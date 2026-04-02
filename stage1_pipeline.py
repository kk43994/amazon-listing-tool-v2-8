"""
第一阶段：AI商品内容处理 Pipeline V2
读取Excel → AI图片处理 + 文案生成(含前后对比) → 输出对比Excel
"""
import os
import json
import logging
import base64
import time
import re
import threading
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from contextlib import contextmanager
from typing import Callable, Dict, List, Optional
from io import BytesIO
from urllib.parse import quote
from uuid import uuid4

from PIL import Image

from amazon.mapper import MARKETPLACE_CODE_BY_ID, MARKETPLACE_LANGUAGE, MARKETPLACE_IDS, SEARCH_TERM_BYTE_LIMIT
from config import get_config
from core.excel.processor import ExcelProcessor
from core.ai_client import ai_text, ai_image_edit_url
from core.media_store import get_media_store
from core.search_term_utils import count_search_term_bytes, dedup_search_terms, truncate_search_terms
from core.utils import filter_rows
from core.prompts.amazon_prompts import (
    TITLE_PROMPT, BULLET_POINTS_PROMPT, DESCRIPTION_PROMPT,
    SEARCH_TERMS_PROMPT, SPECIAL_FEATURE_PROMPT, TARGET_AUDIENCE_PROMPT,
    SUBJECT_KEYWORDS_PROMPT, IMAGE_BG_GRADIENT_PROMPT, IMAGE_BG_LIFESTYLE_PROMPT,
    IMAGE_BG_WHITE_PROMPT, PRODUCT_TYPE_SUGGEST_PROMPT
)

logger = logging.getLogger(__name__)


class Stage1Pipeline:
    """第一阶段：AI商品内容处理(含前后对比)"""

    def __init__(self):
        self.config = get_config()
        self.excel = ExcelProcessor()
        self.media_store = get_media_store()
        self.stats = {'total': 0, 'success': 0, 'failed': 0, 'skipped': 0}

    def run(self, input_file: str, output_file: str = None,
            process_images: bool = True, process_text: bool = True,
            rows: str = None, resume: bool = False,
            text_fields: List[str] = None, overwrite_existing: bool = True,
            progress_callback: Optional[Callable[[int, int, str], None]] = None,
            image_scope: str = 'main', image_style: str = 'white',
            image_custom_prompt: str = '',
            image_reference_url: str = '',
            cancel_event: Optional[threading.Event] = None):
        """
        运行第一阶段处理

        Args:
            input_file: 输入Excel路径
            output_file: 输出Excel路径
            process_images: 是否处理图片
            process_text: 是否生成文案
            rows: 行范围 (如 "1-10")
            resume: 断点续传
            text_fields: 指定生成哪些文案字段
            overwrite_existing: 是否覆盖已有AI结果
        """
        self.stats = {'total': 0, 'success': 0, 'failed': 0, 'skipped': 0}
        image_scope = self._normalize_image_scope(image_scope)
        image_style = self._normalize_bg_style(image_style)
        cancel_event = cancel_event or threading.Event()

        if not output_file:
            timestamp = time.strftime('%Y%m%d_%H%M%S')
            output_file = os.path.join(self.config.OUTPUT_DIR,
                                       f'对比结果_{timestamp}.xlsx')

        logger.info("🚀 第一阶段启动 (V2 前后对比模式)")
        logger.info(f"  输入: {input_file}")
        logger.info(f"  输出: {output_file}")
        logger.info(f"  图片处理: {'✅' if process_images else '❌'}")
        logger.info(f"  文案生成: {'✅' if process_text else '❌'}")
        if process_images:
            logger.info(f"  图片范围: {'全部图片' if image_scope == 'all' else '主图'}")
            logger.info(f"  背景风格: {image_style}")

        # 1. 读取Excel
        data = self.excel.read_input(input_file)
        col_map = self.excel.detect_columns()

        # 行范围过滤
        if rows:
            data = filter_rows(data, rows)
            logger.info(f"📏 行范围过滤: {rows} → {len(data)} 条")

        # 断点续传
        progress_file = os.path.join(self.config.OUTPUT_DIR, '.progress.json')
        start_idx = 0
        if resume and os.path.exists(progress_file):
            try:
                with open(progress_file, 'r', encoding='utf-8') as f:
                    progress = json.load(f)

                saved_input = os.path.abspath(progress.get('input', ''))
                current_input = os.path.abspath(input_file)
                saved_mtime = progress.get('input_mtime')
                current_mtime = os.path.getmtime(input_file)
                same_file = saved_input == current_input
                same_mtime = saved_mtime == current_mtime
                same_rows = progress.get('rows') == rows

                if same_file and same_mtime and same_rows:
                    start_idx = progress.get('last_completed', 0)
                    logger.info(f"🔄 断点续传: 从第 {start_idx + 1} 条继续")
                else:
                    logger.warning("⚠️ 发现旧的断点文件，但与当前输入不一致，已忽略 resume")
            except Exception as e:
                logger.warning(f"⚠️ 读取断点文件失败，已忽略 resume: {e}")

        self.stats['total'] = len(data)
        self.stats['skipped'] = start_idx
        logger.info(f"📊 待处理: {len(data)} 条商品")
        pending_entries = [(idx, item) for idx, item in enumerate(data) if idx >= start_idx]
        cancelled = False
        if not pending_entries:
            logger.info("⏭️ 没有新的商品需要处理")
            if progress_callback:
                progress_callback(len(data), len(data), '无待处理商品')
        else:
            text_limit, image_limit, worker_count = self._resolve_worker_counts(
                process_text=process_text,
                process_images=process_images,
                pending_count=len(pending_entries),
            )
            logger.info(
                f"⚙️ 并发配置: 文案 {text_limit} / 图片 {image_limit} / 工作线程 {worker_count}"
            )

            text_semaphore = threading.BoundedSemaphore(text_limit) if process_text else None
            image_semaphore = threading.BoundedSemaphore(image_limit) if process_images else None
            progress_lock = threading.Lock()
            completed_indices = set(range(start_idx))
            contiguous_completed = start_idx
            pending_iter = iter(pending_entries)

            self._save_progress(progress_file, input_file, rows, contiguous_completed)

            with ThreadPoolExecutor(max_workers=worker_count, thread_name_prefix='stage1') as executor:
                future_map = {}
                while len(future_map) < worker_count and not cancel_event.is_set():
                    submitted = self._submit_next_future(
                        executor=executor,
                        future_map=future_map,
                        pending_iter=pending_iter,
                        total=len(data),
                        col_map=col_map,
                        process_text=process_text,
                        process_images=process_images,
                        text_fields=text_fields,
                        overwrite_existing=overwrite_existing,
                        text_semaphore=text_semaphore,
                        image_semaphore=image_semaphore,
                        image_scope=image_scope,
                        image_style=image_style,
                        image_custom_prompt=image_custom_prompt,
                        image_reference_url=image_reference_url,
                        cancel_event=cancel_event,
                    )
                    if not submitted:
                        break

                while future_map:
                    done, _ = wait(list(future_map.keys()), return_when=FIRST_COMPLETED)
                    for future in done:
                        future_map.pop(future, None)
                        result = future.result()
                        with progress_lock:
                            if result.get('cancelled'):
                                self.stats['skipped'] += 1
                            elif result.get('ok'):
                                self.stats['success'] += 1
                            else:
                                self.stats['failed'] += 1

                            completed_indices.add(result['idx'])
                            while contiguous_completed in completed_indices:
                                contiguous_completed += 1

                            self._save_progress(progress_file, input_file, rows, contiguous_completed)
                            progress_message = (
                                f"已完成 {contiguous_completed}/{len(data)}"
                                f" (成功:{self.stats['success']} 失败:{self.stats['failed']})"
                            )
                            if result.get('cancelled'):
                                progress_message += " [已取消]"
                            logger.info(f"  进度: {progress_message}")
                            if progress_callback:
                                progress_callback(
                                    contiguous_completed,
                                    len(data),
                                    result.get('current_item', ''),
                                )

                    if cancel_event.is_set():
                        cancelled = True
                        for future in list(future_map.keys()):
                            future.cancel()
                        break

                    while len(future_map) < worker_count and not cancel_event.is_set():
                        submitted = self._submit_next_future(
                            executor=executor,
                            future_map=future_map,
                            pending_iter=pending_iter,
                            total=len(data),
                            col_map=col_map,
                            process_text=process_text,
                            process_images=process_images,
                            text_fields=text_fields,
                            overwrite_existing=overwrite_existing,
                            text_semaphore=text_semaphore,
                            image_semaphore=image_semaphore,
                            image_scope=image_scope,
                            image_style=image_style,
                            image_custom_prompt=image_custom_prompt,
                            image_reference_url=image_reference_url,
                            cancel_event=cancel_event,
                        )
                        if not submitted:
                            break

                cancelled = cancelled or cancel_event.is_set()

        # 3. 输出前后对比Excel
        self.excel.write_comparison_output(data, output_file, col_map)

        # 4. 统计
        logger.info(f"\n{'='*50}")
        logger.info("⏹️ 第一阶段已取消，已输出当前结果" if cancelled else "🎉 第一阶段完成!")
        logger.info(f"  总计: {self.stats['total']}")
        logger.info(f"  成功: {self.stats['success']}")
        logger.info(f"  失败: {self.stats['failed']}")
        logger.info(f"  输出: {output_file}")

        if progress_callback:
            progress_callback(
                contiguous_completed if 'contiguous_completed' in locals() else len(data),
                len(data),
                '已取消' if cancelled else '完成',
            )

        return output_file

    def _resolve_worker_counts(self, process_text: bool, process_images: bool, pending_count: int):
        text_limit = max(1, int(getattr(self.config, 'AI_CONCURRENCY', 3) or 3))
        image_limit = max(1, int(getattr(self.config, 'IMAGE_CONCURRENCY', 2) or 2))
        slots = 0
        if process_text:
            slots += text_limit
        if process_images:
            slots += image_limit
        worker_count = max(1, min(pending_count or 1, slots or 1))
        return text_limit, image_limit, worker_count

    def _submit_next_future(self, executor, future_map: Dict, pending_iter, total: int, col_map: Dict,
                            process_text: bool, process_images: bool,
                            text_fields: Optional[List[str]], overwrite_existing: bool,
                            text_semaphore, image_semaphore,
                            image_scope: str, image_style: str,
                            image_custom_prompt: str = '',
                            image_reference_url: str = '',
                            cancel_event: threading.Event = None) -> bool:
        try:
            idx, item = next(pending_iter)
        except StopIteration:
            return False

        future = executor.submit(
            self._process_single_item,
            idx=idx,
            total=total,
            item=item,
            col_map=col_map,
            process_text=process_text,
            process_images=process_images,
            text_fields=text_fields,
            overwrite_existing=overwrite_existing,
            text_semaphore=text_semaphore,
            image_semaphore=image_semaphore,
            image_scope=image_scope,
            image_style=image_style,
            image_custom_prompt=image_custom_prompt,
            image_reference_url=image_reference_url,
            cancel_event=cancel_event,
        )
        future_map[future] = idx
        return True

    def _normalize_image_scope(self, scope: str) -> str:
        return 'all' if str(scope or '').strip().lower() == 'all' else 'main'

    def _normalize_bg_style(self, style: str) -> str:
        normalized = str(style or '').strip().lower()
        aliases = {
            'white': 'white',
            'scene': 'scene',
            'lifestyle': 'scene',
            'gradient': 'gradient',
        }
        return aliases.get(normalized, 'white')

    def _save_progress(self, progress_file: str, input_file: str, rows: Optional[str], last_completed: int):
        os.makedirs(self.config.OUTPUT_DIR, exist_ok=True)
        with open(progress_file, 'w', encoding='utf-8') as f:
            json.dump({
                'last_completed': last_completed,
                'input': os.path.abspath(input_file),
                'input_mtime': os.path.getmtime(input_file),
                'rows': rows,
            }, f)

    def _process_single_item(self, idx: int, total: int, item: Dict, col_map: Dict,
                             process_text: bool, process_images: bool,
                             text_fields: Optional[List[str]], overwrite_existing: bool,
                             text_semaphore, image_semaphore,
                             image_scope: str, image_style: str,
                             image_custom_prompt: str = '',
                             image_reference_url: str = '',
                             cancel_event: Optional[threading.Event] = None) -> Dict:
        sku = str(item.get(col_map.get('sku', ''), '') or item.get('SKU', '') or f'ROW-{idx+1}')
        logger.info(f"\n{'='*50}")
        logger.info(f"📦 处理第 {idx+1}/{total} 条商品 (SKU={sku})")
        cancel_event = cancel_event or threading.Event()

        try:
            generated_any = False

            if cancel_event.is_set():
                return {
                    'idx': idx,
                    'sku': sku,
                    'ok': False,
                    'cancelled': True,
                    'current_item': f'{sku} 已取消',
                }

            if process_text:
                product_info = self._build_product_info(item, col_map)
                with self._acquire_limiter(text_semaphore):
                    if cancel_event.is_set():
                        return {
                            'idx': idx,
                            'sku': sku,
                            'ok': False,
                            'cancelled': True,
                            'current_item': f'{sku} 已取消',
                        }
                    product_type = self._detect_product_type(item, col_map, product_info)
                    logger.info(f"  产品类型: {product_type}")
                    try:
                        generated_any = self._generate_text_v2(
                            item, product_info, product_type, col_map,
                            text_fields=text_fields, overwrite_existing=overwrite_existing,
                            cancel_event=cancel_event,
                        )
                    except TypeError as e:
                        if any(token in str(e) for token in ('text_fields', 'overwrite_existing', 'cancel_event')):
                            generated_any = self._generate_text_v2(item, product_info, product_type, col_map)
                        else:
                            raise

            if cancel_event.is_set():
                return {
                    'idx': idx,
                    'sku': sku,
                    'ok': False,
                    'cancelled': True,
                    'current_item': f'{sku} 已取消',
                }

            if process_images:
                with self._acquire_limiter(image_semaphore):
                    generated_any = self._process_images(
                        item,
                        col_map,
                        idx,
                        scope=image_scope,
                        bg_style=image_style,
                        custom_prompt=image_custom_prompt,
                        reference_image_url=image_reference_url,
                        cancel_event=cancel_event,
                    ) or generated_any

            item['submit_status'] = 'PENDING'
            if process_text or process_images:
                item['AI状态'] = 'completed' if (generated_any or self._has_existing_ai_output(item)) else 'failed'

            return {
                'idx': idx,
                'sku': sku,
                'ok': True,
                'current_item': f'{sku} 已完成',
            }

        except Exception as e:
            logger.error(f"  ❌ 处理失败: {e}")
            import traceback
            traceback.print_exc()
            item['submit_status'] = 'AI_ERROR'
            item['AI状态'] = 'failed'
            item['issues'] = str(e)
            return {
                'idx': idx,
                'sku': sku,
                'ok': False,
                'current_item': f'{sku} 失败',
            }

    @contextmanager
    def _acquire_limiter(self, limiter):
        if limiter is None:
            yield
            return
        limiter.acquire()
        try:
            yield
        finally:
            limiter.release()

    def _build_product_info(self, item: Dict, col_map: Dict) -> str:
        """从已有数据构建产品信息文本"""
        parts = []
        # 优先级排列关键字段
        priority_fields = ['title', 'brand', 'description', 'bullet_points',
                          'price', 'color', 'size', 'material', 'product_type']

        for field in priority_fields:
            col_name = col_map.get(field)
            if col_name and item.get(col_name):
                parts.append(f"{field}: {item[col_name]}")

        # 其余字段
        for field, col_name in col_map.items():
            if field not in priority_fields and col_name and item.get(col_name):
                val = str(item[col_name])
                if len(val) > 5:  # 跳过太短的值
                    parts.append(f"{field}: {val}")

        return "\n".join(parts) if parts else "No product information available"

    def _detect_product_type(self, item: Dict, col_map: Dict, product_info: str) -> str:
        """检测或推断产品类型"""
        # 1. 直接从Excel取
        pt_col = col_map.get('product_type')
        if pt_col and item.get(pt_col):
            return str(item[pt_col]).strip()

        # 2. AI推断
        try:
            pt = ai_text(
                PRODUCT_TYPE_SUGGEST_PROMPT.format(product_info=product_info),
                temperature=0.3, max_tokens=50,
            )
            logger.info(f"  🤖 AI推荐产品类型: {pt}")
            item['product_type'] = pt
            return pt
        except Exception as e:
            logger.warning(f"  产品类型推荐失败: {e}")
            return "PRODUCT"

    def _generate_text_v2(self, item: Dict, product_info: str,
                          product_type: str, col_map: Dict,
                          text_fields: List[str] = None,
                          overwrite_existing: bool = True,
                          cancel_event: Optional[threading.Event] = None):
        """
        AI生成文案 V2 — 分条生成，严格长度控制
        结果写入item的AI字段，保留原始值用于对比
        """
        cancel_event = cancel_event or threading.Event()
        selected_fields = set(text_fields or ['title', 'bullets', 'description', 'keywords'])
        generated_any = False
        text_attempts = 0
        text_errors = []
        marketplace_language = self._resolve_marketplace_language()
        text_generated_at = ''

        # === 标题 ===
        title = item.get('AI标题', '')
        if 'title' in selected_fields and (overwrite_existing or not title):
            if cancel_event.is_set():
                return generated_any
            logger.info("  📝 生成标题...")
            title_result = self._ai_text_with_retry(
                TITLE_PROMPT.format(
                    product_info=product_info,
                    product_type=product_type,
                    language=marketplace_language,
                )
            )
            text_attempts += title_result['attempts']
            if title_result['error']:
                text_errors.append(f"标题: {title_result['error']}")
            title = title_result['text']
            if len(title) > 200:
                title = title[:197] + "..."
                logger.warning("  ⚠️ 标题超长，已截断至200字符")
            # Amazon 合规校验：禁用字符 + 同词重复
            from core.title_validation import validate_title, fix_title
            brand = item.get(col_map.get('brand', ''), '') or item.get('brand', '')
            title_check = validate_title(title, brand=brand)
            if not title_check['valid']:
                title, fix_changes = fix_title(title, brand=brand)
                for change in fix_changes:
                    logger.warning(f"  ⚠️ 标题自动修正: {change}")
                if len(title) > 200:
                    title = title[:197] + "..."
            item['AI标题'] = title
            logger.info(f"     → [{len(title)}字符] {title[:80]}...")
            generated_any = generated_any or bool(str(title).strip())
            if title:
                text_generated_at = self._now_text()

        # === Bullet Points(分5条) ===
        bullets = [item.get(f'AI卖点{i}', '') for i in range(1, 6)]
        if 'bullets' in selected_fields and (overwrite_existing or not any(bullets)):
            if cancel_event.is_set():
                return generated_any
            logger.info("  📝 生成Bullet Points...")
            bullets_result = self._ai_text_with_retry(
                BULLET_POINTS_PROMPT.format(
                    product_info=product_info,
                    product_type=product_type,
                    language=marketplace_language,
                )
            )
            text_attempts += bullets_result['attempts']
            if bullets_result['error']:
                text_errors.append(f"卖点: {bullets_result['error']}")
            bullets_raw = bullets_result['text']
            bullets = self._parse_bullets(bullets_raw)
            for i, bp in enumerate(bullets, 1):
                if len(bp) > 500:
                    bp = bp[:497] + "..."
                    logger.warning(f"  ⚠️ 卖点{i}超长，已截断至500字符")
                item[f'AI卖点{i}'] = bp
                logger.info(f"     卖点{i} [{len(bp)}字符]: {bp[:60]}...")
                generated_any = generated_any or bool(str(bp).strip())
            item['AI五点描述'] = "\n".join(f"• {bp}" for bp in bullets)
            if any(str(bp).strip() for bp in bullets):
                text_generated_at = self._now_text()

        # === 描述 ===
        desc = item.get('AI商品描述', '')
        if 'description' in selected_fields and (overwrite_existing or not desc):
            if cancel_event.is_set():
                return generated_any
            logger.info("  📝 生成商品描述...")
            desc_result = self._ai_text_with_retry(
                DESCRIPTION_PROMPT.format(
                    product_info=product_info,
                    product_type=product_type,
                    language=marketplace_language,
                )
            )
            text_attempts += desc_result['attempts']
            if desc_result['error']:
                text_errors.append(f"描述: {desc_result['error']}")
            desc = desc_result['text']
            if len(desc) > 2000:
                desc = desc[:1997] + "..."
                logger.warning("  ⚠️ 描述超长，已截断至2000字符")
            item['AI商品描述'] = desc
            logger.info(f"     [{len(desc)}字符]")
            generated_any = generated_any or bool(str(desc).strip())
            if desc:
                text_generated_at = self._now_text()

        # === 搜索关键词 ===
        keywords = item.get('AI搜索关键词', '')
        if 'keywords' in selected_fields and (overwrite_existing or not keywords):
            if cancel_event.is_set():
                return generated_any
            logger.info("  📝 生成搜索关键词...")
            existing_title = title or item.get(col_map.get('title', ''), '')
            marketplace_id = str(getattr(self.config, 'AMAZON_MARKETPLACE', '') or '').strip()
            kw_byte_limit = SEARCH_TERM_BYTE_LIMIT.get(marketplace_id, 250)
            keywords_result = self._ai_text_with_retry(
                SEARCH_TERMS_PROMPT.format(
                    product_info=product_info,
                    product_type=product_type,
                    title=existing_title,
                    language=marketplace_language,
                    byte_limit=kw_byte_limit,
                )
            )
            text_attempts += keywords_result['attempts']
            if keywords_result['error']:
                text_errors.append(f"搜索词: {keywords_result['error']}")
            keywords = keywords_result['text']

            # 去重：移除标题和五点中已有的词
            existing_bullets = [item.get(f'AI卖点{i}', '') for i in range(1, 6)]
            keywords = dedup_search_terms(keywords, existing_title, existing_bullets)

            # 按 Amazon 规则截断（仅统计单词字节，空格和标点不计入）
            kw_bytes = count_search_term_bytes(keywords)
            if kw_bytes > kw_byte_limit:
                keywords = truncate_search_terms(keywords, kw_byte_limit)
                logger.warning(f"  ⚠️ 搜索词超过{kw_byte_limit}字节限制，已截断")
            item['AI搜索关键词'] = keywords
            kw_bytes = count_search_term_bytes(keywords)
            logger.info(f"     [{kw_bytes}/{kw_byte_limit}字节] {keywords[:60]}...")

            generated_any = generated_any or bool(str(keywords).strip())
            if keywords:
                text_generated_at = self._now_text()

        # === 特殊功能/亮点 ===
        special_feature = item.get('AI特殊功能', '')
        if 'special_feature' in selected_fields and (overwrite_existing or not special_feature):
            if not cancel_event.is_set():
                logger.info("  📝 生成特殊功能亮点...")
                sf_result = self._ai_text_with_retry(
                    SPECIAL_FEATURE_PROMPT.format(
                        product_info=product_info,
                        product_type=product_type,
                        language=marketplace_language,
                    )
                )
                text_attempts += sf_result['attempts']
                if sf_result['error']:
                    text_errors.append(f"特殊功能: {sf_result['error']}")
                special_feature = sf_result['text'].strip()
                item['AI特殊功能'] = special_feature
                if special_feature:
                    logger.info(f"     → {special_feature[:80]}...")
                    generated_any = True
                    text_generated_at = self._now_text()

        # === 目标受众 ===
        target_audience = item.get('AI目标受众', '')
        if 'target_audience' in selected_fields and (overwrite_existing or not target_audience):
            if not cancel_event.is_set():
                logger.info("  📝 生成目标受众关键词...")
                ta_result = self._ai_text_with_retry(
                    TARGET_AUDIENCE_PROMPT.format(
                        product_info=product_info,
                        product_type=product_type,
                        language=marketplace_language,
                    )
                )
                text_attempts += ta_result['attempts']
                if ta_result['error']:
                    text_errors.append(f"目标受众: {ta_result['error']}")
                target_audience = ta_result['text'].strip()
                item['AI目标受众'] = target_audience
                if target_audience:
                    logger.info(f"     → {target_audience[:80]}...")
                    generated_any = True
                    text_generated_at = self._now_text()

        # === 主题关键词 ===
        subject_kw = item.get('AI主题关键词', '')
        if 'subject_keywords' in selected_fields and (overwrite_existing or not subject_kw):
            if not cancel_event.is_set():
                logger.info("  📝 生成主题关键词...")
                existing_title = title or item.get(col_map.get('title', ''), '')
                sk_result = self._ai_text_with_retry(
                    SUBJECT_KEYWORDS_PROMPT.format(
                        product_info=product_info,
                        product_type=product_type,
                        title=existing_title,
                        language=marketplace_language,
                    )
                )
                text_attempts += sk_result['attempts']
                if sk_result['error']:
                    text_errors.append(f"主题关键词: {sk_result['error']}")
                subject_kw = sk_result['text'].strip()
                item['AI主题关键词'] = subject_kw
                if subject_kw:
                    logger.info(f"     → {subject_kw[:80]}...")
                    generated_any = True
                    text_generated_at = self._now_text()

        if text_attempts:
            self._apply_ai_trace(
                item,
                kind='text',
                attempts=text_attempts,
                success=generated_any,
                error=' | '.join(text_errors),
                generated_at=text_generated_at,
            )

        return generated_any

    def _parse_bullets(self, text: str) -> List[str]:
        """解析Bullet Points为5条独立文本 — V2增强版"""
        import re

        # 预处理：移除markdown粗体/斜体
        text = re.sub(r'\*\*(.+?)\*\*', r'\1', text)
        text = re.sub(r'\*(.+?)\*', r'\1', text)

        lines = text.strip().split('\n')
        bullets = []
        current_bullet = ""

        for line in lines:
            line = line.strip()
            if not line:
                continue

            # 检测是否是新的bullet point开头
            is_new_bullet = False

            # 模式1: 数字前缀 (1. / 1) / 1: / 1、)
            num_match = re.match(r'^(\d+)\s*[.):\-、]\s*(.+)', line)
            if num_match:
                is_new_bullet = True
                line = num_match.group(2)

            # 模式2: 符号前缀 (• / - / * / ▪ / → / ✅ / 🔹)
            elif re.match(r'^[•\-\*▪→✅🔹►◆■□●○]\s*', line):
                is_new_bullet = True
                line = re.sub(r'^[•\-\*▪→✅🔹►◆■□●○]\s*', '', line)

            # 模式3: Bullet Point N: / BP N: 格式
            elif re.match(r'^(?:Bullet\s*Point|BP)\s*\d+\s*[:：]\s*', line, re.IGNORECASE):
                is_new_bullet = True
                line = re.sub(r'^(?:Bullet\s*Point|BP)\s*\d+\s*[:：]\s*', '', line, flags=re.IGNORECASE)

            # 模式4: 全大写标题行后跟描述（如 "PREMIUM QUALITY - ..."）
            elif re.match(r'^[A-Z][A-Z\s&]+[\-–—:]\s*', line):
                is_new_bullet = True

            if is_new_bullet:
                # 保存之前的bullet
                if current_bullet:
                    bullets.append(current_bullet.strip())
                current_bullet = line.strip()
            else:
                # 续行 — 追加到当前bullet
                if current_bullet:
                    current_bullet += " " + line
                else:
                    current_bullet = line

        # 保存最后一个
        if current_bullet:
            bullets.append(current_bullet.strip())

        # 移除空白项
        bullets = [b for b in bullets if b.strip()]

        # 如果解析出的太少（<3），可能AI返回了段落格式
        # 尝试按句号分割
        if len(bullets) < 3 and len(text) > 200:
            logger.warning(f"  ⚠️ Bullet解析只得到{len(bullets)}条，尝试按段落/句号分割")
            # 先试双换行分割
            paragraphs = re.split(r'\n\n+', text.strip())
            if len(paragraphs) >= 3:
                bullets = [p.strip() for p in paragraphs if p.strip()]
            else:
                # 按句号分割（保留句号后的内容完整性）
                sentences = re.split(r'(?<=[.!])\s+(?=[A-Z])', text.strip())
                if len(sentences) >= 3:
                    bullets = [s.strip() for s in sentences if s.strip()]

        # 确保恰好5条
        while len(bullets) < 5:
            bullets.append("")

        result = bullets[:5]
        logger.info(f"  📋 解析到 {len([b for b in result if b])} 条有效Bullet Points")
        return result

    def _process_images(self, item: Dict, col_map: Dict, idx: int,
                        scope: str = 'main', bg_style: str = 'white',
                        custom_prompt: str = '',
                        reference_image_url: str = '',
                        cancel_event: Optional[threading.Event] = None):
        """AI图片处理，支持主图或全图生成。支持自定义提示词。"""
        cancel_event = cancel_event or threading.Event()
        scope = self._normalize_image_scope(scope)
        bg_style = self._normalize_bg_style(bg_style)
        image_sources = self._collect_image_sources(item, col_map, scope)

        if not image_sources:
            logger.info("  ⏭️ 无图片URL，跳过图片处理")
            return False

        generated_any = False
        # 自定义提示词优先，否则使用预设风格
        prompt = custom_prompt.strip() if custom_prompt and custom_prompt.strip() else self._build_image_prompt(bg_style, item, col_map)
        if custom_prompt and custom_prompt.strip():
            logger.info(f"  🖼️ 使用自定义图片提示词: {prompt[:60]}...")
        image_attempts = 0
        image_errors = []
        image_generated_at = ''

        for image_source in image_sources:
            if cancel_event.is_set():
                logger.info("  ⏹️ 检测到取消信号，停止图片处理")
                break

            slot = image_source['slot']
            image_url = image_source['url']
            label = '主图' if slot == 'main' else f'副图{slot}'

            try:
                output_path = self._build_image_output_path(item, col_map, idx, slot=slot)
                os.makedirs(os.path.dirname(output_path), exist_ok=True)

                logger.info(f"  🖼️ {label}URL: {image_url[:60]}...")
                logger.info(f"  🖼️ 调用图片编辑API生成{label} ({bg_style})...")

                result_info = self._ai_image_edit_url_with_retry(image_url, prompt, reference_image_url=reference_image_url)
                image_attempts += result_info['attempts']
                if result_info['error']:
                    image_errors.append(f"{label}: {result_info['error']}")
                result = result_info['data']
                if not result:
                    logger.warning(f"  ⚠️ {label}处理失败")
                    continue

                self._save_image(result, output_path)

                # Amazon 图片合规校验
                from core.image_validation import validate_image, ensure_jpeg
                try:
                    with Image.open(output_path) as saved_img:
                        is_main = (slot == 'main')
                        img_check = validate_image(saved_img, is_main=is_main)
                        for issue in img_check['issues']:
                            icon = '❌' if issue['level'] == 'error' else '⚠️'
                            logger.warning(f"     {icon} {label}: {issue['message']}")
                            image_errors.append(f"{label}: {issue['message']}")
                        if img_check['format_hint']:
                            logger.info(f"     ℹ️ {label}: {img_check['format_hint']}")
                        # 确保保存为 JPEG
                        rgb_img = ensure_jpeg(saved_img)
                        rgb_img.save(output_path, format='JPEG', quality=95)
                except Exception as img_err:
                    logger.warning(f"     ⚠️ {label}合规校验失败: {img_err}")

                upload_result = self._upload_generated_image(item, col_map, slot, output_path)
                self._store_generated_image(item, slot, output_path, upload_result)
                logger.info(f"     → {label}保存: {output_path}")
                if upload_result.get('success'):
                    logger.info(f"     → {label}上传: {upload_result.get('locator')}")
                elif upload_result.get('error'):
                    logger.warning(f"     → {label}上传未完成: {upload_result.get('error')}")
                generated_any = True
                image_generated_at = self._now_text()
            except Exception as e:
                logger.error(f"  ❌ {label}处理失败: {e}")
                image_errors.append(f"{label}: {e}")

        if image_attempts:
            self._apply_ai_trace(
                item,
                kind='image',
                attempts=image_attempts,
                success=generated_any,
                error=' | '.join(image_errors),
                generated_at=image_generated_at,
            )

        return generated_any

    def _has_existing_ai_output(self, item: Dict) -> bool:
        fields = [
            'AI标题', 'AI商品描述', 'AI搜索关键词',
            'AI主图路径', 'AI主图URL', 'AI主图预览URL',
            'AI主图上传状态', 'AI主图上传错误',
        ]
        fields.extend([f'AI卖点{i}' for i in range(1, 6)])
        fields.extend([f'AI副图{i}路径' for i in range(2, 10)])
        fields.extend([f'AI副图{i}URL' for i in range(2, 10)])
        fields.extend([f'AI副图{i}预览URL' for i in range(2, 10)])
        fields.extend([f'AI副图{i}上传状态' for i in range(2, 10)])
        fields.extend([f'AI副图{i}上传错误' for i in range(2, 10)])
        return any(str(item.get(field, '') or '').strip() for field in fields)

    def _build_image_output_path(self, item: Dict, col_map: Dict, idx: int, slot='main') -> str:
        """为生成图片构建稳定且不冲突的输出路径。"""
        sku_col = col_map.get('sku')
        raw_sku = item.get(sku_col) if sku_col else item.get('SKU')
        row_index = item.get('_row_index', idx + 1)

        sku_part = re.sub(r'[^0-9A-Za-z_-]+', '_', str(raw_sku or '').strip()).strip('_')
        if not sku_part:
            sku_part = f'row_{row_index}'

        unique_suffix = uuid4().hex[:8]
        slot_prefix = 'main' if slot == 'main' else f'sub_{slot}'
        filename = f"{slot_prefix}_{sku_part}_{row_index}_{unique_suffix}.jpg"
        return os.path.join(self.config.OUTPUT_DIR, 'images', filename)

    def _collect_image_sources(self, item: Dict, col_map: Dict, scope: str) -> List[Dict]:
        sources: List[Dict] = []
        primary_field = ''

        def add_source(slot, field_names):
            for field_name in field_names:
                col_name = col_map.get(field_name)
                value = item.get(col_name) if col_name else item.get(field_name)
                text = str(value or '').strip()
                if text.startswith('http'):
                    sources.append({'slot': slot, 'url': text, 'field': field_name})
                    return field_name
            return ''

        primary_field = add_source('main', ['image_url', 'main_image_url', 'image_2', 'image_3'])
        if scope == 'all':
            for slot in range(2, 10):
                candidate_fields = [f'image_{slot}', f'other_image_url_{slot-1}']
                if primary_field and primary_field in candidate_fields:
                    continue
                add_source(slot, candidate_fields)

        unique_sources = []
        seen = set()
        for source in sources:
            key = (source['slot'], source['url'])
            if key in seen:
                continue
            seen.add(key)
            unique_sources.append(source)
        return unique_sources

    def _build_image_prompt(self, bg_style: str, item: Dict, col_map: Dict) -> str:
        bg_style = self._normalize_bg_style(bg_style)
        if bg_style == 'scene':
            product_type = str(
                item.get(col_map.get('product_type', ''), '')
                or item.get('product_type', '')
                or item.get(col_map.get('title', ''), '')
                or item.get('item_name', '')
                or 'the product category'
            ).strip()
            return IMAGE_BG_LIFESTYLE_PROMPT.format(style_hint=product_type or 'the product category')
        if bg_style == 'gradient':
            return IMAGE_BG_GRADIENT_PROMPT
        return IMAGE_BG_WHITE_PROMPT

    def _upload_generated_image(self, item: Dict, col_map: Dict, slot, output_path: str) -> Dict:
        sku = self._resolve_sku(item, col_map)
        marketplace_code = self._resolve_marketplace_code()
        if getattr(self.media_store, 'enabled', lambda: False)():
            result = self.media_store.upload_image(
                output_path,
                sku=sku,
                slot='main' if slot == 'main' else f'sub{slot}',
                marketplace=marketplace_code,
            )
            return result.to_dict()

        public_url = self._build_public_image_url(output_path)
        if public_url:
            return {
                'success': True,
                'locator': public_url,
                'preview_url': '',
                'provider': 'legacy_public_base',
                'bucket': '',
                'key': '',
                'etag': '',
                'error': '',
            }

        return {
            'success': False,
            'locator': '',
            'preview_url': '',
            'provider': 'disabled',
            'bucket': '',
            'key': '',
            'etag': '',
            'error': '媒体存储未启用，且未配置 OUTPUT_IMAGE_PUBLIC_BASE',
        }

    def _store_generated_image(self, item: Dict, slot, output_path: str, upload_result: Optional[Dict] = None):
        upload_result = upload_result or {}
        locator = str(upload_result.get('locator', '') or '').strip()
        preview_url = str(upload_result.get('preview_url', '') or '').strip()
        status = 'uploaded' if upload_result.get('success') else 'failed'
        error = str(upload_result.get('error', '') or '').strip()
        if upload_result.get('provider') == 'legacy_public_base' and locator:
            status = 'ready'

        if slot == 'main':
            item['AI主图路径'] = output_path
            item['AI主图URL'] = locator
            item['AI主图预览URL'] = preview_url
            item['AI主图上传状态'] = status
            item['AI主图上传错误'] = error
            return

        item[f'AI副图{slot}路径'] = output_path
        item[f'AI副图{slot}URL'] = locator
        item[f'AI副图{slot}预览URL'] = preview_url
        item[f'AI副图{slot}上传状态'] = status
        item[f'AI副图{slot}上传错误'] = error

    def _resolve_sku(self, item: Dict, col_map: Dict) -> str:
        sku_col = col_map.get('sku')
        raw_sku = item.get(sku_col) if sku_col else item.get('SKU')
        return str(raw_sku or '').strip() or 'unknown'

    def _resolve_marketplace_code(self) -> str:
        marketplace_id = str(getattr(self.config, 'AMAZON_MARKETPLACE', '') or '').strip()
        return MARKETPLACE_CODE_BY_ID.get(marketplace_id, 'US')

    def _resolve_marketplace_language(self) -> str:
        marketplace_id = str(getattr(self.config, 'AMAZON_MARKETPLACE', '') or '').strip()
        return MARKETPLACE_LANGUAGE.get(marketplace_id, 'en_US')

    def _ai_text(self, prompt: str) -> str:
        """调用AI生成文本"""
        return ai_text(prompt, temperature=0.7, max_tokens=2000)

    def _ai_text_with_retry(self, prompt: str) -> Dict:
        attempts_allowed = max(1, int(getattr(self.config, 'OPENAI_MAX_RETRIES', 1) or 1))
        last_error = ''
        for attempt in range(1, attempts_allowed + 1):
            try:
                result = ai_text(
                    prompt,
                    temperature=0.7,
                    max_tokens=2000,
                    raise_on_error=True,
                )
                text = str(result or '').strip()
                if not text:
                    raise ValueError('AI返回空文本')
                return {'text': text, 'attempts': attempt, 'error': ''}
            except Exception as exc:
                last_error = str(exc)
                logger.warning(f"  ⚠️ 文本生成第 {attempt}/{attempts_allowed} 次失败: {last_error}")
                if attempt < attempts_allowed:
                    time.sleep(min(2 ** (attempt - 1), 3))
        return {'text': '', 'attempts': attempts_allowed, 'error': last_error or 'AI返回空文本'}

    def _ai_image_edit_url_with_retry(self, image_url: str, prompt: str,
                                      reference_image_url: str = '') -> Dict:
        attempts_allowed = max(1, int(getattr(self.config, 'OPENAI_MAX_RETRIES', 1) or 1))
        last_error = ''
        for attempt in range(1, attempts_allowed + 1):
            try:
                result = ai_image_edit_url(image_url, prompt,
                                           reference_image_url=reference_image_url,
                                           raise_on_error=True)
                payload = str(result or '').strip()
                if not payload:
                    raise ValueError('AI图片接口未返回图片数据')
                return {'data': payload, 'attempts': attempt, 'error': ''}
            except Exception as exc:
                last_error = str(exc)
                logger.warning(f"  ⚠️ 图片生成第 {attempt}/{attempts_allowed} 次失败: {last_error}")
                if attempt < attempts_allowed:
                    time.sleep(min(2 ** (attempt - 1), 3))
        return {'data': None, 'attempts': attempts_allowed, 'error': last_error or 'AI图片接口未返回图片数据'}

    def _apply_ai_trace(self, item: Dict, kind: str, attempts: int, success: bool, error: str = '', generated_at: str = ''):
        if kind == 'text':
            item['AI文案模型'] = str(getattr(self.config, 'AI_TEXT_MODEL', '') or '').strip()
            item['AI文案协议'] = str(getattr(self.config, 'AI_TEXT_PROTOCOL', '') or '').strip()
            item['AI文案尝试次数'] = attempts
            item['AI文案最后错误'] = '' if success else str(error or '').strip()
            if generated_at:
                item['AI文案生成时间'] = generated_at
        elif kind == 'image':
            item['AI图片模型'] = str(getattr(self.config, 'AI_IMAGE_MODEL', '') or '').strip()
            item['AI图片协议'] = str(getattr(self.config, 'AI_IMAGE_PROTOCOL', '') or '').strip()
            item['AI图片尝试次数'] = attempts
            item['AI图片最后错误'] = '' if success else str(error or '').strip()
            if generated_at:
                item['AI图片生成时间'] = generated_at

    def _now_text(self) -> str:
        return time.strftime('%Y-%m-%d %H:%M:%S')

    def _image_to_base64(self, img: Image.Image) -> str:
        """PIL Image → base64 (缩小到合理尺寸以节省API费用)"""
        buffered = BytesIO()
        if img.mode in ('RGBA', 'LA', 'P'):
            img = img.convert('RGB')
        # 限制到800px以内，既保持质量又不浪费token
        max_size = min(self.config.IMAGE_MAX_SIZE, 800)
        if max(img.size) > max_size:
            ratio = max_size / max(img.size)
            new_size = (int(img.width * ratio), int(img.height * ratio))
            img = img.resize(new_size, Image.LANCZOS)
        img.save(buffered, format="JPEG", quality=85)
        return base64.b64encode(buffered.getvalue()).decode('utf-8')

    def _save_image(self, data: str, path: str):
        """保存图片"""
        try:
            if data.startswith('http'):
                import requests
                resp = requests.get(data, timeout=30)
                with open(path, 'wb') as f:
                    f.write(resp.content)
            else:
                img_data = base64.b64decode(data)
                with open(path, 'wb') as f:
                    f.write(img_data)
        except Exception as e:
            logger.error(f"保存图片失败: {e}")

    def _build_public_image_url(self, local_path: str) -> str:
        """按配置将本地输出图映射为公开URL。"""
        public_base = getattr(self.config, 'OUTPUT_IMAGE_PUBLIC_BASE', '') or ''
        if not public_base:
            return ''

        filename = quote(os.path.basename(local_path))
        if '{filename}' in public_base:
            return public_base.format(filename=filename)
        return f"{public_base.rstrip('/')}/{filename}"
