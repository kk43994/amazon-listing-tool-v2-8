"""
第二阶段：Amazon SP-API 自动化提交 Pipeline
读取处理后的Excel → 映射字段 → 提交到亚马逊 → 轮询结果 → 生成报告
"""
import os
import json
import time
import logging
from datetime import datetime
from typing import Dict, List, Optional

from config import get_config
from core.excel.processor import ExcelProcessor
from amazon.auth import AmazonAuth
from amazon.listings import ListingsAPI
from amazon.feeds import FeedsAPI
from amazon.mapper import FieldMapper

logger = logging.getLogger(__name__)


class Stage2Pipeline:
    """第二阶段：SP-API自动化提交"""

    def __init__(self):
        self.config = get_config()
        self.excel = ExcelProcessor()
        self.stats = {
            'total': 0, 'submitted': 0, 'success': 0,
            'failed': 0, 'skipped': 0,
        }

        # 初始化Amazon API
        self.auth = AmazonAuth(
            client_id=self.config.AMAZON_CLIENT_ID,
            client_secret=self.config.AMAZON_CLIENT_SECRET,
            refresh_token=self.config.AMAZON_REFRESH_TOKEN,
        )
        self.listings = ListingsAPI(
            auth=self.auth,
            seller_id=self.config.AMAZON_SELLER_ID,
            marketplace_id=self.config.AMAZON_MARKETPLACE,
        )
        self.feeds = FeedsAPI(
            auth=self.auth,
            marketplace_id=self.config.AMAZON_MARKETPLACE,
        )
        self.mapper = FieldMapper(self.config.AMAZON_MARKETPLACE)

    def run(self, input_file: str, output_file: str = None,
            mode: str = 'individual', preview: bool = False,
            rows: Optional[str] = None):
        """
        运行第二阶段提交

        Args:
            input_file: 第一阶段输出的Excel路径
            output_file: 提交报告输出路径
            mode: 提交模式 ('individual' 逐条 | 'batch' 批量Feed)
            preview: 预览模式（不实际提交）
            rows: 行范围 (如 "1-10")
        """
        if not output_file:
            timestamp = time.strftime('%Y%m%d_%H%M%S')
            output_file = os.path.join(
                self.config.OUTPUT_DIR, f'提交报告_{timestamp}.xlsx'
            )

        logger.info(f"🚀 第二阶段启动 {'[预览模式]' if preview else ''}")
        logger.info(f"  输入: {input_file}")
        logger.info(f"  模式: {mode}")

        # 检查SP-API凭证
        if not preview:
            self._check_credentials()

        # 1. 读取Excel
        data = self.excel.read_input(input_file)

        # 2. 行范围过滤
        if rows:
            data = self._filter_rows(data, rows)

        self.stats['total'] = len(data)

        # 3. 字段映射
        col_map = self.excel.detect_columns()
        products = [self.mapper.map_excel_row(item, col_map) for item in data]

        if preview:
            # 预览模式: 调用 Amazon VALIDATION_PREVIEW
            logger.info(f"\n{'='*50}")
            logger.info(f"📋 提交预览 — 共 {len(products)} 条商品")
            logger.info(f"{'='*50}")
            valid_count = 0
            for idx, product in enumerate(products):
                sku = product.get('sku', 'N/A')
                title = product.get('title', 'N/A')[:60]
                price = product.get('price', 'N/A')
                validation = self.mapper.validate_required_fields(product)
                errors = list(validation['errors'])
                warnings = list(validation['warnings'])
                preview_status = 'LOCAL_INVALID'

                if validation['valid']:
                    preview_result = self.listings.put_listings_item(sku, product, preview=True)
                    preview_status = str(preview_result.get('status', 'UNKNOWN')).upper()
                    for issue in preview_result.get('issues', []) or []:
                        message = issue.get('message', '')
                        severity = str(issue.get('severity', '')).upper()
                        if severity == 'ERROR':
                            errors.append(message)
                        else:
                            warnings.append(message)

                status = '✅ 通过' if preview_status == 'VALID' and not errors else f"❌ {len(errors)}个错误"
                if preview_status == 'VALID' and not errors:
                    valid_count += 1
                logger.info(f"  {idx+1}. SKU={sku}  标题={title}  价格={price}  校验={status}")
                for err in errors:
                    logger.info(f"       ❌ {err}")
                for warn in warnings:
                    logger.info(f"       ⚠️ {warn}")
                time.sleep(1.0)
            logger.info(f"\n📋 预览完成: {valid_count}/{len(products)} 条可提交")
            return

        # 4. 提交
        results = []
        if mode == 'individual':
            results = self._submit_individual(products)
        elif mode == 'batch':
            results = self._submit_batch(products)

        # 5. 生成报告
        self._generate_report(data, results, output_file)

        # 6. 统计
        logger.info(f"\n{'='*50}")
        logger.info(f"🎉 第二阶段完成!")
        logger.info(f"  总计: {self.stats['total']}")
        logger.info(f"  成功: {self.stats['success']}")
        logger.info(f"  失败: {self.stats['failed']}")
        logger.info(f"  跳过: {self.stats['skipped']}")
        logger.info(f"  报告: {output_file}")

        return output_file

    def _check_credentials(self):
        """检查SP-API凭证是否配置"""
        missing = []
        if not self.config.AMAZON_CLIENT_ID:
            missing.append('AMAZON_CLIENT_ID')
        if not self.config.AMAZON_CLIENT_SECRET:
            missing.append('AMAZON_CLIENT_SECRET')
        if not self.config.AMAZON_REFRESH_TOKEN:
            missing.append('AMAZON_REFRESH_TOKEN')
        if not self.config.AMAZON_SELLER_ID:
            missing.append('AMAZON_SELLER_ID')

        if missing:
            raise ValueError(
                f"❌ 缺少SP-API凭证: {', '.join(missing)}\n"
                f"请在 .env 文件中配置"
            )

    def _filter_rows(self, data: List[Dict], rows: str) -> List[Dict]:
        """根据行范围过滤数据"""
        try:
            if '-' in rows:
                start, end = rows.split('-')
                start, end = int(start) - 1, int(end)
                return data[start:end]
            else:
                idx = int(rows) - 1
                return [data[idx]]
        except (ValueError, IndexError) as e:
            logger.error(f"行范围解析错误: {rows} → {e}")
            return data

    def _submit_individual(self, products: List[Dict]) -> List[Dict]:
        """逐条提交 (使用Listings API)"""
        results = []
        for idx, product in enumerate(products):
            sku = product.get('sku')
            if not sku:
                self.stats['skipped'] += 1
                continue

            logger.info(f"\n📦 提交 {idx+1}/{len(products)}: SKU={sku}")

            try:
                result = self.listings.put_listings_item(sku, product)
                status = result.get('status', 'UNKNOWN')

                if status == 'ACCEPTED':
                    self.stats['success'] += 1
                    logger.info(f"  ✅ 已接收")
                else:
                    self.stats['failed'] += 1
                    logger.warning(f"  ⚠️ 状态: {status}")

                results.append({
                    'sku': sku,
                    'status': status,
                    'issues': result.get('issues', []),
                    'submission_id': result.get('submissionId', ''),
                })
                self.stats['submitted'] += 1

            except Exception as e:
                logger.error(f"  ❌ 提交失败: {e}")
                self.stats['failed'] += 1
                results.append({
                    'sku': sku,
                    'status': 'ERROR',
                    'issues': [{'message': str(e)}],
                })

            # 限流：每秒最多1个请求
            time.sleep(1)

        return results

    def _submit_batch(self, products: List[Dict]) -> List[Dict]:
        """批量提交 (使用Feeds API)"""
        # 构建Feed items
        feed_items = []
        for product in products:
            if not product.get('sku'):
                self.stats['skipped'] += 1
                continue

            feed_items.append({
                'sku': product['sku'],
                'product_type': product.get('product_type', 'PRODUCT'),
                'attributes': self.mapper.build_listing_attributes(product),
            })

        if not feed_items:
            logger.warning("⚠️ 没有有效商品可提交")
            return []

        # 提交Feed
        result = self.feeds.submit_and_wait(
            items=feed_items,
            seller_id=self.config.AMAZON_SELLER_ID,
            timeout_minutes=10,
        )

        item_results = result.get('item_results') or []
        status = result.get('status', 'UNKNOWN')
        self.stats['submitted'] = len(feed_items)

        if item_results:
            success_statuses = {'ACCEPTED', 'ACCEPTED_WITH_WARNINGS'}
            self.stats['success'] = sum(
                1 for item in item_results
                if str(item.get('status', '')).upper() in success_statuses
            )
            self.stats['failed'] = len(item_results) - self.stats['success']
            logger.info(
                f"📋 批量结果: 成功 {self.stats['success']} 条, 失败 {self.stats['failed']} 条"
            )
            return item_results

        if status == 'DONE':
            self.stats['failed'] = len(feed_items)
            logger.warning("⚠️ Feed 已完成，但未返回逐条结果，按逐条状态未知处理")
            return [{
                'sku': item['sku'],
                'status': 'UNKNOWN',
                'issues': [{'message': 'Feed 已完成，但处理报告缺少逐条结果，请到提交记录中核对最终状态'}],
                'feed_id': result.get('feed_id', ''),
            } for item in feed_items]

        self.stats['failed'] = len(feed_items)
        logger.warning(f"⚠️ 批量提交状态: {status}")
        return [{
            'sku': item['sku'],
            'status': status,
            'issues': [{'message': f'Feed 状态: {status}'}],
            'feed_id': result.get('feed_id', ''),
        } for item in feed_items]

    def _generate_report(self, original_data: List[Dict],
                         results: List[Dict], output_path: str):
        """生成提交报告Excel"""
        # 合并结果到原始数据
        result_map = {r['sku']: r for r in results if isinstance(r, dict) and 'sku' in r}

        for item in original_data:
            sku = item.get('SKU') or item.get('sku') or item.get('商品编号')
            if sku and sku in result_map:
                r = result_map[sku]
                item['提交状态'] = r.get('status', '')
                item['提交ID'] = r.get('submission_id', '')
                issues = r.get('issues', [])
                if issues:
                    item['问题详情'] = '; '.join(
                        i.get('message', '') for i in issues
                    )

        extra_cols = ['提交状态', '提交ID', '问题详情']
        self.excel.write_output(original_data, output_path, extra_columns=extra_cols)

        # 同时输出JSON日志
        json_path = output_path.replace('.xlsx', '.json')
        with open(json_path, 'w', encoding='utf-8') as f:
            json.dump({
                'timestamp': datetime.now().isoformat(),
                'stats': self.stats,
                'results': results,
            }, f, ensure_ascii=False, indent=2)
        logger.info(f"📄 JSON日志: {json_path}")
