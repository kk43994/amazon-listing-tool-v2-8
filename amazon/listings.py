"""
Amazon SP-API Listings API V2
支持完整上架流程: 搜索目录→获取类型Schema→上架→状态追踪
"""
import logging
import requests
import time
from typing import Dict, Optional, List

from amazon.mapper import FieldMapper, ENDPOINTS, MARKETPLACE_REGION

logger = logging.getLogger(__name__)


# 常见 Amazon issue code → 中文说明 + 修复建议
_ISSUE_HINTS = {
    # 通用错误
    'MISSING_REQUIRED_ATTRIBUTE': ('缺少必填属性', '请在 Excel 中补充该字段的值'),
    'INVALID_ATTRIBUTE_VALUE': ('属性值无效', '请检查该字段的值是否符合亚马逊要求的格式'),
    'ATTRIBUTE_VALUE_TOO_LONG': ('属性值超长', '请缩短该字段的内容'),
    'ATTRIBUTE_VALUE_TOO_SHORT': ('属性值过短', '请补充更多内容'),
    'DUPLICATE_VALUE': ('重复值', '该字段的值与已有商品重复，请修改'),
    # 产品标识
    'INVALID_UPC': ('UPC 条码无效', '请核实 UPC 是否为有效的 12 位数字'),
    'INVALID_EAN': ('EAN 条码无效', '请核实 EAN 是否为有效的 13 位数字'),
    'INVALID_GTIN': ('GTIN 无效', '请检查产品标识码是否正确'),
    'PRODUCT_IDENTIFIER_NOT_FOUND': ('产品标识未找到', 'UPC/EAN 在 GS1 数据库中未找到，请确认条码正确或申请豁免'),
    # 图片
    'INVALID_IMAGE_URL': ('图片 URL 无效', '请确保图片链接可访问且为 JPEG/PNG 格式'),
    'IMAGE_TOO_SMALL': ('图片尺寸过小', '主图最小 1000x1000 像素，建议 2000x2000'),
    # 价格
    'INVALID_PRICE': ('价格无效', '请检查价格格式（数字，无货币符号）'),
    'PRICE_TOO_LOW': ('价格过低', '该价格低于亚马逊允许的最低价格'),
    'PRICE_TOO_HIGH': ('价格过高', '该价格高于亚马逊允许的最高价格'),
    # 变体
    'MISSING_PARENT_SKU': ('缺少父体 SKU', '子体商品必须填写 parent_sku'),
    'INVALID_VARIATION_THEME': ('变体主题无效', '请检查 variation_theme 是否为该类目支持的值'),
    # 权限
    'UNAUTHORIZED': ('无权操作', '请检查账号是否有该类目的销售权限'),
    'BRAND_NOT_APPROVED': ('品牌未授权', '请先在 Brand Registry 中注册或获得品牌授权'),
}


class ListingsAPI:
    """Amazon SP-API Listings + Catalog + Product Type Definitions"""

    def __init__(self, auth, seller_id: str, marketplace_id: str = "ATVPDKIKX0DER"):
        self.auth = auth
        self.seller_id = seller_id
        self.marketplace_id = marketplace_id
        self.mapper = FieldMapper(marketplace_id)
        self.region = MARKETPLACE_REGION.get(marketplace_id, 'NA')
        self.base_url = ENDPOINTS.get(self.region, ENDPOINTS['NA'])

    # ===== 1. 目录搜索 (Catalog Items API) =====

    def search_catalog(self, keywords: str = None,
                       identifiers: List[str] = None,
                       identifier_type: str = 'UPC') -> Optional[Dict]:
        """
        搜索亚马逊目录，检查商品是否已存在

        Args:
            keywords: 关键词搜索
            identifiers: 标识符搜索(UPC/EAN/ASIN)
            identifier_type: 标识符类型

        Returns:
            搜索结果 或 None
        """
        url = f"{self.base_url}/catalog/2022-04-01/items"
        params = {
            'marketplaceIds': self.marketplace_id,
            'includedData': 'summaries,identifiers',
            'pageSize': 5,
        }

        if identifiers:
            params['identifiers'] = ','.join(identifiers)
            params['identifiersType'] = identifier_type
        elif keywords:
            params['keywords'] = keywords
        else:
            return None

        logger.info("🔍 搜索亚马逊目录...")
        try:
            response = requests.get(
                url, params=params,
                headers=self.auth.get_headers(),
                timeout=30,
            )
            if response.status_code == 200:
                result = response.json()
                items = result.get('items', [])
                logger.info(f"  找到 {len(items)} 个匹配商品")
                return result
            else:
                logger.warning(f"  目录搜索失败: {response.status_code}")
                return None
        except Exception as e:
            logger.error(f"  目录搜索异常: {e}")
            return None

    # ===== 2. 产品类型定义 (Product Type Definitions API) =====

    def get_product_type_schema(self, product_type: str,
                                requirements: str = 'LISTING') -> Optional[Dict]:
        """
        获取产品类型的JSON Schema(字段要求)

        Args:
            product_type: 产品类型(如 LUGGAGE, WIRELESS_ACCESSORY)
            requirements: LISTING / LISTING_PRODUCT_ONLY / LISTING_OFFER_ONLY

        Returns:
            ProductTypeDefinition 包含 propertyGroups 和 schema link
        """
        url = f"{self.base_url}/definitions/2020-09-01/productTypes/{product_type}"
        params = {
            'marketplaceIds': self.marketplace_id,
            'requirements': requirements,
            'locale': 'en_US',
            'sellerId': self.seller_id,
        }

        logger.info(f"📋 获取产品类型定义: {product_type}")
        try:
            response = requests.get(
                url, params=params,
                headers=self.auth.get_headers(),
                timeout=30,
            )
            if response.status_code == 200:
                result = response.json()
                groups = result.get('propertyGroups', {})
                logger.info(f"  属性分组: {list(groups.keys())}")
                return result
            else:
                logger.warning(f"  获取失败: {response.status_code} {response.text[:200]}")
                return None
        except Exception as e:
            logger.error(f"  获取异常: {e}")
            return None

    def search_product_types(self, item_name: str = None) -> Optional[List[Dict]]:
        """
        搜索可用的产品类型

        Args:
            item_name: 商品名称(用于获取推荐)

        Returns:
            产品类型列表
        """
        url = f"{self.base_url}/definitions/2020-09-01/productTypes"
        params = {
            'marketplaceIds': self.marketplace_id,
        }
        if item_name:
            params['itemName'] = item_name

        try:
            response = requests.get(
                url, params=params,
                headers=self.auth.get_headers(),
                timeout=30,
            )
            if response.status_code == 200:
                return response.json().get('productTypes', [])
            return None
        except Exception as e:
            logger.error(f"  搜索产品类型异常: {e}")
            return None

    # ===== 3. Listings Items API =====

    def put_listings_item(self, sku: str, product_data: Dict,
                          preview: bool = False) -> Dict:
        """
        创建或完全更新一个Listing

        Args:
            sku: 商品SKU
            product_data: 标准化产品数据
            preview: 是否仅预览验证(不实际提交)

        Returns:
            API响应 {status, submissionId, issues, identifiers}
        """
        url = f"{self.base_url}/listings/2021-08-01/items/{self.seller_id}/{sku}"
        params = {
            'marketplaceIds': self.marketplace_id,
            'includedData': 'issues,identifiers' if preview else 'issues',
        }

        if preview:
            params['mode'] = 'VALIDATION_PREVIEW'

        # 构建请求体
        body = self.mapper.build_put_body(product_data)

        action = "预览验证" if preview else "提交上架"
        logger.info(f"📤 {action}: SKU={sku}, 类型={body['productType']}")

        retried_auth = False
        while True:
            try:
                response = requests.put(
                    url, params=params,
                    headers=self.auth.get_headers(),
                    json=body,
                    timeout=30,
                )

                # Token 过期时自动刷新并重试一次
                if response.status_code in (401, 403) and not retried_auth:
                    logger.warning("  🔄 Token 可能过期，刷新后重试...")
                    retried_auth = True
                    self.auth.force_refresh()
                    continue

                try:
                    result = response.json()
                except ValueError:
                    result = {}

                result = self._normalize_put_response(response.status_code, result)
                status = result.get('status', 'UNKNOWN')

                if status in ('ACCEPTED', 'ACCEPTED_WITH_WARNINGS'):
                    asin = self.resolve_submission_asin(sku, result) if not preview else ''
                    if asin:
                        result['asin'] = asin
                    logger.info(
                        f"  ✅ 已接受{'（含警告）' if status == 'ACCEPTED_WITH_WARNINGS' else ''} "
                        f"(submissionId: {result.get('submissionId', 'N/A')})"
                    )
                    if result.get('asin'):
                        logger.info(f"  📌 ASIN: {result['asin']}")
                elif status == 'VALID':
                    logger.info("  ✅ 验证通过 (预览模式)")
                elif status == 'INVALID':
                    logger.error("  ❌ 验证失败")
                else:
                    logger.warning(f"  ⚠️ 状态: {status}")

                # 输出issues
                issues = result.get('issues', [])
                for issue in issues:
                    severity = issue.get('severity', 'INFO')
                    code = issue.get('code', '')
                    msg = issue.get('message', '')
                    attrs = issue.get('attributeNames', [])
                    hint = issue.get('hint', '')
                    fix = issue.get('fix', '')
                    icon = {'ERROR': '❌', 'WARNING': '⚠️', 'INFO': 'ℹ️'}.get(severity, '?')
                    line = f"  {icon} [{code}] {msg} (字段: {', '.join(attrs)})"
                    if hint:
                        line += f"\n       → {hint}"
                    if fix:
                        line += f"\n       💡 {fix}"
                    logger.log(
                        logging.ERROR if severity == 'ERROR' else logging.WARNING,
                        line,
                    )

                return result

            except requests.exceptions.Timeout:
                logger.error("  ❌ 请求超时")
                return {'status': 'ERROR', 'issues': [{'message': '请求超时，请稍后重试'}]}
            except Exception as e:
                logger.error(f"  ❌ 请求异常: {e}")
                return {'status': 'ERROR', 'issues': [{'message': str(e)}]}

    def _normalize_put_response(self, status_code: int, result: Dict) -> Dict:
        """把 Amazon 顶层 errors 统一折叠成 issues，便于前后端复用。"""
        payload = dict(result or {})
        raw_errors = payload.get('errors') or []
        if raw_errors and not payload.get('issues'):
            payload['issues'] = [{
                'code': err.get('code', ''),
                'message': err.get('message', ''),
                'severity': 'ERROR',
                'attributeNames': err.get('attributeNames', []),
            } for err in raw_errors]

        if 'status' not in payload:
            if payload.get('issues'):
                payload['status'] = 'ERROR' if status_code >= 400 else 'INVALID'
            elif status_code >= 400:
                payload['status'] = 'ERROR'

        payload.setdefault('http_status', status_code)

        # 为每个 issue 追加中文说明和修复建议
        for issue in payload.get('issues', []):
            code = str(issue.get('code', '')).strip()
            hint = _ISSUE_HINTS.get(code)
            if hint:
                issue.setdefault('hint', hint[0])
                issue.setdefault('fix', hint[1])
            attrs = issue.get('attributeNames', [])
            if attrs and 'hint' not in issue:
                issue['hint'] = f"相关字段: {', '.join(attrs)}"

        return payload

    def _extract_asin(self, payload: Optional[Dict]) -> str:
        payload = payload or {}
        for ident in payload.get('identifiers', []) or []:
            asin = str(ident.get('asin', '') or '').strip()
            if asin:
                return asin
        for summary in payload.get('summaries', []) or []:
            asin = str(summary.get('asin', '') or '').strip()
            if asin:
                return asin
        return str(payload.get('asin', '') or '').strip()

    def resolve_submission_asin(self, sku: str, submit_result: Optional[Dict] = None,
                                max_attempts: int = 2, delay: float = 0.5) -> str:
        """优先读取提交响应中的 ASIN，缺失时补查 Listings API。"""
        asin = self._extract_asin(submit_result)
        if asin:
            return asin

        status = str((submit_result or {}).get('status', '') or '').strip().upper()
        if status != 'ACCEPTED':
            return ''

        attempts = max(1, int(max_attempts or 1))
        wait_seconds = max(0.0, float(delay or 0.0))
        for attempt in range(attempts):
            listing_payload = self.get_listings_item(sku)
            asin = self._extract_asin(listing_payload)
            if asin:
                return asin
            if attempt < attempts - 1 and wait_seconds > 0:
                time.sleep(wait_seconds)
        return ''

    def patch_listings_item(self, sku: str, patches: List[Dict]) -> Dict:
        """部分更新Listing"""
        url = f"{self.base_url}/listings/2021-08-01/items/{self.seller_id}/{sku}"
        params = {'marketplaceIds': self.marketplace_id}

        body = {
            'productType': 'PRODUCT',
            'patches': patches,
        }

        logger.info(f"📝 局部更新: SKU={sku}, {len(patches)}个patch")
        try:
            response = requests.patch(
                url, params=params,
                headers=self.auth.get_headers(),
                json=body,
                timeout=30,
            )
            return response.json()
        except Exception as e:
            return {'status': 'ERROR', 'issues': [{'message': str(e)}]}

    def get_listings_item(self, sku: str) -> Optional[Dict]:
        """获取Listing信息"""
        url = f"{self.base_url}/listings/2021-08-01/items/{self.seller_id}/{sku}"
        params = {
            'marketplaceIds': self.marketplace_id,
            'includedData': 'summaries,attributes,issues,offers,fulfillmentAvailability',
            'issueLocale': 'en_US',
        }

        try:
            response = requests.get(
                url, params=params,
                headers=self.auth.get_headers(),
                timeout=30,
            )
            if response.status_code == 200:
                return response.json()
            elif response.status_code == 404:
                return None
            else:
                logger.error(f"获取Listing失败: {response.status_code}")
                return None
        except Exception as e:
            logger.error(f"获取Listing异常: {e}")
            return None

    def probe_connection(self, probe_sku: str = None) -> Dict:
        """用一次 seller-scoped Listings 请求验证账号、站点和权限是否可用。"""
        probe_sku = str(probe_sku or f"__sp_api_probe__{int(time.time())}").strip()
        url = f"{self.base_url}/listings/2021-08-01/items/{self.seller_id}/{probe_sku}"
        params = {
            'marketplaceIds': self.marketplace_id,
            'includedData': 'summaries',
            'issueLocale': 'en_US',
        }

        try:
            response = requests.get(
                url,
                params=params,
                headers=self.auth.get_headers(),
                timeout=30,
            )
        except requests.exceptions.Timeout:
            return {'success': False, 'message': 'Listings API 探测超时'}
        except Exception as e:
            return {'success': False, 'message': f'Listings API 探测异常: {e}'}

        if response.status_code in (200, 404):
            return {
                'success': True,
                'status_code': response.status_code,
                'message': 'Listings API 可访问',
            }

        detail = ''
        try:
            payload = response.json()
            detail = (
                payload.get('message')
                or payload.get('errors', [{}])[0].get('message', '')
                or response.text[:200]
            )
        except Exception:
            detail = response.text[:200]

        detail = str(detail or '').strip()
        if detail:
            detail = f": {detail}"

        return {
            'success': False,
            'status_code': response.status_code,
            'message': f'Listings API 探测失败({response.status_code}){detail}',
        }

    def delete_listings_item(self, sku: str) -> Dict:
        """删除Listing"""
        url = f"{self.base_url}/listings/2021-08-01/items/{self.seller_id}/{sku}"
        params = {'marketplaceIds': self.marketplace_id}

        logger.info(f"🗑️ 删除Listing: SKU={sku}")
        try:
            response = requests.delete(
                url, params=params,
                headers=self.auth.get_headers(),
                timeout=30,
            )
            return response.json()
        except Exception as e:
            return {'status': 'ERROR', 'issues': [{'message': str(e)}]}

    # ===== 4. 批量提交流程 =====

    def submit_listings(self, products: List[Dict],
                        preview_first: bool = True,
                        delay: float = 0.3) -> List[Dict]:
        """
        批量提交Listings

        Args:
            products: 标准化产品数据列表(必须有sku)
            preview_first: 是否先预览验证
            delay: 每次请求间隔(秒,避免限流)

        Returns:
            提交结果列表
        """
        results = []
        total = len(products)

        processed_skus = set()

        for idx, product in enumerate(products):
            sku = str(product.get('sku', '') or '').strip()
            if not sku:
                results.append({
                    'sku': 'MISSING',
                    'status': 'SKIPPED',
                    'issues': [{'message': '缺少SKU'}]
                })
                continue
            if sku in processed_skus:
                logger.warning(f"  ⚠️ 检测到重复 SKU: {sku}，跳过重复项")
                results.append({
                    'sku': sku,
                    'status': 'SKIPPED',
                    'issues': [{'message': f'重复 SKU 已跳过: {sku}'}],
                })
                continue
            processed_skus.add(sku)

            logger.info(f"\n--- 提交 {idx+1}/{total}: SKU={sku} ---")

            # 字段验证
            validation = self.mapper.validate_required_fields(product)
            if not validation['valid']:
                logger.error(f"  ❌ 验证失败: {validation['errors']}")
                results.append({
                    'sku': sku,
                    'status': 'VALIDATION_ERROR',
                    'issues': [{'message': e} for e in validation['errors']],
                    'warnings': validation['warnings'],
                })
                continue

            # 预览验证
            if preview_first:
                preview_result = self.put_listings_item(sku, product, preview=True)
                preview_status = str(preview_result.get('status', '') or '').strip().upper()
                preview_issues = preview_result.get('issues', []) or []
                preview_has_error = any(
                    str(issue.get('severity', '') or '').strip().upper() == 'ERROR'
                    for issue in preview_issues
                )
                if preview_status != 'VALID' or preview_has_error:
                    results.append({
                        'sku': sku,
                        'status': 'PREVIEW_INVALID' if preview_status == 'INVALID' else f'PREVIEW_{preview_status or "ERROR"}',
                        'issues': preview_issues or [{'message': f'Amazon 预览未通过: {preview_status or "UNKNOWN"}'}],
                    })
                    continue
                time.sleep(delay)

            # 正式提交
            submit_result = self.put_listings_item(sku, product, preview=False)
            submit_result['sku'] = sku
            submit_result['submit_time'] = time.strftime('%Y-%m-%d %H:%M:%S')

            asin = self.resolve_submission_asin(sku, submit_result, delay=delay)
            if asin:
                submit_result['asin'] = asin

            results.append(submit_result)
            time.sleep(delay)

        # 汇总
        accepted = sum(
            1 for r in results
            if str(r.get('status', '')).upper() in ('ACCEPTED', 'ACCEPTED_WITH_WARNINGS')
        )
        invalid = sum(1 for r in results if 'INVALID' in r.get('status', ''))
        errors = sum(1 for r in results if r.get('status') == 'ERROR')

        logger.info(f"\n{'='*50}")
        logger.info("📊 批量提交完成:")
        logger.info(f"  ✅ 接受: {accepted}")
        logger.info(f"  ❌ 失败: {invalid + errors}")
        logger.info(f"  总计: {total}")

        return results
