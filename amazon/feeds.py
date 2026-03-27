"""
Amazon SP-API Feeds API 模块
用于批量提交商品数据 (JSON Feed)
"""
import json
import time
import logging
import requests
from typing import Dict, List, Optional

from amazon.mapper import ENDPOINTS, MARKETPLACE_REGION

logger = logging.getLogger(__name__)


class FeedsAPI:
    """Amazon Feeds API 批量提交"""

    def __init__(self, auth, marketplace_id: str = "ATVPDKIKX0DER"):
        """
        初始化 Feeds API

        Args:
            auth: AmazonAuth 实例
            marketplace_id: 市场ID
        """
        self.auth = auth
        self.marketplace_id = marketplace_id
        region = MARKETPLACE_REGION.get(marketplace_id, 'NA')
        self.base_url = ENDPOINTS.get(region, ENDPOINTS['NA'])

    def create_feed_document(self, content_type: str = "application/json") -> Dict:
        """
        创建 Feed Document (获取上传URL)

        Returns:
            {feedDocumentId, url} 用于上传内容
        """
        url = f"{self.base_url}/feeds/2021-06-30/documents"
        body = {"contentType": content_type}

        response = requests.post(
            url,
            headers=self.auth.get_headers(),
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        return response.json()

    def upload_feed_content(self, upload_url: str, content: str):
        """
        上传 Feed 内容到预签名URL

        Args:
            upload_url: 从 create_feed_document 获取的URL
            content: JSON格式的Feed内容
        """
        response = requests.put(
            upload_url,
            data=content.encode('utf-8'),
            headers={'Content-Type': 'application/json; charset=UTF-8'},
            timeout=60,
        )
        response.raise_for_status()
        logger.info("✅ Feed内容上传成功")

    def create_feed(self, feed_document_id: str,
                    feed_type: str = "JSON_LISTINGS_FEED") -> str:
        """
        创建并提交 Feed

        Args:
            feed_document_id: Feed Document ID
            feed_type: Feed类型

        Returns:
            feedId
        """
        url = f"{self.base_url}/feeds/2021-06-30/feeds"
        body = {
            "feedType": feed_type,
            "marketplaceIds": [self.marketplace_id],
            "inputFeedDocumentId": feed_document_id,
        }

        response = requests.post(
            url,
            headers=self.auth.get_headers(),
            json=body,
            timeout=30,
        )
        response.raise_for_status()
        result = response.json()
        feed_id = result['feedId']
        logger.info(f"📋 Feed已提交: {feed_id}")
        return feed_id

    def get_feed(self, feed_id: str) -> Dict:
        """
        获取 Feed 处理状态

        Returns:
            Feed状态信息
        """
        url = f"{self.base_url}/feeds/2021-06-30/feeds/{feed_id}"
        response = requests.get(
            url,
            headers=self.auth.get_headers(),
            timeout=30,
        )
        response.raise_for_status()
        return response.json()

    def get_feed_result(self, feed_document_id: str) -> Optional[Dict]:
        """
        获取 Feed 处理结果文档

        Returns:
            处理结果 或 None
        """
        url = f"{self.base_url}/feeds/2021-06-30/documents/{feed_document_id}"
        response = requests.get(
            url,
            headers=self.auth.get_headers(),
            timeout=30,
        )
        if response.status_code != 200:
            return None

        doc = response.json()
        # 下载结果内容
        result_url = doc.get('url')
        if result_url:
            result_response = requests.get(result_url, timeout=30)
            try:
                return result_response.json()
            except:
                return {'raw': result_response.text}
        return doc

    def submit_and_wait(self, items: List[Dict], seller_id: str,
                        timeout_minutes: int = 10,
                        poll_interval: int = 30) -> Dict:
        """
        完整的Feed提交+等待流程

        Args:
            items: 商品数据列表
            seller_id: 卖家ID
            timeout_minutes: 等待超时(分钟)
            poll_interval: 轮询间隔(秒)

        Returns:
            {feed_id, status, result}
        """
        # 1. 构建JSON Feed内容
        feed_content = self._build_json_feed(items, seller_id)
        message_lookup = {
            message['messageId']: message['sku']
            for message in feed_content.get('messages', [])
        }
        logger.info(f"📦 准备提交 {len(items)} 条商品")

        # 2. 创建Feed Document
        doc = self.create_feed_document()
        feed_doc_id = doc['feedDocumentId']
        upload_url = doc['url']

        # 3. 上传内容
        self.upload_feed_content(upload_url, json.dumps(feed_content, ensure_ascii=False))

        # 4. 创建Feed
        feed_id = self.create_feed(feed_doc_id)

        # 5. 轮询等待结果
        logger.info(f"⏳ 等待平台处理 (最多{timeout_minutes}分钟)...")
        deadline = time.time() + timeout_minutes * 60

        while time.time() < deadline:
            time.sleep(poll_interval)
            feed_status = self.get_feed(feed_id)
            status = feed_status.get('processingStatus', 'UNKNOWN')
            logger.info(f"  状态: {status}")

            if status in ('DONE', 'FATAL'):
                # 获取结果
                result_doc_id = feed_status.get('resultFeedDocumentId')
                result = None
                if result_doc_id:
                    result = self.get_feed_result(result_doc_id)

                return {
                    'feed_id': feed_id,
                    'status': status,
                    'result': result,
                    'item_results': self._parse_processing_report(
                        result=result,
                        message_lookup=message_lookup,
                        default_status='ACCEPTED' if status == 'DONE' else status,
                    ),
                    'processing_status': feed_status,
                }

            elif status == 'CANCELLED':
                return {
                    'feed_id': feed_id,
                    'status': 'CANCELLED',
                    'result': None,
                }

        logger.warning(f"⚠️ 等待超时 ({timeout_minutes}分钟)")
        return {
            'feed_id': feed_id,
            'status': 'TIMEOUT',
            'result': None,
            'item_results': [],
        }

    def _build_json_feed(self, items: List[Dict], seller_id: str) -> Dict:
        """
        构建 JSON Listings Feed 内容

        Args:
            items: 商品数据列表 (每项需包含 sku 和 attributes)
            seller_id: 卖家ID
        """
        messages = []
        for idx, item in enumerate(items):
            messages.append({
                "messageId": idx + 1,
                "sku": item['sku'],
                "operationType": "PARTIAL_UPDATE" if item.get('update') else "UPDATE",
                "productType": item.get('product_type', 'PRODUCT'),
                "attributes": item.get('attributes', {}),
            })

        return {
            "header": {
                "sellerId": seller_id,
                "version": "2.0",
                "issueLocale": "en_US",
                "report": {
                    "includedData": ["issues"],
                },
            },
            "messages": messages,
        }

    def _parse_processing_report(self, result: Optional[Dict],
                                 message_lookup: Dict[int, str],
                                 default_status: str) -> List[Dict]:
        """将 Feed 处理结果尽量还原为逐 SKU 状态。"""
        if not message_lookup:
            return []

        parsed = {
            message_id: {
                'sku': sku,
                'status': default_status,
                'issues': [],
            }
            for message_id, sku in message_lookup.items()
        }

        if not isinstance(result, dict):
            return list(parsed.values())

        issues_sources = []
        for key in ('issues', 'results', 'messages'):
            value = result.get(key)
            if isinstance(value, list):
                issues_sources.extend(value)

        processing_report = result.get('processingReport')
        if isinstance(processing_report, dict):
            for key in ('issues', 'results', 'messages'):
                value = processing_report.get(key)
                if isinstance(value, list):
                    issues_sources.extend(value)

        for entry in issues_sources:
            if not isinstance(entry, dict):
                continue

            message_id = entry.get('messageId') or entry.get('message_id')
            try:
                message_id = int(message_id)
            except (TypeError, ValueError):
                message_id = None
            if message_id not in parsed:
                continue

            item = parsed[message_id]
            item_status = self._derive_item_status(entry, default_status)
            if item_status:
                item['status'] = item_status

            item['issues'].extend(self._extract_issue_messages(entry))

        return list(parsed.values())

    def _derive_item_status(self, entry: Dict, default_status: str) -> str:
        status = (
            entry.get('status')
            or entry.get('processingStatus')
            or entry.get('resultCode')
            or entry.get('code')
            or ''
        )
        text = str(status).strip().upper()
        if not text:
            severity = str(entry.get('severity', '')).strip().upper()
            if severity == 'ERROR':
                return 'INVALID'
            if severity == 'WARNING':
                return 'ACCEPTED_WITH_WARNINGS'
            return default_status

        if text in ('SUCCESS', 'DONE', 'OK', 'ACCEPTED'):
            return 'ACCEPTED'
        if text in ('WARNING', 'WARN'):
            return 'ACCEPTED_WITH_WARNINGS'
        if 'ERROR' in text or 'INVALID' in text or 'FAIL' in text or 'FATAL' in text:
            return 'INVALID'
        return text

    def _extract_issue_messages(self, entry: Dict) -> List[Dict]:
        issues = []

        def append_issue(message, severity='INFO', code=''):
            text = str(message or '').strip()
            if not text:
                return
            issues.append({
                'severity': str(severity or 'INFO').upper(),
                'code': str(code or ''),
                'message': text,
            })

        nested_lists = []
        for key in ('issues', 'errors', 'warnings'):
            value = entry.get(key)
            if isinstance(value, list):
                nested_lists.extend(value)

        for nested in nested_lists:
            if isinstance(nested, dict):
                append_issue(
                    nested.get('message') or nested.get('description') or nested.get('details'),
                    severity=nested.get('severity') or nested.get('level') or entry.get('severity'),
                    code=nested.get('code') or entry.get('code'),
                )
            else:
                append_issue(nested, severity=entry.get('severity'), code=entry.get('code'))

        if not issues:
            append_issue(
                entry.get('message') or entry.get('description') or entry.get('details'),
                severity=entry.get('severity'),
                code=entry.get('code'),
            )

        return issues
