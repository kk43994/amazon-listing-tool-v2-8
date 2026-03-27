"""
统一 SP-API 客户端
使用 python-amazon-sp-api 库替换手写 HTTP 调用
"""
import logging
from typing import Dict, List, Optional

from sp_api.api import ProductTypeDefinitions, ListingsItemsV20210801, CatalogItemsV20220401
from sp_api.base.marketplaces import Marketplaces

from amazon.accounts import AccountManager

logger = logging.getLogger(__name__)

# marketplace_id -> Marketplaces enum
_MP_MAP = {
    'ATVPDKIKX0DER': Marketplaces.US,
    'A2EUQ1WTGCTBG2': Marketplaces.CA,
    'A1AM78C64UM0Y8': Marketplaces.MX,
    'A1F83G8C2ARO7P': Marketplaces.UK,
    'A1PA6795UKMFR9': Marketplaces.DE,
    'A13V1IB3VIYZZH': Marketplaces.FR,
    'APJ6JRA9NG5V4': Marketplaces.IT,
    'A1RKKUPIHCS9HS': Marketplaces.ES,
    'A1VC38T7YXB528': Marketplaces.JP,
    'A39IBJ37TRP1C6': Marketplaces.AU,
}


class SPClient:
    """统一 SP-API 客户端，从 accounts.json 读取凭证"""

    def __init__(self, seller_id: str = None, marketplace_id: str = None):
        mgr = AccountManager()
        acc = mgr.get_account(seller_id) if seller_id else mgr.get_default_account()
        if not acc:
            raise ValueError("未找到可用的亚马逊账号，请先在 accounts.json 中配置")

        self.seller_id = acc['seller_id']
        self.marketplace_id = marketplace_id or acc.get('marketplace_id', 'ATVPDKIKX0DER')
        self.marketplace = _MP_MAP.get(self.marketplace_id, Marketplaces.US)
        self.credentials = {
            'lwa_app_id': acc['lwa_client_id'],
            'lwa_client_secret': acc['lwa_client_secret'],
            'refresh_token': acc['refresh_token'],
        }

    def _make_client(self, cls):
        return cls(
            marketplace=self.marketplace,
            credentials=self.credentials,
            refresh_token=self.credentials['refresh_token'],
        )

    def search_product_types(self, keyword: str) -> List[Dict]:
        """搜索推荐的产品类型"""
        client = self._make_client(ProductTypeDefinitions)
        resp = client.search_definitions_product_types(
            itemName=keyword,
            marketplaceIds=[self.marketplace_id],
        )
        return resp.payload.get('productTypes', [])

    def get_schema(self, product_type: str) -> Dict:
        """获取产品类型的 JSON Schema 定义"""
        client = self._make_client(ProductTypeDefinitions)
        resp = client.get_definitions_product_type(
            product_type,
            marketplaceIds=[self.marketplace_id],
            requirements='LISTING',
            locale='en_US',
            sellerId=self.seller_id,
        )
        return resp.payload

    def put_listing(self, sku: str, body: Dict) -> Dict:
        """创建或更新 Listing"""
        client = self._make_client(ListingsItemsV20210801)
        resp = client.put_listings_item(
            self.seller_id, sku,
            body=body,
            marketplaceIds=[self.marketplace_id],
        )
        return resp.payload

    def search_catalog(self, keywords: str) -> Dict:
        """搜索亚马逊目录"""
        client = self._make_client(CatalogItemsV20220401)
        resp = client.search_catalog_items(
            keywords=[keywords],
            marketplaceIds=[self.marketplace_id],
            includedData=['summaries', 'identifiers'],
            pageSize=5,
        )
        return resp.payload
