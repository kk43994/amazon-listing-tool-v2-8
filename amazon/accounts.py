"""
亚马逊卖家账号管理
支持多账号配置、切换、测试连接
"""
import json
import os
import logging
from typing import Dict, List, Optional

from amazon.auth import AmazonAuth, AmazonAuthError, AmazonTokenError, AmazonNetworkError
from amazon.listings import ListingsAPI
from amazon.mapper import MARKETPLACE_IDS, ENDPOINTS, MARKETPLACE_REGION
from core.runtime_paths import runtime_path

logger = logging.getLogger(__name__)

# 默认配置文件路径
DEFAULT_ACCOUNTS_FILE = runtime_path('accounts.json')


def _is_real_credential(value: str) -> bool:
    text = str(value or '').strip()
    return bool(text) and 'YOUR_' not in text.upper()


def account_has_real_credentials(account: Optional[Dict], require_seller_id: bool = True) -> bool:
    """判断账号是否已填入可用凭证，而不是仓库里的模板占位值。"""
    if not isinstance(account, dict):
        return False

    required_fields = ['lwa_client_id', 'lwa_client_secret', 'refresh_token']
    if require_seller_id:
        required_fields.insert(0, 'seller_id')

    return all(_is_real_credential(account.get(field)) for field in required_fields)


class AccountManager:
    """亚马逊多账号管理器"""

    def __init__(self, accounts_file: str = None):
        self.accounts_file = accounts_file or DEFAULT_ACCOUNTS_FILE
        self.accounts: List[Dict] = []
        self._load_accounts()

    def _load_accounts(self):
        """加载账号配置"""
        if os.path.exists(self.accounts_file):
            try:
                self._secure_file_permissions()
                with open(self.accounts_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                self.accounts = data.get('accounts', [])
                logger.info(f"📂 加载 {len(self.accounts)} 个亚马逊账号")
            except Exception as e:
                logger.error(f"加载账号配置失败: {e}")
                self.accounts = []
        else:
            logger.info("📂 账号配置文件不存在，创建模板...")
            self._create_template()

    def _create_template(self):
        """创建配置模板"""
        template = {
            "accounts": [
                {
                    "name": "示例账号-美国站",
                    "seller_id": "YOUR_SELLER_ID",
                    "marketplace_id": "ATVPDKIKX0DER",
                    "marketplace_name": "Amazon US",
                    "lwa_client_id": "amzn1.application-oa2-client.YOUR_CLIENT_ID",
                    "lwa_client_secret": "YOUR_CLIENT_SECRET",
                    "refresh_token": "Atzr|YOUR_REFRESH_TOKEN",
                    "is_default": True,
                    "enabled": True
                }
            ],
            "_注释": {
                "seller_id": "卖家ID，在Seller Central的Account Info中查看",
                "marketplace_id": "参考: US=ATVPDKIKX0DER, UK=A1F83G8C2ARO7P, DE=A1PA6795UKMFR9",
                "lwa_client_id": "在Developer Central的应用信息中获取",
                "lwa_client_secret": "在Developer Central的应用信息中获取",
                "refresh_token": "通过OAuth2授权流程获取",
            }
        }
        with open(self.accounts_file, 'w', encoding='utf-8') as f:
            json.dump(template, f, ensure_ascii=False, indent=2)
        self.accounts = template['accounts']
        self._secure_file_permissions()
        logger.info(f"  模板已创建: {self.accounts_file}")

    def save_accounts(self):
        """保存账号配置"""
        data = {'accounts': self.accounts}
        with open(self.accounts_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self._secure_file_permissions()
        logger.info("💾 账号配置已保存")

    def _secure_file_permissions(self):
        """尽量收紧账号配置权限，避免凭证被同机其他用户读取。"""
        if os.name == 'nt' or not os.path.exists(self.accounts_file):
            return
        try:
            os.chmod(self.accounts_file, 0o600)
        except OSError as exc:
            logger.warning("收紧账号配置文件权限失败: %s", exc)

    def list_accounts(self) -> List[Dict]:
        """列出所有账号(隐藏敏感信息)"""
        safe_list = []
        for acc in self.accounts:
            safe = {
                'name': acc.get('name', ''),
                'seller_id': acc.get('seller_id', ''),
                'marketplace_id': acc.get('marketplace_id', ''),
                'marketplace_name': acc.get('marketplace_name', ''),
                'is_default': acc.get('is_default', False),
                'enabled': acc.get('enabled', True),
                'has_credentials': account_has_real_credentials(acc),
            }
            safe_list.append(safe)
        return safe_list

    def get_default_account(self) -> Optional[Dict]:
        """获取默认账号"""
        for acc in self.accounts:
            if acc.get('is_default') and acc.get('enabled', True):
                return acc
        # 没有默认，返回第一个启用的
        for acc in self.accounts:
            if acc.get('enabled', True):
                return acc
        return None

    def get_account(self, name_or_id: str) -> Optional[Dict]:
        """通过名称或seller_id获取账号"""
        for acc in self.accounts:
            if (acc.get('name') == name_or_id or
                acc.get('seller_id') == name_or_id):
                return acc
        return None

    def add_account(self, account: Dict) -> bool:
        """添加新账号"""
        # 检查重复
        for acc in self.accounts:
            if acc.get('seller_id') == account.get('seller_id'):
                logger.warning(f"账号已存在: {account.get('seller_id')}")
                return False

        # 如果是第一个，设为默认
        if not self.accounts:
            account['is_default'] = True

        self.accounts.append(account)
        self.save_accounts()
        logger.info(f"✅ 添加账号: {account.get('name')}")
        return True

    def update_account(self, seller_id: str, updates: Dict) -> bool:
        """更新账号信息"""
        for acc in self.accounts:
            if acc.get('seller_id') == seller_id:
                acc.update(updates)
                self.save_accounts()
                logger.info(f"✅ 更新账号: {seller_id}")
                return True
        return False

    def remove_account(self, seller_id: str) -> bool:
        """删除账号"""
        self.accounts = [a for a in self.accounts
                        if a.get('seller_id') != seller_id]
        self.save_accounts()
        return True

    def set_default(self, seller_id: str) -> bool:
        """设置默认账号"""
        found = False
        for acc in self.accounts:
            if acc.get('seller_id') == seller_id:
                acc['is_default'] = True
                found = True
            else:
                acc['is_default'] = False
        if found:
            self.save_accounts()
        return found

    def test_connection(self, seller_id: str = None) -> Dict:
        """
        测试账号连接

        Returns:
            {'success': bool, 'message': str, 'seller_info': {...}}
        """
        acc = self.get_account(seller_id) if seller_id else self.get_default_account()
        if not acc:
            return {'success': False, 'message': '未找到账号'}

        if not all([acc.get('lwa_client_id'),
                    acc.get('lwa_client_secret'),
                    acc.get('refresh_token')]):
            return {'success': False, 'message': '凭证不完整(缺少client_id/secret/refresh_token)'}
        if not account_has_real_credentials(acc):
            return {'success': False, 'message': '凭证未配置完成，仍包含模板占位值'}

        try:
            auth = AmazonAuth(
                client_id=acc['lwa_client_id'],
                client_secret=acc['lwa_client_secret'],
                refresh_token=acc['refresh_token'],
            )
            # 尝试获取token
            token = auth.get_access_token()
            if not token:
                return {'success': False, 'message': 'Token获取失败'}

            listings = ListingsAPI(
                auth=auth,
                seller_id=acc.get('seller_id', ''),
                marketplace_id=acc.get('marketplace_id', 'ATVPDKIKX0DER'),
            )
            probe = listings.probe_connection()
            if not probe.get('success'):
                return {
                    'success': False,
                    'message': probe.get('message', 'Listings API 探测失败'),
                    'seller_id': acc.get('seller_id'),
                    'marketplace': acc.get('marketplace_name'),
                    'token_ok': True,
                }

            return {
                'success': True,
                'message': f"连接成功! 账号: {acc.get('name')}",
                'seller_id': acc.get('seller_id'),
                'marketplace': acc.get('marketplace_name'),
                'probe_status': probe.get('status_code'),
            }
        except AmazonTokenError as e:
            return {'success': False, 'message': f'凭证错误: {e}'}
        except AmazonNetworkError as e:
            return {'success': False, 'message': f'网络错误: {e}'}
        except Exception as e:
            return {'success': False, 'message': f'连接失败: {e}'}

    def get_auth(self, seller_id: str = None) -> Optional[AmazonAuth]:
        """获取Auth实例"""
        acc = self.get_account(seller_id) if seller_id else self.get_default_account()
        if not acc:
            return None

        return AmazonAuth(
            client_id=acc['lwa_client_id'],
            client_secret=acc['lwa_client_secret'],
            refresh_token=acc['refresh_token'],
        )

    @staticmethod
    def get_marketplace_options() -> List[Dict]:
        """获取可选的Marketplace列表"""
        options = []
        names = {
            'US': '🇺🇸 美国', 'CA': '🇨🇦 加拿大', 'MX': '🇲🇽 墨西哥',
            'UK': '🇬🇧 英国', 'DE': '🇩🇪 德国', 'FR': '🇫🇷 法国',
            'IT': '🇮🇹 意大利', 'ES': '🇪🇸 西班牙',
            'JP': '🇯🇵 日本', 'AU': '🇦🇺 澳大利亚',
        }
        for code, mp_id in MARKETPLACE_IDS.items():
            options.append({
                'code': code,
                'name': names.get(code, code),
                'marketplace_id': mp_id,
                'region': MARKETPLACE_REGION.get(mp_id, 'NA'),
                'endpoint': ENDPOINTS.get(MARKETPLACE_REGION.get(mp_id, 'NA')),
            })
        return options
