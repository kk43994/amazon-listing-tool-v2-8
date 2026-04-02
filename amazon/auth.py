"""
Amazon SP-API OAuth2 认证模块
处理 Login with Amazon (LWA) OAuth2 token 获取和刷新
"""
import threading
import time
import logging
import requests
from typing import Optional, Dict

logger = logging.getLogger(__name__)

LWA_TOKEN_URL = "https://api.amazon.com/auth/o2/token"
LWA_TIMEOUT_SECONDS = 20


class AmazonAuthError(Exception):
    """Amazon 认证相关错误基类"""


class AmazonTokenError(AmazonAuthError):
    """Token 获取/刷新失败（凭证无效、权限不足等）"""


class AmazonNetworkError(AmazonAuthError):
    """网络层错误（超时、连接失败等）"""


class AmazonAuth:
    """Amazon SP-API OAuth2 认证"""

    def __init__(self, client_id: str, client_secret: str, refresh_token: str):
        """
        初始化认证

        Args:
            client_id: LWA应用Client ID
            client_secret: LWA应用Client Secret
            refresh_token: 卖家授权的Refresh Token
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.refresh_token = refresh_token
        self._access_token: Optional[str] = None
        self._token_expiry: float = 0
        self._token_lock = threading.Lock()

    def get_access_token(self) -> str:
        """
        获取有效的 Access Token (自动刷新，线程安全)

        Returns:
            有效的access_token
        """
        if self._access_token and time.time() < self._token_expiry - 60:
            return self._access_token

        with self._token_lock:
            if self._access_token and time.time() < self._token_expiry - 60:
                return self._access_token

            logger.info("🔑 刷新 Amazon Access Token...")
            try:
                response = requests.post(
                    LWA_TOKEN_URL,
                    data={
                        'grant_type': 'refresh_token',
                        'refresh_token': self.refresh_token,
                        'client_id': self.client_id,
                        'client_secret': self.client_secret,
                    },
                    timeout=LWA_TIMEOUT_SECONDS,
                )
            except requests.exceptions.Timeout as exc:
                raise AmazonNetworkError(f"Token刷新超时，请稍后重试: {exc}") from exc
            except requests.exceptions.RequestException as exc:
                raise AmazonNetworkError(f"Token刷新请求失败: {exc}") from exc

            if response.status_code == 401:
                raise AmazonTokenError(f"凭证无效(401): {response.text}")
            if response.status_code == 403:
                raise AmazonTokenError(f"权限不足(403): {response.text}")
            if response.status_code != 200:
                raise AmazonTokenError(f"Token刷新失败: {response.status_code} {response.text}")

            data = response.json()
            self._access_token = data['access_token']
            self._token_expiry = time.time() + data.get('expires_in', 3600)

            logger.info(f"✅ Token获取成功，有效期 {data.get('expires_in', 3600)}秒")
            return self._access_token

    def force_refresh(self) -> str:
        """强制刷新 Token（忽略缓存）"""
        with self._token_lock:
            self._access_token = None
            self._token_expiry = 0
        return self.get_access_token()

    def get_headers(self) -> Dict[str, str]:
        """获取API请求头"""
        return {
            'x-amz-access-token': self.get_access_token(),
            'Content-Type': 'application/json',
            'Accept': 'application/json',
        }
