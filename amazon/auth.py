"""
Amazon SP-API OAuth2 认证模块
处理 Login with Amazon (LWA) OAuth2 token 获取和刷新
"""
import time
import logging
import requests
from typing import Optional, Dict

logger = logging.getLogger(__name__)

LWA_TOKEN_URL = "https://api.amazon.com/auth/o2/token"


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

    def get_access_token(self) -> str:
        """
        获取有效的 Access Token (自动刷新)

        Returns:
            有效的access_token
        """
        if self._access_token and time.time() < self._token_expiry - 60:
            return self._access_token

        logger.info("🔑 刷新 Amazon Access Token...")
        response = requests.post(LWA_TOKEN_URL, data={
            'grant_type': 'refresh_token',
            'refresh_token': self.refresh_token,
            'client_id': self.client_id,
            'client_secret': self.client_secret,
        })

        if response.status_code != 200:
            raise Exception(f"Token刷新失败: {response.status_code} {response.text}")

        data = response.json()
        self._access_token = data['access_token']
        self._token_expiry = time.time() + data.get('expires_in', 3600)

        logger.info(f"✅ Token获取成功，有效期 {data.get('expires_in', 3600)}秒")
        return self._access_token

    def get_headers(self) -> Dict[str, str]:
        """获取API请求头"""
        return {
            'x-amz-access-token': self.get_access_token(),
            'Content-Type': 'application/json',
            'Accept': 'application/json',
        }
