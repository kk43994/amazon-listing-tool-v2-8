"""
媒体存储抽象。

负责把 Stage1 生成的本地图片上传到外部媒体存储，并返回：
- Amazon 提交使用的 media locator（例如 s3://bucket/key）
- Web 侧可选预览地址（例如 CloudFront URL）
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from typing import Dict

from config import get_config


@dataclass
class MediaUploadResult:
    success: bool
    locator: str = ""
    preview_url: str = ""
    provider: str = ""
    bucket: str = ""
    key: str = ""
    etag: str = ""
    error: str = ""

    def to_dict(self) -> Dict[str, str]:
        return asdict(self)


class DisabledMediaStore:
    provider = "disabled"

    def enabled(self) -> bool:
        return False

    def healthcheck(self) -> Dict[str, str]:
        return {
            "success": False,
            "provider": self.provider,
            "message": "媒体存储未启用",
        }

    def upload_image(self, local_path: str, *, sku: str, slot: str, marketplace: str = "US") -> MediaUploadResult:
        return MediaUploadResult(
            success=False,
            provider=self.provider,
            error="媒体存储未启用",
        )


def get_media_store():
    config = get_config()
    if not getattr(config, "MEDIA_STORE_ENABLED", False):
        return DisabledMediaStore()

    provider = str(getattr(config, "MEDIA_STORE_PROVIDER", "") or "").strip().lower()
    if provider == "s3":
        from core.media_providers.s3_store import S3MediaStore

        return S3MediaStore(config)

    return DisabledMediaStore()
