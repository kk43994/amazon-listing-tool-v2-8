"""
Amazon S3 媒体存储 provider。
"""
from __future__ import annotations

import hashlib
import mimetypes
import os
import re
import time
from typing import Dict
from uuid import uuid4

from core.media_store import MediaUploadResult


class S3MediaStore:
    provider = "s3"

    def __init__(self, config):
        self.config = config
        self.bucket = str(getattr(config, "MEDIA_S3_BUCKET", "") or "").strip()
        self.region = str(getattr(config, "MEDIA_S3_REGION", "") or "us-east-1").strip()
        self.prefix = str(getattr(config, "MEDIA_S3_PREFIX", "amazon28") or "amazon28").strip().strip("/")
        self.submit_scheme = str(getattr(config, "MEDIA_S3_SUBMIT_SCHEME", "s3") or "s3").strip().lower()
        self.preview_base = str(getattr(config, "MEDIA_S3_PREVIEW_BASE", "") or "").strip().rstrip("/")
        self._client = None

    def enabled(self) -> bool:
        return bool(self.bucket)

    def _get_client(self):
        if self._client is not None:
            return self._client
        try:
            import boto3
        except ImportError as exc:
            raise RuntimeError("未安装 boto3，无法使用 S3 媒体存储") from exc

        self._client = boto3.client("s3", region_name=self.region)
        return self._client

    def healthcheck(self) -> Dict[str, str]:
        if not self.enabled():
            return {
                "success": False,
                "provider": self.provider,
                "message": "未配置 MEDIA_S3_BUCKET",
            }
        try:
            client = self._get_client()
            client.head_bucket(Bucket=self.bucket)
            return {
                "success": True,
                "provider": self.provider,
                "message": "S3 媒体存储可用",
                "bucket": self.bucket,
                "prefix": self.prefix,
            }
        except Exception as exc:
            return {
                "success": False,
                "provider": self.provider,
                "message": f"S3 媒体存储不可用: {exc}",
                "bucket": self.bucket,
                "prefix": self.prefix,
            }

    def build_key(self, local_path: str, *, sku: str, slot: str, marketplace: str = "US") -> str:
        ext = os.path.splitext(local_path)[1].lower() or ".jpg"
        safe_sku = re.sub(r"[^0-9A-Za-z_-]+", "_", str(sku or "").strip()).strip("_") or "unknown"
        safe_slot = re.sub(r"[^0-9A-Za-z_-]+", "_", str(slot or "main").strip()).strip("_") or "main"
        safe_marketplace = re.sub(r"[^0-9A-Za-z_-]+", "_", str(marketplace or "US").strip()).strip("_") or "US"
        date_part = time.strftime("%Y%m%d")
        digest = self._file_digest(local_path)[:12]
        filename = f"{uuid4().hex[:8]}_{digest}_{safe_slot}{ext}"
        parts = [segment for segment in (self.prefix, safe_marketplace, safe_sku, date_part, filename) if segment]
        return "/".join(parts)

    def build_locator(self, key: str) -> str:
        if self.submit_scheme == "https" and self.preview_base:
            return f"{self.preview_base}/{key}"
        return f"s3://{self.bucket}/{key}"

    def build_preview_url(self, key: str) -> str:
        if not key:
            return ""
        if self.preview_base:
            return f"{self.preview_base}/{key}"
        return ""

    def upload_image(self, local_path: str, *, sku: str, slot: str, marketplace: str = "US") -> MediaUploadResult:
        if not self.enabled():
            return MediaUploadResult(
                success=False,
                provider=self.provider,
                error="未配置 S3 bucket",
            )
        if not os.path.exists(local_path):
            return MediaUploadResult(
                success=False,
                provider=self.provider,
                error=f"本地图片不存在: {local_path}",
            )

        key = self.build_key(local_path, sku=sku, slot=slot, marketplace=marketplace)
        content_type = mimetypes.guess_type(local_path)[0] or "application/octet-stream"
        client = self._get_client()

        extra_args = {
            "ContentType": content_type,
            "CacheControl": "public,max-age=31536000,immutable",
            "Metadata": {
                "sku": str(sku or ""),
                "slot": str(slot or ""),
                "source": "amazon28",
            },
        }

        try:
            with open(local_path, "rb") as fh:
                client.upload_fileobj(fh, self.bucket, key, ExtraArgs=extra_args)
            head = client.head_object(Bucket=self.bucket, Key=key)
            etag = str(head.get("ETag", "") or "").strip('"')
            return MediaUploadResult(
                success=True,
                provider=self.provider,
                locator=self.build_locator(key),
                preview_url=self.build_preview_url(key),
                bucket=self.bucket,
                key=key,
                etag=etag,
            )
        except Exception as exc:
            return MediaUploadResult(
                success=False,
                provider=self.provider,
                bucket=self.bucket,
                key=key,
                error=str(exc),
            )

    @staticmethod
    def _file_digest(local_path: str) -> str:
        hasher = hashlib.sha1()
        with open(local_path, "rb") as fh:
            for chunk in iter(lambda: fh.read(1024 * 1024), b""):
                hasher.update(chunk)
        return hasher.hexdigest()
