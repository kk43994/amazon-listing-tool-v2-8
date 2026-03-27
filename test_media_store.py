import sys
import shutil
from pathlib import Path
from types import SimpleNamespace
from uuid import uuid4

from config import Config
from core.media_providers.s3_store import S3MediaStore


def _make_workspace_temp_dir() -> Path:
    path = Path(__file__).resolve().parent / f'tmp_media_store_{uuid4().hex[:8]}'
    path.mkdir(parents=True, exist_ok=True)
    return path


def test_s3_media_store_builds_locator_and_key():
    tmp_dir = _make_workspace_temp_dir()
    img_path = tmp_dir / 'demo.jpg'
    img_path.write_bytes(b'test-image')

    try:
        cfg = Config()
        cfg.MEDIA_S3_BUCKET = 'demo-bucket'
        cfg.MEDIA_S3_REGION = 'us-east-1'
        cfg.MEDIA_S3_PREFIX = 'amazon28'
        cfg.MEDIA_S3_SUBMIT_SCHEME = 's3'
        cfg.MEDIA_S3_PREVIEW_BASE = 'https://cdn.example.com'

        store = S3MediaStore(cfg)
        key = store.build_key(str(img_path), sku='SKU-1', slot='main', marketplace='US')
        assert key.startswith('amazon28/US/SKU-1/')
        assert key.endswith('_main.jpg')

        locator = store.build_locator(key)
        preview = store.build_preview_url(key)
        assert locator == f's3://demo-bucket/{key}'
        assert preview == f'https://cdn.example.com/{key}'
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def test_s3_media_store_upload_image_with_fake_boto3(monkeypatch):
    tmp_dir = _make_workspace_temp_dir()
    img_path = tmp_dir / 'demo.jpg'
    img_path.write_bytes(b'test-image')
    calls = {}

    class FakeClient:
        def upload_fileobj(self, fh, bucket, key, ExtraArgs=None):
            calls['upload'] = {
                'bucket': bucket,
                'key': key,
                'extra': ExtraArgs,
                'body': fh.read(),
            }

        def head_object(self, Bucket=None, Key=None):
            calls['head'] = {'bucket': Bucket, 'key': Key}
            return {'ETag': '"etag-1"'}

    try:
        cfg = Config()
        cfg.MEDIA_S3_BUCKET = 'demo-bucket'
        cfg.MEDIA_S3_REGION = 'us-east-1'
        cfg.MEDIA_S3_PREFIX = 'amazon28'
        cfg.MEDIA_S3_SUBMIT_SCHEME = 's3'
        cfg.MEDIA_S3_PREVIEW_BASE = ''

        fake_boto3 = SimpleNamespace(client=lambda *args, **kwargs: FakeClient())
        monkeypatch.setitem(sys.modules, 'boto3', fake_boto3)

        store = S3MediaStore(cfg)
        result = store.upload_image(str(img_path), sku='SKU-1', slot='main', marketplace='US')

        assert result.success is True
        assert result.locator.startswith('s3://demo-bucket/amazon28/US/SKU-1/')
        assert result.etag == 'etag-1'
        assert calls['upload']['bucket'] == 'demo-bucket'
        assert calls['head']['bucket'] == 'demo-bucket'
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
