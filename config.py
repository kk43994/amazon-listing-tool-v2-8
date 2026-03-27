"""
亚马逊商品处理工具 - 配置文件
"""
import os
import re
from dotenv import load_dotenv

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

_config_instance = None


def _normalize_base_url(value: str, default: str) -> str:
    text = str(value or "").strip()
    if not text:
        return default
    if not re.match(r'^https?://', text, re.IGNORECASE):
        text = f'https://{text.lstrip("/")}'
    return text.rstrip('/')


def _infer_protocol(endpoint_template: str, default: str) -> str:
    template = str(endpoint_template or "").strip().lower()
    if "generatecontent" in template:
        return "gemini_generate_content"
    if "/chat/completions" in template:
        return "openai_chat_completions"
    if "/images/" in template:
        return "openai_images"
    return default


def _resolve_dir_setting(value: str, default_name: str) -> str:
    text = str(value or "").strip()
    if not text:
        return os.path.join(BASE_DIR, default_name)
    if os.path.isabs(text):
        return text
    return os.path.abspath(os.path.join(BASE_DIR, text))


def _read_int_env(name: str, default: int, minimum: int = 1) -> int:
    raw = str(os.getenv(name, str(default)) or "").strip()
    try:
        value = int(raw)
    except (TypeError, ValueError):
        value = default
    return max(minimum, value)


def _read_bool_env(name: str, default: bool = False) -> bool:
    raw = str(os.getenv(name, "1" if default else "0") or "").strip().lower()
    return raw in ("1", "true", "yes", "on")

class Config:
    def __init__(self):
        self.reload()

    def reload(self):
        """重新加载环境变量"""
        # 进程级环境变量优先，.env 作为默认值来源。
        load_dotenv(override=False)

        # AI Provider 配置
        raw_ai_api_key = os.getenv('AI_API_KEY', '').strip()
        raw_ai_api_base = os.getenv('AI_API_BASE', '').strip()

        self.AI_API_KEY = raw_ai_api_key
        self.AI_API_BASE = _normalize_base_url(raw_ai_api_base, 'https://api.openai.com/v1')
        self.AI_TEXT_MODEL = os.getenv('AI_TEXT_MODEL', 'gpt-4o')
        self.AI_IMAGE_MODEL = os.getenv('AI_IMAGE_MODEL', 'gpt-4o')

        self.AI_TEXT_API_KEY = os.getenv('AI_TEXT_API_KEY', raw_ai_api_key).strip()
        self.AI_TEXT_API_BASE = _normalize_base_url(
            os.getenv('AI_TEXT_API_BASE', '').strip() or raw_ai_api_base,
            'https://api.openai.com/v1',
        )
        self.AI_TEXT_ENDPOINT_TEMPLATE = os.getenv('AI_TEXT_ENDPOINT_TEMPLATE', '/chat/completions').strip()
        self.AI_TEXT_PROTOCOL = os.getenv(
            'AI_TEXT_PROTOCOL',
            _infer_protocol(self.AI_TEXT_ENDPOINT_TEMPLATE, 'openai_chat_completions'),
        ).strip()

        self.AI_IMAGE_API_KEY = os.getenv('AI_IMAGE_API_KEY', raw_ai_api_key).strip()
        self.AI_IMAGE_API_BASE = _normalize_base_url(
            os.getenv('AI_IMAGE_API_BASE', '').strip() or raw_ai_api_base,
            'https://api.openai.com/v1',
        )
        self.AI_IMAGE_ENDPOINT_TEMPLATE = os.getenv('AI_IMAGE_ENDPOINT_TEMPLATE', '').strip()
        self.AI_IMAGE_PROTOCOL = os.getenv(
            'AI_IMAGE_PROTOCOL',
            _infer_protocol(self.AI_IMAGE_ENDPOINT_TEMPLATE, 'openai_images'),
        ).strip()

        if not self.AI_API_KEY:
            self.AI_API_KEY = self.AI_TEXT_API_KEY or self.AI_IMAGE_API_KEY

        # Amazon SP-API 配置
        self.AMAZON_CLIENT_ID = os.getenv('AMAZON_CLIENT_ID', '')
        self.AMAZON_CLIENT_SECRET = os.getenv('AMAZON_CLIENT_SECRET', '')
        self.AMAZON_REFRESH_TOKEN = os.getenv('AMAZON_REFRESH_TOKEN', '')
        self.AMAZON_MARKETPLACE = os.getenv('AMAZON_MARKETPLACE', 'ATVPDKIKX0DER')
        self.AMAZON_SELLER_ID = os.getenv('AMAZON_SELLER_ID', '')

        # 路径
        self.INPUT_DIR = _resolve_dir_setting(os.getenv('INPUT_DIR', ''), 'input')
        self.OUTPUT_DIR = _resolve_dir_setting(os.getenv('OUTPUT_DIR', ''), 'output')
        self.LOGS_DIR = _resolve_dir_setting(os.getenv('LOGS_DIR', ''), 'logs')
        self.OUTPUT_IMAGE_PUBLIC_BASE = os.getenv('OUTPUT_IMAGE_PUBLIC_BASE', '').strip().rstrip('/')
        self.MEDIA_STORE_ENABLED = _read_bool_env('MEDIA_STORE_ENABLED', False)
        self.MEDIA_STORE_PROVIDER = os.getenv('MEDIA_STORE_PROVIDER', '').strip().lower()
        self.MEDIA_S3_BUCKET = os.getenv('MEDIA_S3_BUCKET', '').strip()
        self.MEDIA_S3_REGION = os.getenv('MEDIA_S3_REGION', '').strip() or 'us-east-1'
        self.MEDIA_S3_PREFIX = os.getenv('MEDIA_S3_PREFIX', 'amazon28').strip().strip('/')
        self.MEDIA_S3_SUBMIT_SCHEME = os.getenv('MEDIA_S3_SUBMIT_SCHEME', 's3').strip().lower() or 's3'
        self.MEDIA_S3_PREVIEW_BASE = os.getenv('MEDIA_S3_PREVIEW_BASE', '').strip().rstrip('/')
        self.DEFAULT_LANG = os.getenv('DEFAULT_LANG', 'zh').strip().lower() or 'zh'
        if self.DEFAULT_LANG not in ('zh', 'en'):
            self.DEFAULT_LANG = 'zh'
        self.BATCH_LIMIT = _read_int_env('BATCH_LIMIT', 500, minimum=1)
        self.AI_CONCURRENCY = _read_int_env('AI_CONCURRENCY', 3, minimum=1)
        self.IMAGE_CONCURRENCY = _read_int_env('IMAGE_CONCURRENCY', 2, minimum=1)

        # 图片处理
        self.IMAGE_MAX_SIZE = 2000
        self.IMAGE_QUALITY = 95
        self.IMAGE_BG_STYLE = 'white'

        # 超时
        self.OPENAI_TIMEOUT = int(os.getenv('OPENAI_TIMEOUT', '120'))
        self.OPENAI_MAX_RETRIES = int(os.getenv('OPENAI_MAX_RETRIES', '3'))
        self.WEB_DEBUG = os.getenv('WEB_DEBUG', '').strip().lower() in ('1', 'true', 'yes', 'on')
        self.WEB_PORT = _read_int_env('WEB_PORT', 5000, minimum=1)

def get_config():
    global _config_instance
    if _config_instance is None:
        _config_instance = Config()
    return _config_instance

def reload_config():
    """强制重新加载配置"""
    global _config_instance
    if _config_instance is None:
        _config_instance = Config()
    else:
        _config_instance.reload()
    return _config_instance
