"""
亚马逊 2.8 环境检测脚本

设计目标：
- 源码部署时，提前发现 Python、依赖、端口、目录权限等问题。
- 发行版运行时，配合 `AmazonListingTool --env-check` 做客户电脑上的一键诊断。
- 尽量只使用标准库，避免检测脚本本身被第三方依赖卡住。
"""
from __future__ import annotations

import argparse
import http.client
import importlib.util
import json
import os
import platform
import shutil
import socket
import ssl
import sys
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable
from urllib.parse import urlparse


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PORT = 5000
MIN_PYTHON = (3, 10)
RECOMMENDED_PYTHON = (3, 11)
RECOMMENDED_AI_BASE = "https://api.kk666.best"
RECOMMENDED_AI_ENDPOINT = "/v1beta/models/{model}:generateContent"
RECOMMENDED_TEXT_MODEL = "gemini-3.1-flash-lite-preview"
RECOMMENDED_IMAGE_MODEL = "gemini-3.1-flash-image-preview"

# requirements.txt 中包名和 import 名不完全一致，统一在这里做映射。
REQUIRED_IMPORTS = {
    "openai": "openai",
    "openpyxl": "openpyxl",
    "Pillow": "PIL",
    "python-dotenv": "dotenv",
    "requests": "requests",
    "httpx": "httpx",
    "flask": "flask",
    "xlrd": "xlrd",
    "xlwt": "xlwt",
    "boto3": "boto3",
    "python-amazon-sp-api": "sp_api",
}

OPTIONAL_NETWORK_TARGETS = (
    ("AI 中转域名", RECOMMENDED_AI_BASE),
    ("Amazon LWA 登录授权", "https://api.amazon.com/auth/o2/token"),
    ("Amazon SP-API 北美", "https://sellingpartnerapi-na.amazon.com/"),
    ("Amazon SP-API 欧洲", "https://sellingpartnerapi-eu.amazon.com/"),
    ("Amazon SP-API 远东", "https://sellingpartnerapi-fe.amazon.com/"),
)


def _configure_output_encoding() -> None:
    """Windows CI/客户终端可能不是 UTF-8，尽量避免中文输出触发编码异常。"""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


@dataclass
class CheckItem:
    name: str
    status: str
    message: str
    detail: str = ""

    def as_dict(self) -> dict:
        payload = {"name": self.name, "status": self.status, "message": self.message}
        if self.detail:
            payload["detail"] = self.detail
        return payload


@dataclass
class CheckReport:
    items: list[CheckItem] = field(default_factory=list)

    def add(self, name: str, status: str, message: str, detail: str = "") -> None:
        self.items.append(CheckItem(name=name, status=status, message=message, detail=detail))

    @property
    def errors(self) -> list[CheckItem]:
        return [item for item in self.items if item.status == "ERROR"]

    @property
    def warnings(self) -> list[CheckItem]:
        return [item for item in self.items if item.status == "WARN"]

    @property
    def ok(self) -> bool:
        return not self.errors

    def as_dict(self) -> dict:
        return {
            "ok": self.ok,
            "errors": len(self.errors),
            "warnings": len(self.warnings),
            "items": [item.as_dict() for item in self.items],
        }


def _normalize_requirement_name(name: str) -> str:
    return str(name or "").strip().lower().replace("_", "-")


def _is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def _runtime_root() -> Path:
    if _is_frozen_app():
        return Path(sys.executable).resolve().parent
    return PROJECT_ROOT


def _ensure_basic_runtime_files(report: CheckReport) -> None:
    """Create the files/folders customers usually forget before running checks."""
    root = _runtime_root()
    fixed: list[str] = []

    env_path = root / ".env"
    if not env_path.exists():
        example = root / ".env.example"
        if not example.exists() and not _is_frozen_app():
            example = PROJECT_ROOT / ".env.example"
        if example.exists():
            shutil.copyfile(example, env_path)
            fixed.append(".env")

    accounts_path = root / "accounts.json"
    if not accounts_path.exists():
        accounts_path.write_text(json.dumps({"accounts": []}, ensure_ascii=False, indent=2), encoding="utf-8")
        fixed.append("accounts.json")

    for dirname in ("input", "output", "logs", "config"):
        path = root / dirname
        if not path.exists():
            path.mkdir(parents=True, exist_ok=True)
            fixed.append(f"{dirname}/")

    if fixed:
        report.add("自动修复", "OK", "已自动补齐基础运行文件：" + ", ".join(fixed))
    else:
        report.add("自动修复", "OK", "基础运行文件已存在，无需修复")


def _read_env_file(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.exists():
        return values
    try:
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip('"').strip("'")
    except OSError:
        return values
    return values


def _env_value(name: str, default: str = "") -> str:
    env_path = _runtime_root() / ".env"
    file_values = _read_env_file(env_path)
    return str(os.getenv(name) or file_values.get(name) or default).strip()


def _write_env_file_values(path: Path, updates: dict[str, str]) -> None:
    lines = path.read_text(encoding="utf-8").splitlines(keepends=True) if path.exists() else []
    updated: set[str] = set()
    output: list[str] = []
    for line in lines:
        key = line.split("=", 1)[0].strip() if "=" in line else ""
        if key in updates:
            output.append(f"{key}={updates[key]}\n")
            updated.add(key)
        else:
            output.append(line)
    for key, value in updates.items():
        if key not in updated:
            output.append(f"{key}={value}\n")
    path.write_text("".join(output), encoding="utf-8")


def _normalize_ai_base(value: str) -> str:
    text = str(value or "").strip()
    if not text:
        return ""
    if not text.lower().startswith(("http://", "https://")):
        text = "https://" + text.lstrip("/")
    text = text.rstrip("/")
    if text.lower().endswith("/v1") or text.lower().endswith("/v1beta"):
        text = text.rsplit("/", 1)[0]
    return text


def _looks_like_bad_ai_base(value: str) -> bool:
    text = _normalize_ai_base(value).lower()
    return (not text) or "api.example.com" in text or "api.kk666.online" in text


def _repair_common_config(report: CheckReport) -> None:
    """Fix config typos that repeatedly break non-technical customers."""
    root = _runtime_root()
    env_path = root / ".env"
    if not env_path.exists():
        return

    values = _read_env_file(env_path)
    updates: dict[str, str] = {}
    for key in ("AI_TEXT_API_BASE", "AI_IMAGE_API_BASE"):
        current = values.get(key) or values.get("AI_API_BASE", "")
        normalized = _normalize_ai_base(current)
        if _looks_like_bad_ai_base(current) or normalized != current:
            updates[key] = RECOMMENDED_AI_BASE if _looks_like_bad_ai_base(current) else normalized

    for key in ("AI_TEXT_ENDPOINT_TEMPLATE", "AI_IMAGE_ENDPOINT_TEMPLATE"):
        endpoint = str(values.get(key, "") or "").strip()
        endpoint_lower = endpoint.lower()
        if not endpoint or ("generatecontent" in endpoint_lower and "{model}" not in endpoint):
            updates[key] = RECOMMENDED_AI_ENDPOINT

    if updates.get("AI_TEXT_ENDPOINT_TEMPLATE") == RECOMMENDED_AI_ENDPOINT or "AI_TEXT_API_BASE" in updates:
        updates.setdefault("AI_TEXT_PROTOCOL", "gemini_generate_content")
        updates.setdefault("AI_TEXT_MODEL", values.get("AI_TEXT_MODEL") or RECOMMENDED_TEXT_MODEL)
    if updates.get("AI_IMAGE_ENDPOINT_TEMPLATE") == RECOMMENDED_AI_ENDPOINT or "AI_IMAGE_API_BASE" in updates:
        updates.setdefault("AI_IMAGE_PROTOCOL", "gemini_generate_content")
        updates.setdefault("AI_IMAGE_MODEL", values.get("AI_IMAGE_MODEL") or RECOMMENDED_IMAGE_MODEL)

    if updates:
        _write_env_file_values(env_path, updates)
        report.add("配置自动修复", "OK", "已修正常见 AI 配置错误：" + ", ".join(sorted(updates)))


def _read_port(explicit_port: int | None = None) -> int:
    if explicit_port:
        return explicit_port
    raw = _env_value("WEB_PORT", str(DEFAULT_PORT))
    try:
        port = int(raw)
    except ValueError:
        return DEFAULT_PORT
    return port if 1 <= port <= 65535 else DEFAULT_PORT


def _tcp_connect(host: str, port: int, timeout: float = 2.0) -> tuple[bool, str]:
    try:
        with socket.create_connection((host, port), timeout=timeout):
            return True, ""
    except OSError as exc:
        return False, str(exc)


def _can_bind_local_port(port: int) -> tuple[bool, str]:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind(("127.0.0.1", port))
        return True, ""
    except OSError as exc:
        return False, str(exc)


def _check_write_dir(path: Path) -> tuple[bool, str]:
    try:
        path.mkdir(parents=True, exist_ok=True)
        with tempfile.NamedTemporaryFile(prefix="amazon28_env_", dir=str(path), delete=True) as handle:
            handle.write(b"ok")
            handle.flush()
        return True, ""
    except OSError as exc:
        return False, str(exc)


def _looks_like_placeholder(value: str) -> bool:
    text = str(value or "").strip().upper()
    return not text or "YOUR_" in text or text.startswith("REPLACE-WITH")


def _load_accounts_status(accounts_path: Path) -> tuple[str, str]:
    if not accounts_path.exists():
        return "WARN", "未找到 accounts.json，首次启动会自动创建账号配置模板"
    try:
        payload = json.loads(accounts_path.read_text(encoding="utf-8"))
        accounts = payload.get("accounts", []) if isinstance(payload, dict) else []
    except Exception as exc:  # noqa: BLE001 - 环境诊断要尽量捕获并转成可读结果。
        return "ERROR", f"accounts.json 无法解析：{exc}"
    if not accounts:
        return "WARN", "accounts.json 中还没有账号，请在设置页添加 Amazon SP-API 账号"
    default = next((acc for acc in accounts if acc.get("is_default") and acc.get("enabled", True)), accounts[0])
    required = ("seller_id", "lwa_client_id", "lwa_client_secret", "refresh_token")
    missing = [key for key in required if _looks_like_placeholder(str(default.get(key, "")))]
    if missing:
        return "WARN", "默认账号凭证未填完整，进入设置页补齐后才能预览/提交：" + ", ".join(missing)
    return "OK", "默认 Amazon 账号配置已填写"


def _iter_requirements(paths: Iterable[Path]) -> list[str]:
    packages: list[str] = []
    seen: set[str] = set()
    for path in paths:
        if not path.exists():
            continue
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line or line.startswith("#") or line.startswith("-"):
                continue
            name = line.split(";", 1)[0].strip()
            for sep in (">=", "==", "~=", "<=", ">", "<"):
                name = name.split(sep, 1)[0].strip()
            key = _normalize_requirement_name(name)
            if name and key not in seen:
                seen.add(key)
                packages.append(name)
    return packages


def _find_wheelhouse() -> Path | None:
    root = _runtime_root()
    candidates = [
        root / "vendor" / "wheelhouse",
        root / "wheelhouse",
        root / "dependency-bundle" / "wheelhouse",
        PROJECT_ROOT / "vendor" / "wheelhouse",
        PROJECT_ROOT / "wheelhouse",
        PROJECT_ROOT / "dependency-bundle" / "wheelhouse",
    ]
    candidates.extend(root.glob("AmazonListingTool-dependencies-*/vendor/wheelhouse"))
    candidates.extend(PROJECT_ROOT.glob("AmazonListingTool-dependencies-*/vendor/wheelhouse"))
    candidates.extend((PROJECT_ROOT / "release").glob("AmazonListingTool-dependencies-*/vendor/wheelhouse"))
    for candidate in candidates:
        if candidate.exists() and any(candidate.glob("*")):
            return candidate
    return None


def _check_wheelhouse(report: CheckReport, packages: list[str]) -> None:
    wheelhouse = _find_wheelhouse()
    if not wheelhouse:
        if _is_frozen_app():
            report.add("离线依赖包", "OK", "客户发行版已内置依赖，不需要额外安装 Python 包")
        else:
            report.add("离线依赖包", "WARN", "未找到 vendor/wheelhouse；源码部署缺依赖时需要联网或解压离线依赖包")
        return

    wheel_names = {_normalize_requirement_name(path.name) for path in wheelhouse.glob("*")}
    missing = []
    for package in packages:
        normalized = _normalize_requirement_name(package)
        normalized_alt = normalized.replace("-", "_")
        if not any(name.startswith(normalized) or name.startswith(normalized_alt) for name in wheel_names):
            missing.append(package)
    if missing:
        report.add("离线依赖包", "WARN", "wheelhouse 存在，但可能缺少：" + ", ".join(missing), str(wheelhouse))
    else:
        report.add("离线依赖包", "OK", f"离线依赖包完整：{wheelhouse}")


def _check_python(report: CheckReport) -> None:
    if _is_frozen_app():
        report.add("运行模式", "OK", "当前是已打包发行版，不需要客户安装 Python 或 pip")
        return

    version = sys.version_info[:3]
    version_text = ".".join(str(part) for part in version)
    if version < MIN_PYTHON:
        report.add("Python", "ERROR", f"Python {version_text} 过低，请安装 Python 3.10+，推荐 3.11")
    elif version[:2] != RECOMMENDED_PYTHON:
        report.add("Python", "WARN", f"当前 Python {version_text} 可运行；推荐使用 Python 3.11 构建发行包")
    else:
        report.add("Python", "OK", f"Python {version_text} 符合推荐版本")


def _check_dependencies(report: CheckReport, require_source_dependencies: bool | None, check_runtime_imports: bool) -> None:
    if require_source_dependencies is None:
        require_source_dependencies = not _is_frozen_app()

    packages = _iter_requirements((_runtime_root() / "requirements.txt", PROJECT_ROOT / "requirements.txt"))
    if not packages:
        packages = list(REQUIRED_IMPORTS)
    _check_wheelhouse(report, packages)

    if not require_source_dependencies and not check_runtime_imports:
        report.add("Python 依赖", "OK", "发行版已内置运行依赖")
        return

    missing: list[str] = []
    for package in packages:
        import_name = REQUIRED_IMPORTS.get(package, package.replace("-", "_"))
        if importlib.util.find_spec(import_name) is None:
            missing.append(f"{package}({import_name})")
    if missing:
        if _is_frozen_app():
            report.add("内置依赖", "ERROR", "发行包内置依赖缺失：" + ", ".join(missing), "请重新下载完整压缩包，或把整个解压后的文件夹发给技术支持")
        else:
            report.add("Python 依赖", "ERROR", "依赖未安装完整：" + ", ".join(missing), "运行 python tools/bootstrap.py 可自动使用离线包/联网安装")
    else:
        report.add("Python 依赖", "OK", "必需依赖均可导入")


def _check_runtime_package(report: CheckReport) -> None:
    root = _runtime_root()
    if _is_frozen_app():
        exe = Path(sys.executable).resolve()
        internal_dir = root / "_internal"
        if not exe.exists():
            report.add("发行包完整性", "ERROR", "主程序不存在，请重新解压完整压缩包", str(exe))
        elif internal_dir.exists():
            report.add("发行包完整性", "OK", "主程序和 _internal 目录存在")
        else:
            report.add("发行包完整性", "WARN", "未找到 _internal 目录；如果是单文件包可忽略，否则请重新解压", str(root))
        return

    required_files = [
        PROJECT_ROOT / "web" / "app.py",
        PROJECT_ROOT / "requirements.txt",
        PROJECT_ROOT / ".env.example",
    ]
    missing = [str(path.relative_to(PROJECT_ROOT)) for path in required_files if not path.exists()]
    if missing:
        report.add("源码完整性", "ERROR", "源码目录缺少关键文件：" + ", ".join(missing))
    else:
        report.add("源码完整性", "OK", "源码关键文件存在")


def _check_runtime_location(report: CheckReport) -> None:
    root = _runtime_root().resolve()
    root_text = str(root)
    lower = root_text.lower()
    risky_markers = [
        (".zip", "看起来可能在压缩包/临时解压路径中运行，请先完整解压到普通文件夹"),
        ("appdata\\local\\temp", "当前在 Windows 临时目录，容易被清理或被安全软件拦截"),
        ("/tmp/", "当前在临时目录，容易被清理"),
        ("\\temp\\", "当前在临时目录，容易被清理"),
    ]
    cloud_markers = ("onedrive", "dropbox", "icloud drive", "google drive", "nutstore", "坚果云")
    if root_text.startswith("\\\\"):
        report.add("运行位置", "WARN", "当前像是在网络共享路径运行，建议复制到本机桌面或 D 盘普通目录", root_text)
    elif any(marker in lower for marker in cloud_markers):
        report.add("运行位置", "WARN", "当前在云同步目录，可能出现文件锁/同步冲突，建议复制到本机普通目录", root_text)
    else:
        risky = next((message for marker, message in risky_markers if marker in lower), "")
        if risky:
            report.add("运行位置", "WARN", risky, root_text)
        else:
            report.add("运行位置", "OK", "运行目录看起来正常", root_text)


def _read_dependency_inventory() -> dict:
    for candidate in (_runtime_root() / "dependency-inventory.json", PROJECT_ROOT / "dependency-inventory.json"):
        if candidate.exists():
            try:
                payload = json.loads(candidate.read_text(encoding="utf-8"))
                if isinstance(payload, dict):
                    payload["_path"] = str(candidate)
                    return payload
            except Exception:
                return {"_path": str(candidate), "_error": "无法解析 dependency-inventory.json"}
    return {}


def _check_dependency_inventory(report: CheckReport) -> None:
    inventory = _read_dependency_inventory()
    if not inventory:
        report.add("依赖清单", "WARN" if _is_frozen_app() else "OK", "未找到 dependency-inventory.json；源码模式可忽略")
        return
    if inventory.get("_error"):
        report.add("依赖清单", "WARN", inventory["_error"], inventory.get("_path", ""))
        return
    requirements = inventory.get("requirements", [])
    if not isinstance(requirements, list):
        report.add("依赖清单", "WARN", "dependency-inventory.json 格式不正确", inventory.get("_path", ""))
        return
    missing: list[str] = []
    for item in requirements:
        if isinstance(item, dict):
            package = item.get("package", "")
            import_name = item.get("import_name") or str(package).replace("-", "_")
        else:
            package = str(item)
            import_name = REQUIRED_IMPORTS.get(package, package.replace("-", "_"))
        if import_name and importlib.util.find_spec(str(import_name)) is None:
            missing.append(f"{package}({import_name})")
    if missing:
        report.add("依赖清单", "ERROR", "清单中的运行依赖无法导入：" + ", ".join(missing), inventory.get("_path", ""))
    else:
        report.add("依赖清单", "OK", f"依赖清单可用，共 {len(requirements)} 项", inventory.get("_path", ""))


def _check_browser(report: CheckReport) -> None:
    try:
        import webbrowser

        controller = webbrowser.get()
        report.add("浏览器", "OK", "已检测到系统默认浏览器", getattr(controller, "name", ""))
    except Exception as exc:
        report.add("浏览器", "WARN", "没有检测到默认浏览器；启动后可手动复制 http://127.0.0.1:端口 打开", str(exc))


def _check_paths(report: CheckReport) -> None:
    root = _runtime_root()
    report.add("运行目录", "OK", str(root))

    dirs = {
        "input": _env_value("INPUT_DIR") or str(root / "input"),
        "output": _env_value("OUTPUT_DIR") or str(root / "output"),
        "logs": _env_value("LOGS_DIR") or str(root / "logs"),
        "config": str(root / "config"),
    }
    failed = []
    for name, raw_path in dirs.items():
        path = Path(raw_path)
        if not path.is_absolute():
            path = root / path
        ok, detail = _check_write_dir(path)
        if not ok:
            failed.append(f"{name}: {path} ({detail})")
    if failed:
        report.add("目录权限", "ERROR", "以下目录不可写：" + "; ".join(failed))
    else:
        report.add("目录权限", "OK", "input/output/logs/config 均可写")


def _check_config(report: CheckReport) -> None:
    root = _runtime_root()
    env_path = root / ".env"
    if env_path.exists():
        report.add(".env", "OK", f"已找到配置文件：{env_path}")
    else:
        report.add(".env", "WARN", "未找到 .env，首次启动会从 .env.example 自动创建")

    status, message = _load_accounts_status(root / "accounts.json")
    report.add("Amazon 账号", status, message)

    ai_text_key = _env_value("AI_TEXT_API_KEY") or _env_value("AI_API_KEY")
    ai_image_key = _env_value("AI_IMAGE_API_KEY") or _env_value("AI_API_KEY")
    text_base = _env_value("AI_TEXT_API_BASE") or _env_value("AI_API_BASE")
    image_base = _env_value("AI_IMAGE_API_BASE") or _env_value("AI_API_BASE")
    if _looks_like_placeholder(ai_text_key) or _looks_like_placeholder(ai_image_key):
        report.add("AI Key", "WARN", "AI Key 还未配置完整，AI 文案/图片功能会不可用")
    else:
        report.add("AI Key", "OK", "AI 文案/图片 Key 已填写")

    bad_base = [
        value for value in (text_base, image_base)
        if _looks_like_bad_ai_base(value)
    ]
    if bad_base:
        report.add("AI 中转域名", "WARN", f"建议把文字/图片 Base URL 都填为 {RECOMMENDED_AI_BASE}")
    else:
        report.add("AI 中转域名", "OK", f"AI Base URL 已填写：文字={text_base or '-'} 图片={image_base or '-'}")


def _check_port(report: CheckReport, port: int) -> None:
    ok, detail = _can_bind_local_port(port)
    if ok:
        report.add("Web 端口", "OK", f"127.0.0.1:{port} 可用")
    else:
        report.add("Web 端口", "WARN", f"127.0.0.1:{port} 已被占用，发行版启动时会自动尝试下一个端口", detail)


def _network_status(strict_network: bool) -> str:
    return "ERROR" if strict_network else "WARN"


def _http_probe(url: str, timeout: float = 3.5) -> tuple[bool, str]:
    parsed = urlparse(url)
    host = parsed.hostname or ""
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    path = parsed.path or "/"
    if parsed.query:
        path = f"{path}?{parsed.query}"
    conn_cls = http.client.HTTPSConnection if parsed.scheme == "https" else http.client.HTTPConnection
    try:
        conn = conn_cls(host, port, timeout=timeout)
        conn.request("HEAD", path, headers={"User-Agent": "AmazonListingTool-Doctor/1.0"})
        response = conn.getresponse()
        response.read(256)
        return True, f"HTTP {response.status} {response.reason}".strip()
    except Exception as exc:
        return False, str(exc)
    finally:
        try:
            conn.close()  # type: ignore[name-defined]
        except Exception:
            pass


def _tls_probe(host: str, port: int, timeout: float = 3.5) -> tuple[bool, str]:
    try:
        context = ssl.create_default_context()
        with socket.create_connection((host, port), timeout=timeout) as raw_sock:
            with context.wrap_socket(raw_sock, server_hostname=host) as tls_sock:
                cert = tls_sock.getpeercert() or {}
                subject = cert.get("subject", ())
                common_name = ""
                for group in subject:
                    for key, value in group:
                        if key == "commonName":
                            common_name = str(value)
                            break
                return True, common_name or tls_sock.version() or "TLS OK"
    except Exception as exc:
        return False, str(exc)


def _check_network(report: CheckReport, strict_network: bool) -> None:
    # 先报告当前代理配置（傻瓜在公司网时常见）
    proxy_http = os.environ.get("HTTP_PROXY") or os.environ.get("http_proxy") or ""
    proxy_https = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy") or ""
    no_proxy = os.environ.get("NO_PROXY") or os.environ.get("no_proxy") or ""
    if proxy_http or proxy_https:
        report.add(
            "网络：代理配置",
            "OK",
            "已配置 HTTP/HTTPS 代理",
            f"HTTP_PROXY={proxy_http or '(空)'}; HTTPS_PROXY={proxy_https or '(空)'}; NO_PROXY={no_proxy or '(空)'}",
        )
    else:
        report.add(
            "网络：代理配置",
            "OK",
            "未配置代理（直连）",
            "公司网络/防火墙拦截外网时，可在 .env 里填 HTTP_PROXY / HTTPS_PROXY",
        )

    network_failed_any = False

    targets = list(OPTIONAL_NETWORK_TARGETS)
    target_urls = {url for _, url in targets}
    for label, env_name in (("当前文字 AI Base URL", "AI_TEXT_API_BASE"), ("当前图片 AI Base URL", "AI_IMAGE_API_BASE")):
        value = _env_value(env_name)
        if value and value not in target_urls:
            targets.append((label, value))
            target_urls.add(value)

    for name, url in targets:
        parsed = urlparse(url)
        host = parsed.hostname or ""
        port = parsed.port or (443 if parsed.scheme == "https" else 80)
        if not host:
            report.add(f"网络：{name}", "WARN", f"地址格式不正确：{url}")
            continue

        try:
            addresses = socket.getaddrinfo(host, port, type=socket.SOCK_STREAM)
            unique_addresses = sorted({item[4][0] for item in addresses})
            report.add(f"网络：{name} DNS", "OK", f"{host} 可解析", ", ".join(unique_addresses[:4]))
        except OSError as exc:
            report.add(f"网络：{name} DNS", _network_status(strict_network), f"{host} 无法解析", str(exc))
            network_failed_any = True
            continue

        start = time.monotonic()
        ok, detail = _tcp_connect(host, port, timeout=2.5)
        elapsed_ms = int((time.monotonic() - start) * 1000)
        if ok:
            report.add(f"网络：{name} TCP", "OK", f"可连接 {host}:{port} ({elapsed_ms}ms)")
        else:
            report.add(f"网络：{name} TCP", _network_status(strict_network), f"无法连接 {host}:{port}", detail)
            network_failed_any = True
            continue

        if parsed.scheme == "https":
            ok, detail = _tls_probe(host, port)
            if ok:
                report.add(f"网络：{name} TLS", "OK", "TLS 握手成功", detail)
            else:
                report.add(f"网络：{name} TLS", _network_status(strict_network), "TLS 握手失败", detail)
                network_failed_any = True
                continue

        ok, detail = _http_probe(url)
        if ok:
            report.add(f"网络：{name} HTTP", "OK", "HTTP 可达", detail)
        else:
            report.add(f"网络：{name} HTTP", _network_status(strict_network), "HTTP 请求失败", detail)
            network_failed_any = True

    # 网络失败时给傻瓜级建议
    if network_failed_any and not (proxy_http or proxy_https):
        report.add(
            "网络：建议",
            "WARN",
            "外网检查有失败，且未配代理",
            "如果是在公司/学校网络里，请联系 IT 拿到代理地址，填到 .env 的 HTTP_PROXY / HTTPS_PROXY；如果在家可以试试 VPN",
        )


def run_checks(
    *,
    require_source_dependencies: bool | None = None,
    check_network: bool = False,
    strict_network: bool = False,
    check_runtime_imports: bool = False,
    fix: bool = False,
    port: int | None = None,
) -> CheckReport:
    report = CheckReport()
    if fix:
        _ensure_basic_runtime_files(report)
        _repair_common_config(report)
    system = platform.system() or "Unknown"
    machine = platform.machine() or "unknown"
    report.add("操作系统", "OK", f"{system} {platform.release()} ({machine})")
    if system not in {"Darwin", "Windows", "Linux"}:
        report.add("系统支持", "WARN", "当前系统不在常规测试范围内，推荐 macOS 或 Windows")

    _check_python(report)
    _check_runtime_package(report)
    _check_runtime_location(report)
    _check_dependencies(report, require_source_dependencies, check_runtime_imports)
    _check_dependency_inventory(report)
    _check_paths(report)
    _check_config(report)
    _check_browser(report)
    _check_port(report, _read_port(port))
    if check_network or strict_network:
        _check_network(report, strict_network=strict_network)
    else:
        report.add("网络检测", "OK", "已跳过外网连通性检测；需要时运行 --network")
    return report


def _format_report_lines(report: CheckReport, quiet: bool = False) -> list[str]:
    lines: list[str] = []
    if not quiet:
        lines.append("亚马逊 2.8 Doctor 一键检测")
        lines.append("=" * 32)
    icon = {"OK": "[OK]", "WARN": "[WARN]", "ERROR": "[ERROR]"}
    for item in report.items:
        lines.append(f"{icon.get(item.status, '[INFO]')} {item.name}: {item.message}")
        if item.detail and not quiet:
            lines.append(f"       {item.detail}")
    lines.append("-" * 32)
    if report.ok:
        if report.warnings:
            lines.append(f"检测完成：可启动，但有 {len(report.warnings)} 个提醒需要留意。")
        else:
            lines.append("检测完成：环境正常，可以启动。")
    else:
        lines.append(f"检测失败：发现 {len(report.errors)} 个必须处理的问题。")
    return lines


def _print_report(report: CheckReport, quiet: bool = False) -> None:
    print("\n".join(_format_report_lines(report, quiet=quiet)))


def _save_report(report: CheckReport) -> Path:
    log_dir = _runtime_root() / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    (log_dir / "doctor-report.json").write_text(json.dumps(report.as_dict(), ensure_ascii=False, indent=2), encoding="utf-8")
    report_path = log_dir / "doctor-report.txt"
    report_path.write_text("\n".join(_format_report_lines(report, quiet=False)) + "\n", encoding="utf-8")
    print(f"报告已保存：{report_path}")
    return report_path


def _mask_secret(value: str) -> str:
    text = str(value or "")
    if not text:
        return ""
    if len(text) <= 8:
        return "***"
    return f"{text[:4]}***{text[-4:]}"


def _is_secret_key(key: str) -> bool:
    lowered = str(key or "").lower()
    return any(token in lowered for token in ("key", "secret", "token", "password", "credential"))


def _masked_env_text(path: Path) -> str:
    if not path.exists():
        return "missing .env\n"
    lines: list[str] = []
    for raw_line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if "=" not in raw_line or raw_line.strip().startswith("#"):
            lines.append(raw_line)
            continue
        key, value = raw_line.split("=", 1)
        lines.append(f"{key}={_mask_secret(value) if _is_secret_key(key) else value}")
    return "\n".join(lines) + "\n"


def _mask_json_secrets(value):
    if isinstance(value, dict):
        return {
            key: (_mask_secret(raw) if _is_secret_key(str(key)) and not isinstance(raw, (dict, list)) else _mask_json_secrets(raw))
            for key, raw in value.items()
        }
    if isinstance(value, list):
        return [_mask_json_secrets(item) for item in value]
    return value


def _safe_read_json(path: Path) -> dict | list | str:
    if not path.exists():
        return {"missing": str(path)}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception as exc:
        return f"无法解析 {path.name}: {exc}"


def create_support_bundle(report: CheckReport | None = None) -> Path:
    """Create a redacted zip that customers can send to support."""
    if report is None:
        report = run_checks(
            require_source_dependencies=False,
            check_network=False,
            strict_network=False,
            check_runtime_imports=_is_frozen_app(),
            fix=True,
        )
    report_path = _save_report(report)
    root = _runtime_root()
    log_dir = root / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    bundle_path = log_dir / f"support-bundle-{stamp}.zip"

    support_info = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "platform": f"{platform.system()} {platform.release()} ({platform.machine()})",
        "python": platform.python_version(),
        "frozen": _is_frozen_app(),
        "runtime_root": str(root),
        "cwd": os.getcwd(),
        "doctor_ok": report.ok,
        "doctor_errors": len(report.errors),
        "doctor_warnings": len(report.warnings),
    }

    with zipfile.ZipFile(bundle_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("support-info.json", json.dumps(support_info, ensure_ascii=False, indent=2))
        zf.writestr("doctor-report.json", json.dumps(report.as_dict(), ensure_ascii=False, indent=2))
        zf.write(report_path, "doctor-report.txt")
        zf.writestr("env.masked.txt", _masked_env_text(root / ".env"))
        accounts_payload = _mask_json_secrets(_safe_read_json(root / "accounts.json"))
        zf.writestr("accounts.masked.json", json.dumps(accounts_payload, ensure_ascii=False, indent=2))

        for candidate_name in ("release-manifest.json", "dependency-inventory.json", "VERSION"):
            candidate = root / candidate_name
            if candidate.exists():
                zf.write(candidate, candidate_name)

        task_history = root / "output" / "task_history.json"
        if task_history.exists():
            zf.write(task_history, "output/task_history.json")

        for log_path in sorted(log_dir.glob("*")):
            if not log_path.is_file() or log_path == bundle_path:
                continue
            if log_path.suffix.lower() not in {".log", ".txt", ".json"}:
                continue
            if log_path.name.startswith("support-bundle-"):
                continue
            zf.write(log_path, f"logs/{log_path.name}")

    print(f"支持包已保存：{bundle_path}")
    return bundle_path


def main(argv: list[str] | None = None) -> int:
    _configure_output_encoding()

    parser = argparse.ArgumentParser(description="亚马逊 2.8 环境检测")
    parser.add_argument("--json", action="store_true", help="输出 JSON，方便自动化读取")
    parser.add_argument("--quiet", action="store_true", help="减少说明性输出")
    parser.add_argument("--network", action="store_true", help="检查 AI 中转、Amazon LWA/SP-API TCP 连通性")
    parser.add_argument("--strict-network", action="store_true", help="网络不通时按错误处理")
    parser.add_argument("--source-deps", action="store_true", help="强制检查源码运行依赖")
    parser.add_argument("--no-source-deps", action="store_true", help="跳过源码依赖导入检查")
    parser.add_argument("--runtime-deps", action="store_true", help="强制导入检查运行依赖，适合发行包 doctor")
    parser.add_argument("--fix", action="store_true", help="自动补齐 .env、accounts.json 和运行目录")
    parser.add_argument("--doctor", action="store_true", help="客户一键检测：自动修复基础文件、检查内置依赖和外网连通性")
    parser.add_argument("--support-bundle", action="store_true", help="导出已脱敏的技术支持包 zip")
    parser.add_argument("--save-report", action="store_true", help="把检测报告保存到 logs/doctor-report.*")
    parser.add_argument("--port", type=int, default=None, help="指定要检查的 Web 端口")
    args = parser.parse_args(argv)

    require_source_dependencies = None
    if args.source_deps:
        require_source_dependencies = True
    elif args.no_source_deps:
        require_source_dependencies = False

    if args.doctor:
        require_source_dependencies = False
        args.network = True
        args.runtime_deps = True
        args.fix = True
        args.save_report = True
    if args.support_bundle:
        require_source_dependencies = False
        args.network = True
        args.runtime_deps = True
        args.fix = True
        args.save_report = True

    report = run_checks(
        require_source_dependencies=require_source_dependencies,
        check_network=args.network,
        strict_network=args.strict_network,
        check_runtime_imports=args.runtime_deps,
        fix=args.fix,
        port=args.port,
    )
    if args.json:
        print(json.dumps(report.as_dict(), ensure_ascii=False, indent=2))
    else:
        _print_report(report, quiet=args.quiet)
    if args.save_report and not args.support_bundle:
        _save_report(report)
    if args.support_bundle:
        create_support_bundle(report)
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
