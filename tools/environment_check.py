"""
亚马逊 2.8 环境检测脚本

设计目标：
- 源码部署时，提前发现 Python、依赖、端口、目录权限等问题。
- 发行版运行时，配合 `AmazonListingTool --env-check` 做客户电脑上的一键诊断。
- 尽量只使用标准库，避免检测脚本本身被第三方依赖卡住。
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import platform
import socket
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
from urllib.parse import urlparse


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_PORT = 5000
MIN_PYTHON = (3, 10)
RECOMMENDED_PYTHON = (3, 11)

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
    ("PyPI", "https://pypi.org/simple/"),
    ("Amazon SP-API US", "https://sellingpartnerapi-na.amazon.com/"),
)


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


def _is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def _runtime_root() -> Path:
    if _is_frozen_app():
        return Path(sys.executable).resolve().parent
    return PROJECT_ROOT


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
            if name:
                packages.append(name)
    return packages


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


def _check_dependencies(report: CheckReport, require_source_dependencies: bool | None) -> None:
    if require_source_dependencies is None:
        require_source_dependencies = not _is_frozen_app()
    if not require_source_dependencies:
        report.add("Python 依赖", "OK", "发行版已内置运行依赖")
        return

    packages = _iter_requirements((PROJECT_ROOT / "requirements.txt",))
    missing: list[str] = []
    for package in packages:
        import_name = REQUIRED_IMPORTS.get(package, package.replace("-", "_"))
        if importlib.util.find_spec(import_name) is None:
            missing.append(f"{package}({import_name})")
    if missing:
        report.add("Python 依赖", "ERROR", "依赖未安装完整：" + ", ".join(missing), "运行 python tools/bootstrap.py 可自动安装")
    else:
        report.add("Python 依赖", "OK", "requirements.txt 依赖均可导入")


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
    if _looks_like_placeholder(ai_text_key) or _looks_like_placeholder(ai_image_key):
        report.add("AI 配置", "WARN", "AI Key 还未配置完整，AI 文案/图片功能会不可用")
    else:
        report.add("AI 配置", "OK", "AI 文案/图片 Key 已填写")


def _check_port(report: CheckReport, port: int) -> None:
    ok, detail = _can_bind_local_port(port)
    if ok:
        report.add("Web 端口", "OK", f"127.0.0.1:{port} 可用")
    else:
        report.add("Web 端口", "WARN", f"127.0.0.1:{port} 已被占用，发行版启动时会自动尝试下一个端口", detail)


def _check_network(report: CheckReport, strict_network: bool) -> None:
    for name, url in OPTIONAL_NETWORK_TARGETS:
        parsed = urlparse(url)
        host = parsed.hostname or ""
        port = parsed.port or (443 if parsed.scheme == "https" else 80)
        start = time.monotonic()
        ok, detail = _tcp_connect(host, port, timeout=2.5)
        elapsed_ms = int((time.monotonic() - start) * 1000)
        if ok:
            report.add(f"网络：{name}", "OK", f"可连接 {host}:{port} ({elapsed_ms}ms)")
        else:
            status = "ERROR" if strict_network else "WARN"
            report.add(f"网络：{name}", status, f"无法连接 {host}:{port}", detail)


def run_checks(
    *,
    require_source_dependencies: bool | None = None,
    check_network: bool = False,
    strict_network: bool = False,
    port: int | None = None,
) -> CheckReport:
    report = CheckReport()
    system = platform.system() or "Unknown"
    machine = platform.machine() or "unknown"
    report.add("操作系统", "OK", f"{system} {platform.release()} ({machine})")
    if system not in {"Darwin", "Windows", "Linux"}:
        report.add("系统支持", "WARN", "当前系统不在常规测试范围内，推荐 macOS 或 Windows")

    _check_python(report)
    _check_dependencies(report, require_source_dependencies)
    _check_paths(report)
    _check_config(report)
    _check_port(report, _read_port(port))
    if check_network or strict_network:
        _check_network(report, strict_network=strict_network)
    else:
        report.add("网络检测", "OK", "已跳过外网连通性检测；需要时运行 --network")
    return report


def _print_report(report: CheckReport, quiet: bool = False) -> None:
    if not quiet:
        print("亚马逊 2.8 环境检测")
        print("=" * 28)
    icon = {"OK": "[OK]", "WARN": "[WARN]", "ERROR": "[ERROR]"}
    for item in report.items:
        print(f"{icon.get(item.status, '[INFO]')} {item.name}: {item.message}")
        if item.detail and not quiet:
            print(f"       {item.detail}")
    print("-" * 28)
    if report.ok:
        if report.warnings:
            print(f"检测完成：可启动，但有 {len(report.warnings)} 个提醒需要留意。")
        else:
            print("检测完成：环境正常，可以启动。")
    else:
        print(f"检测失败：发现 {len(report.errors)} 个必须处理的问题。")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="亚马逊 2.8 环境检测")
    parser.add_argument("--json", action="store_true", help="输出 JSON，方便自动化读取")
    parser.add_argument("--quiet", action="store_true", help="减少说明性输出")
    parser.add_argument("--network", action="store_true", help="检查 PyPI 和 Amazon SP-API TCP 连通性")
    parser.add_argument("--strict-network", action="store_true", help="网络不通时按错误处理")
    parser.add_argument("--source-deps", action="store_true", help="强制检查源码运行依赖")
    parser.add_argument("--no-source-deps", action="store_true", help="跳过源码依赖导入检查")
    parser.add_argument("--port", type=int, default=None, help="指定要检查的 Web 端口")
    args = parser.parse_args(argv)

    require_source_dependencies = None
    if args.source_deps:
        require_source_dependencies = True
    elif args.no_source_deps:
        require_source_dependencies = False

    report = run_checks(
        require_source_dependencies=require_source_dependencies,
        check_network=args.network,
        strict_network=args.strict_network,
        port=args.port,
    )
    if args.json:
        print(json.dumps(report.as_dict(), ensure_ascii=False, indent=2))
    else:
        _print_report(report, quiet=args.quiet)
    return 0 if report.ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
