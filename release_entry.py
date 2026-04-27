"""
发行版启动入口

职责：
- 初始化运行目录和默认配置文件
- 启动本地 Web 服务
- 自动打开浏览器，方便非技术用户直接使用
"""
from __future__ import annotations

import argparse
import os
import shutil
import socket
import sys
import threading
import webbrowser

from amazon.accounts import AccountManager
from core.runtime_paths import resource_path, runtime_path


def _configure_output_encoding() -> None:
    """Avoid UnicodeEncodeError when customers launch from a non-UTF-8 console."""
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


def _ensure_runtime_dirs(config) -> None:
    for dirname in (config.INPUT_DIR, config.OUTPUT_DIR, config.LOGS_DIR, runtime_path("config")):
        os.makedirs(dirname, exist_ok=True)


def _ensure_env_file() -> None:
    env_path = runtime_path(".env")
    if os.path.exists(env_path):
        return
    env_example = resource_path(".env.example")
    if os.path.exists(env_example):
        shutil.copyfile(env_example, env_path)


def _ensure_selected_fields_config() -> None:
    target = runtime_path("config", "selected_fields.json")
    if os.path.exists(target):
        return
    os.makedirs(os.path.dirname(target), exist_ok=True)
    source = resource_path("config", "selected_fields.json")
    if os.path.exists(source):
        shutil.copyfile(source, target)


def _ensure_accounts_template() -> None:
    AccountManager(runtime_path("accounts.json"))


def _prepare_runtime_files() -> None:
    _ensure_env_file()
    _ensure_accounts_template()
    _ensure_selected_fields_config()


def _open_browser_when_ready(port: int) -> None:
    url = f"http://127.0.0.1:{port}"
    try:
        webbrowser.open(url)
    except Exception:
        pass


def _is_port_available(port: int) -> bool:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind(("127.0.0.1", port))
        return True
    except OSError:
        return False


def _choose_web_port(preferred_port: int) -> int:
    if _is_port_available(preferred_port):
        return preferred_port
    for offset in range(1, 20):
        candidate = preferred_port + offset
        if candidate <= 65535 and _is_port_available(candidate):
            return candidate
    return preferred_port


def _run_env_check(json_output: bool = False, quiet: bool = False) -> int:
    _prepare_runtime_files()
    from tools.environment_check import main as env_check_main

    args = ["--no-source-deps"]
    if json_output:
        args.append("--json")
    if quiet:
        args.append("--quiet")
    return env_check_main(args)


def main(argv: list[str] | None = None) -> int:
    _configure_output_encoding()

    parser = argparse.ArgumentParser(description="亚马逊 2.8 发行版启动入口")
    parser.add_argument("--env-check", action="store_true", help="只执行环境检测，不启动 Web 服务")
    parser.add_argument("--env-check-json", action="store_true", help="以 JSON 输出环境检测结果")
    parser.add_argument("--quiet", action="store_true", help="减少环境检测输出")
    parser.add_argument("--no-open-browser", action="store_true", help="启动后不自动打开浏览器")
    args = parser.parse_args(argv)

    if args.no_open_browser:
        os.environ["AMAZON28_NO_OPEN_BROWSER"] = "1"

    if args.env_check or args.env_check_json:
        return _run_env_check(json_output=args.env_check_json, quiet=args.quiet)

    _prepare_runtime_files()

    from web.app import app, config

    _ensure_runtime_dirs(config)
    selected_port = _choose_web_port(config.WEB_PORT)
    if selected_port != config.WEB_PORT:
        print(f"⚠️  端口 {config.WEB_PORT} 已被占用，自动改用 {selected_port}")
        config.WEB_PORT = selected_port

    print("🚀 亚马逊 2.8 已启动")
    print(f"   地址: http://127.0.0.1:{config.WEB_PORT}")
    print(f"   数据目录: {runtime_path()}")

    if str(os.getenv("AMAZON28_NO_OPEN_BROWSER", "")).strip().lower() not in {"1", "true", "yes", "on"}:
        threading.Timer(1.2, _open_browser_when_ready, args=(config.WEB_PORT,)).start()

    try:
        from waitress import serve

        serve(app, host="127.0.0.1", port=config.WEB_PORT)
    except Exception:
        app.run(host="127.0.0.1", port=config.WEB_PORT, debug=False, use_reloader=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
