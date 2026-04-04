"""
发行版启动入口

职责：
- 初始化运行目录和默认配置文件
- 启动本地 Web 服务
- 自动打开浏览器，方便非技术用户直接使用
"""
from __future__ import annotations

import os
import shutil
import threading
import webbrowser

from amazon.accounts import AccountManager
from core.runtime_paths import resource_path, runtime_path


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


def _open_browser_when_ready(port: int) -> None:
    url = f"http://127.0.0.1:{port}"
    try:
        webbrowser.open(url)
    except Exception:
        pass


def main() -> None:
    _ensure_env_file()
    _ensure_accounts_template()
    _ensure_selected_fields_config()

    from web.app import app, config

    _ensure_runtime_dirs(config)

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


if __name__ == "__main__":
    main()
