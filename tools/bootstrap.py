"""
源码部署一键初始化脚本。

客户交付优先使用 PyInstaller 发行包；这个脚本用于内部/售后需要在新电脑上
从源码启动时，把“建虚拟环境 -> 装依赖 -> 生成配置 -> 环境检测”串起来。
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
VENV_DIR = PROJECT_ROOT / ".venv"
WHEELHOUSE_CANDIDATES = (
    PROJECT_ROOT / "vendor" / "wheelhouse",
    PROJECT_ROOT / "wheelhouse",
    PROJECT_ROOT / "dependency-bundle" / "wheelhouse",
)


def _python_in_venv() -> Path:
    if os.name == "nt":
        return VENV_DIR / "Scripts" / "python.exe"
    return VENV_DIR / "bin" / "python"


def _run(cmd: list[str], *, env: dict[str, str] | None = None) -> None:
    print("\n> " + " ".join(cmd))
    subprocess.run(cmd, cwd=PROJECT_ROOT, check=True, env=env)


def _ensure_supported_python() -> None:
    if sys.version_info < (3, 10):
        raise SystemExit("Python 版本过低：请安装 Python 3.10+，推荐 Python 3.11。")
    print(f"Python: {sys.version.split()[0]} ({platform.system()} {platform.machine()})")


def _ensure_venv() -> Path:
    python_path = _python_in_venv()
    if python_path.exists():
        print(f"已找到虚拟环境：{VENV_DIR}")
        return python_path
    _run([sys.executable, "-m", "venv", str(VENV_DIR)])
    return python_path


def _find_wheelhouse() -> Path | None:
    candidates = list(WHEELHOUSE_CANDIDATES)
    candidates.extend(PROJECT_ROOT.glob("AmazonListingTool-dependencies-*/vendor/wheelhouse"))
    for candidate in candidates:
        if candidate.exists() and any(candidate.glob("*")):
            return candidate
    return None


def _install_dependencies(python_path: Path) -> None:
    _run([str(python_path), "-m", "pip", "install", "--upgrade", "pip"])
    wheelhouse = _find_wheelhouse()
    base_cmd = [str(python_path), "-m", "pip", "install"]
    requirements = ["-r", "requirements.txt"]
    if wheelhouse:
        print(f"使用本地离线依赖包：{wheelhouse}")
        _run(base_cmd + ["--no-index", "--find-links", str(wheelhouse)] + requirements)
    else:
        print("未找到本地 wheelhouse，将从 PyPI 在线安装依赖。")
        _run(base_cmd + requirements)


def _copy_default_files() -> None:
    env_path = PROJECT_ROOT / ".env"
    if not env_path.exists() and (PROJECT_ROOT / ".env.example").exists():
        shutil.copyfile(PROJECT_ROOT / ".env.example", env_path)
        print("已生成 .env，请在 Web 设置页或 .env 中补齐 API Key。")

    for dirname in ("input", "output", "logs", "config"):
        (PROJECT_ROOT / dirname).mkdir(exist_ok=True)


def _write_source_launchers(python_path: Path) -> None:
    if os.name == "nt":
        launcher = PROJECT_ROOT / "启动开发版.bat"
        launcher.write_text(
            "@echo off\r\n"
            "chcp 65001 >nul\r\n"
            "cd /d %~dp0\r\n"
            ".venv\\Scripts\\python.exe tools\\environment_check.py --source-deps\r\n"
            "if errorlevel 1 pause & exit /b 1\r\n"
            ".venv\\Scripts\\python.exe web\\app.py\r\n"
            "pause\r\n",
            encoding="utf-8",
        )
    else:
        launcher = PROJECT_ROOT / "启动开发版.command"
        launcher.write_text(
            "#!/bin/bash\n"
            "set -e\n"
            "DIR=\"$(cd \"$(dirname \"$0\")\" && pwd)\"\n"
            "cd \"$DIR\"\n"
            "./.venv/bin/python tools/environment_check.py --source-deps\n"
            "./.venv/bin/python web/app.py\n",
            encoding="utf-8",
        )
        launcher.chmod(0o755)
    print(f"已生成源码启动脚本：{launcher.name}")


def main() -> int:
    print("亚马逊 2.8 源码环境初始化")
    print("=" * 32)
    _ensure_supported_python()
    python_path = _ensure_venv()
    _install_dependencies(python_path)
    _copy_default_files()
    _write_source_launchers(python_path)
    _run([str(python_path), "tools/environment_check.py", "--source-deps"])
    print("\n初始化完成。")
    print("- 源码启动：双击 启动开发版.command / 启动开发版.bat")
    print("- 客户交付：请优先使用 GitHub Release 中的 AmazonListingTool-*.zip")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
