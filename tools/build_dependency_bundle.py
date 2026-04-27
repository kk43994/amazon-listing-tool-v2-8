"""构建离线依赖包，供没有开发经验的新电脑源码部署使用。"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
RELEASE_DIR = PROJECT_ROOT / "release"


def _platform_suffix() -> str:
    system = platform.system().lower()
    machine = platform.machine().lower()
    return f"{system}-{machine}"


def _run(cmd: list[str]) -> None:
    print("> " + " ".join(cmd))
    subprocess.run(cmd, cwd=PROJECT_ROOT, check=True)


def _copy_if_exists(src: Path, dest: Path) -> None:
    if src.exists():
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)


def _write_install_scripts(bundle_dir: Path) -> None:
    (bundle_dir / "离线安装依赖.command").write_text(
        "#!/bin/bash\n"
        "set -e\n"
        "DIR=\"$(cd \"$(dirname \"$0\")\" && pwd)\"\n"
        "cd \"$DIR\"\n"
        "cd ..\n"
        "python3 tools/bootstrap.py\n",
        encoding="utf-8",
    )
    (bundle_dir / "离线安装依赖.command").chmod(0o755)
    (bundle_dir / "离线安装依赖.bat").write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        "cd /d %~dp0\\..\r\n"
        "py -3 tools\\bootstrap.py\r\n"
        "pause\r\n",
        encoding="utf-8",
    )


def _write_readme(bundle_dir: Path) -> None:
    (bundle_dir / "依赖包说明.txt").write_text(
        "亚马逊 2.8 离线依赖包\n\n"
        "用途：给售后/技术人员在源码部署时离线安装 Python 依赖。\n\n"
        "客户交付不要让客户安装依赖，直接发送 GitHub Release 里的 AmazonListingTool-*.zip。\n\n"
        "源码离线部署步骤：\n"
        "1. 把本压缩包解压到项目根目录，确保出现 vendor/wheelhouse/。\n"
        "2. 新电脑需先安装 Python 3.10+，推荐 3.11。\n"
        "3. 运行 python tools/bootstrap.py 或双击本包里的离线安装脚本。\n"
        "4. 脚本会创建 .venv、安装依赖、生成 .env，并运行环境检测。\n",
        encoding="utf-8",
    )


def _remove_macos_metadata(target_dir: Path) -> None:
    for path in target_dir.rglob(".DS_Store"):
        path.unlink(missing_ok=True)


def main() -> int:
    RELEASE_DIR.mkdir(exist_ok=True)
    package_name = f"AmazonListingTool-dependencies-{_platform_suffix()}"
    bundle_dir = RELEASE_DIR / package_name
    wheelhouse = bundle_dir / "vendor" / "wheelhouse"
    shutil.rmtree(bundle_dir, ignore_errors=True)
    wheelhouse.mkdir(parents=True, exist_ok=True)

    _run([
        sys.executable,
        "-m",
        "pip",
        "download",
        "--dest",
        str(wheelhouse),
        "-r",
        "requirements.txt",
        "-r",
        "requirements-release.txt",
    ])

    _copy_if_exists(PROJECT_ROOT / "requirements.txt", bundle_dir / "requirements.txt")
    _copy_if_exists(PROJECT_ROOT / "requirements-release.txt", bundle_dir / "requirements-release.txt")
    _copy_if_exists(PROJECT_ROOT / "tools" / "bootstrap.py", bundle_dir / "tools" / "bootstrap.py")
    _copy_if_exists(PROJECT_ROOT / "tools" / "environment_check.py", bundle_dir / "tools" / "environment_check.py")
    _write_install_scripts(bundle_dir)
    _write_readme(bundle_dir)
    _remove_macos_metadata(bundle_dir)

    archive = shutil.make_archive(str(RELEASE_DIR / package_name), "zip", root_dir=RELEASE_DIR, base_dir=package_name)
    print(f"依赖包目录: {bundle_dir}")
    print(f"依赖包压缩包: {archive}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
