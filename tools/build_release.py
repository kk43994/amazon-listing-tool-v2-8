"""
构建面向终端用户的发行版

输出内容：
- release/<包名>/ 可运行目录
- release/<包名>.zip 压缩包
"""
from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
RELEASE_DIR = PROJECT_ROOT / "release"


def _platform_suffix() -> str:
    system = platform.system().lower()
    machine = platform.machine().lower()
    return f"{system}-{machine}"


def _add_data_arg(src: Path, dest: str) -> str:
    return f"{src}{os.pathsep}{dest}"


def _pyinstaller_cmd() -> list[str]:
    return [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onedir",
        "--name",
        "AmazonListingTool",
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "web" / "templates", "web/templates"),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "config" / "sp_api_fields.json", "config"),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "config" / "selected_fields.json", "config"),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / ".env.example", "."),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "README.md", "."),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "docs" / "商家前端使用说明_飞书图文版.md", "docs"),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "亚马逊商品采集模板_v1.0.xlsx", "."),
        str(PROJECT_ROOT / "release_entry.py"),
    ]


def _write_mac_launcher(target_dir: Path) -> None:
    launcher = target_dir / "启动亚马逊2.8.command"
    executable = "./AmazonListingTool"
    launcher.write_text(
        "#!/bin/bash\n"
        "DIR=\"$(cd \"$(dirname \"$0\")\" && pwd)\"\n"
        "cd \"$DIR\"\n"
        f"{executable}\n",
        encoding="utf-8",
    )
    launcher.chmod(0o755)


def _write_windows_launcher(target_dir: Path) -> None:
    launcher = target_dir / "启动亚马逊2.8.bat"
    launcher.write_text(
        "@echo off\r\n"
        "cd /d %~dp0\r\n"
        "start \"\" AmazonListingTool.exe\r\n",
        encoding="utf-8",
    )


def _write_launcher(target_dir: Path) -> None:
    system = platform.system().lower()
    if system == "windows":
        _write_windows_launcher(target_dir)
    else:
        _write_mac_launcher(target_dir)


def _write_release_readme(target_dir: Path) -> None:
    system = platform.system().lower()
    launcher_name = "启动亚马逊2.8.bat" if system == "windows" else "启动亚马逊2.8.command"
    readme = target_dir / "快速开始.txt"
    readme.write_text(
        "亚马逊 2.8 发行版\n\n"
        f"1. 双击 `{launcher_name}`\n"
        "2. 程序会自动启动本地服务，并尝试打开浏览器\n"
        "3. 第一次使用请先进入设置页配置亚马逊账号\n"
        "4. 运行数据会写到当前目录下的 input/output/logs/.env/accounts.json\n\n"
        "附带文件：\n"
        "- README.md\n"
        "- docs/商家前端使用说明_飞书图文版.md\n"
        "- .env.example\n"
        "- 亚马逊商品采集模板_v1.0.xlsx\n",
        encoding="utf-8",
    )


def main() -> int:
    shutil.rmtree(DIST_DIR, ignore_errors=True)
    shutil.rmtree(BUILD_DIR, ignore_errors=True)
    RELEASE_DIR.mkdir(exist_ok=True)

    subprocess.run(_pyinstaller_cmd(), cwd=PROJECT_ROOT, check=True)

    package_name = f"AmazonListingTool-{_platform_suffix()}"
    package_dir = RELEASE_DIR / package_name
    shutil.rmtree(package_dir, ignore_errors=True)
    shutil.copytree(DIST_DIR / "AmazonListingTool", package_dir)

    _write_launcher(package_dir)
    _write_release_readme(package_dir)

    zip_base = RELEASE_DIR / package_name
    archive = shutil.make_archive(str(zip_base), "zip", root_dir=RELEASE_DIR, base_dir=package_name)

    print(f"发行目录: {package_dir}")
    print(f"压缩包: {archive}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
