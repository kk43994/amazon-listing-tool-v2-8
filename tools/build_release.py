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
        _add_data_arg(PROJECT_ROOT / "docs" / "客户部署说明.md", "docs"),
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
        "echo \"正在检测运行环境...\"\n"
        f"{executable} --env-check --quiet\n"
        "STATUS=$?\n"
        "if [ $STATUS -ne 0 ]; then\n"
        "  echo\n"
        "  echo \"环境检测未通过，请把窗口内容截图发给技术支持。\"\n"
        "  read -n 1 -s -r -p \"按任意键关闭窗口...\"\n"
        "  echo\n"
        "  exit $STATUS\n"
        "fi\n"
        "echo\n"
        f"{executable}\n",
        encoding="utf-8",
    )
    launcher.chmod(0o755)


def _write_windows_launcher(target_dir: Path) -> None:
    launcher = target_dir / "启动亚马逊2.8.bat"
    launcher.write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        "cd /d %~dp0\r\n"
        "echo 正在检测运行环境...\r\n"
        "AmazonListingTool.exe --env-check --quiet\r\n"
        "if errorlevel 1 (\r\n"
        "  echo.\r\n"
        "  echo 环境检测未通过，请把窗口内容截图发给技术支持。\r\n"
        "  pause\r\n"
        "  exit /b 1\r\n"
        ")\r\n"
        "echo.\r\n"
        "AmazonListingTool.exe\r\n"
        "if errorlevel 1 pause\r\n",
        encoding="utf-8",
    )


def _write_launcher(target_dir: Path) -> None:
    system = platform.system().lower()
    if system == "windows":
        _write_windows_launcher(target_dir)
    else:
        _write_mac_launcher(target_dir)


def _write_env_check_launcher(target_dir: Path) -> None:
    system = platform.system().lower()
    if system == "windows":
        launcher = target_dir / "环境检测.bat"
        launcher.write_text(
            "@echo off\r\n"
            "chcp 65001 >nul\r\n"
            "cd /d %~dp0\r\n"
            "AmazonListingTool.exe --env-check\r\n"
            "echo.\r\n"
            "pause\r\n",
            encoding="utf-8",
        )
    else:
        launcher = target_dir / "环境检测.command"
        launcher.write_text(
            "#!/bin/bash\n"
            "DIR=\"$(cd \"$(dirname \"$0\")\" && pwd)\"\n"
            "cd \"$DIR\"\n"
            "./AmazonListingTool --env-check\n"
            "echo\n"
            "read -n 1 -s -r -p \"按任意键关闭窗口...\"\n"
            "echo\n",
            encoding="utf-8",
        )
        launcher.chmod(0o755)


def _write_release_readme(target_dir: Path) -> None:
    system = platform.system().lower()
    launcher_name = "启动亚马逊2.8.bat" if system == "windows" else "启动亚马逊2.8.command"
    check_name = "环境检测.bat" if system == "windows" else "环境检测.command"
    readme = target_dir / "快速开始.txt"
    readme.write_text(
        "亚马逊 2.8 发行版\n\n"
        f"1. 解压后双击 `{launcher_name}`，不要直接在压缩包里运行\n"
        "2. 启动脚本会先自动检测环境，通过后启动本地服务并打开浏览器\n"
        "3. 第一次使用请进入设置页配置 AI Key 和 Amazon SP-API 账号\n"
        f"4. 如果打不开，先双击 `{check_name}`，把检测结果截图发给技术支持\n"
        "5. 运行数据会写到当前目录下的 input/output/logs/.env/accounts.json\n\n"
        "防呆说明：\n"
        "- 客户电脑不需要安装 Python、pip 或 requirements.txt\n"
        "- 端口 5000 被占用时，程序会自动尝试 5001、5002...\n"
        "- 不要删除 .env 和 accounts.json，它们保存本机配置\n\n"
        "附带文件：\n"
        "- README.md\n"
        "- docs/商家前端使用说明_飞书图文版.md\n"
        "- docs/客户部署说明.md\n"
        "- .env.example\n"
        "- 亚马逊商品采集模板_v1.0.xlsx\n",
        encoding="utf-8",
    )


def _copy_customer_files(target_dir: Path) -> None:
    docs_dir = target_dir / "docs"
    docs_dir.mkdir(exist_ok=True)
    for src, dest in (
        (PROJECT_ROOT / "README.md", target_dir / "README.md"),
        (PROJECT_ROOT / ".env.example", target_dir / ".env.example"),
        (PROJECT_ROOT / "亚马逊商品采集模板_v1.0.xlsx", target_dir / "亚马逊商品采集模板_v1.0.xlsx"),
        (PROJECT_ROOT / "docs" / "商家前端使用说明_飞书图文版.md", docs_dir / "商家前端使用说明_飞书图文版.md"),
        (PROJECT_ROOT / "docs" / "客户部署说明.md", docs_dir / "客户部署说明.md"),
    ):
        if src.exists():
            shutil.copy2(src, dest)


def _remove_macos_metadata(target_dir: Path) -> None:
    for path in target_dir.rglob(".DS_Store"):
        path.unlink(missing_ok=True)


def main() -> int:
    shutil.rmtree(DIST_DIR, ignore_errors=True)
    shutil.rmtree(BUILD_DIR, ignore_errors=True)
    RELEASE_DIR.mkdir(exist_ok=True)

    subprocess.run(_pyinstaller_cmd(), cwd=PROJECT_ROOT, check=True)

    package_name = f"AmazonListingTool-{_platform_suffix()}"
    package_dir = RELEASE_DIR / package_name
    shutil.rmtree(package_dir, ignore_errors=True)
    shutil.copytree(DIST_DIR / "AmazonListingTool", package_dir)

    _copy_customer_files(package_dir)
    _write_launcher(package_dir)
    _write_env_check_launcher(package_dir)
    _write_release_readme(package_dir)
    _remove_macos_metadata(package_dir)

    zip_base = RELEASE_DIR / package_name
    archive = shutil.make_archive(str(zip_base), "zip", root_dir=RELEASE_DIR, base_dir=package_name)

    print(f"发行目录: {package_dir}")
    print(f"压缩包: {archive}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
