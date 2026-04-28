"""
构建面向终端用户的发行版

输出内容：
- release/<包名>/ 可运行目录
- release/<包名>.zip 压缩包
"""
from __future__ import annotations

import hashlib
import importlib.metadata as importlib_metadata
import json
import os
import platform
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DIST_DIR = PROJECT_ROOT / "dist"
BUILD_DIR = PROJECT_ROOT / "build"
RELEASE_DIR = PROJECT_ROOT / "release"


def _project_version() -> str:
    version_file = PROJECT_ROOT / "VERSION"
    if version_file.exists():
        return version_file.read_text(encoding="utf-8").strip() or "0.0.0-dev"
    return "0.0.0-dev"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _normalize_requirement_name(name: str) -> str:
    return str(name or "").strip().lower().replace("_", "-")


def _iter_requirements(paths: tuple[Path, ...]) -> list[str]:
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


def _import_name_for_package(package: str) -> str:
    mapping = {
        "Pillow": "PIL",
        "python-dotenv": "dotenv",
        "python-amazon-sp-api": "sp_api",
    }
    return mapping.get(package, package.replace("-", "_"))


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
        _add_data_arg(PROJECT_ROOT / "VERSION", "."),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "docs" / "商家前端使用说明_飞书图文版.md", "docs"),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "docs" / "客户部署说明.md", "docs"),
        "--add-data",
        _add_data_arg(PROJECT_ROOT / "docs" / "RELEASE_CHECKLIST.md", "docs"),
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
        "echo \"亚马逊 2.8 客户启动器\"\n"
        "echo \"请不要关闭这个窗口；关闭窗口会停止软件。\"\n"
        "echo\n"
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
        "echo \"检测通过，正在启动工作台...\"\n"
        "echo \"如果浏览器没有自动打开，请复制窗口里的 http://127.0.0.1:端口 到浏览器。\"\n"
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
        "echo 亚马逊 2.8 客户启动器\r\n"
        "echo 请不要关闭这个窗口；关闭窗口会停止软件。\r\n"
        "echo.\r\n"
        "echo 正在检测运行环境...\r\n"
        "AmazonListingTool.exe --env-check --quiet\r\n"
        "if errorlevel 1 (\r\n"
        "  echo.\r\n"
        "  echo 环境检测未通过，请把窗口内容截图发给技术支持。\r\n"
        "  pause\r\n"
        "  exit /b 1\r\n"
        ")\r\n"
        "echo.\r\n"
        "echo 检测通过，正在启动工作台...\r\n"
        "echo 如果浏览器没有自动打开，请复制窗口里的 http://127.0.0.1:端口 到浏览器。\r\n"
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


def _write_doctor_launcher(target_dir: Path) -> None:
    system = platform.system().lower()
    if system == "windows":
        launcher = target_dir / "一键检测修复.bat"
        launcher.write_text(
            "@echo off\r\n"
            "chcp 65001 >nul\r\n"
            "cd /d %~dp0\r\n"
            "echo 亚马逊 2.8 Doctor 一键检测修复\r\n"
            "echo 会自动补齐 .env、accounts.json、input/output/logs，并检查内置依赖、端口、浏览器和外网连通性。\r\n"
            "echo.\r\n"
            "AmazonListingTool.exe --doctor\r\n"
            "set STATUS=%ERRORLEVEL%\r\n"
            "echo.\r\n"
            "echo 检测报告已保存到 logs\\doctor-report.txt\r\n"
            "if not \"%STATUS%\"==\"0\" echo 检测发现必须处理的问题，请把本窗口和 logs\\doctor-report.txt 发给技术支持。\r\n"
            "pause\r\n"
            "exit /b %STATUS%\r\n",
            encoding="utf-8",
        )
    else:
        launcher = target_dir / "一键检测修复.command"
        launcher.write_text(
            "#!/bin/bash\n"
            "DIR=\"$(cd \"$(dirname \"$0\")\" && pwd)\"\n"
            "cd \"$DIR\"\n"
            "echo \"亚马逊 2.8 Doctor 一键检测修复\"\n"
            "echo \"会自动补齐 .env、accounts.json、input/output/logs，并检查内置依赖、端口、浏览器和外网连通性。\"\n"
            "echo\n"
            "./AmazonListingTool --doctor\n"
            "STATUS=$?\n"
            "echo\n"
            "echo \"检测报告已保存到 logs/doctor-report.txt\"\n"
            "if [ $STATUS -ne 0 ]; then\n"
            "  echo \"检测发现必须处理的问题，请把本窗口和 logs/doctor-report.txt 发给技术支持。\"\n"
            "fi\n"
            "read -n 1 -s -r -p \"按任意键关闭窗口...\"\n"
            "echo\n"
            "exit $STATUS\n",
            encoding="utf-8",
        )
        launcher.chmod(0o755)


def _write_support_bundle_launcher(target_dir: Path) -> None:
    system = platform.system().lower()
    if system == "windows":
        launcher = target_dir / "导出支持包.bat"
        launcher.write_text(
            "@echo off\r\n"
            "chcp 65001 >nul\r\n"
            "cd /d %~dp0\r\n"
            "echo 正在导出技术支持包，请稍等...\r\n"
            "AmazonListingTool.exe --support-bundle\r\n"
            "echo.\r\n"
            "echo 支持包已保存到 logs 目录，请把最新的 support-bundle-*.zip 发给技术支持。\r\n"
            "pause\r\n",
            encoding="utf-8",
        )
    else:
        launcher = target_dir / "导出支持包.command"
        launcher.write_text(
            "#!/bin/bash\n"
            "DIR=\"$(cd \"$(dirname \"$0\")\" && pwd)\"\n"
            "cd \"$DIR\"\n"
            "echo \"正在导出技术支持包，请稍等...\"\n"
            "./AmazonListingTool --support-bundle\n"
            "echo\n"
            "echo \"支持包已保存到 logs 目录，请把最新的 support-bundle-*.zip 发给技术支持。\"\n"
            "read -n 1 -s -r -p \"按任意键关闭窗口...\"\n"
            "echo\n",
            encoding="utf-8",
        )
        launcher.chmod(0o755)


def _write_folder_launchers(target_dir: Path) -> None:
    system = platform.system().lower()
    if system == "windows":
        scripts = {
            "打开输出目录.bat": "if not exist output mkdir output\r\nexplorer output\r\n",
            "打开备份目录.bat": "if not exist backups mkdir backups\r\nexplorer backups\r\n",
        }
        for name, body in scripts.items():
            (target_dir / name).write_text(
                "@echo off\r\n"
                "chcp 65001 >nul\r\n"
                "cd /d %~dp0\r\n"
                + body,
                encoding="utf-8",
            )
    else:
        scripts = {
            "打开输出目录.command": "mkdir -p output\nopen output\n",
            "打开备份目录.command": "mkdir -p backups\nopen backups\n",
        }
        for name, body in scripts.items():
            path = target_dir / name
            path.write_text(
                "#!/bin/bash\n"
                "DIR=\"$(cd \"$(dirname \"$0\")\" && pwd)\"\n"
                "cd \"$DIR\"\n"
                + body,
                encoding="utf-8",
            )
            path.chmod(0o755)


def _write_release_readme(target_dir: Path) -> None:
    system = platform.system().lower()
    launcher_name = "启动亚马逊2.8.bat" if system == "windows" else "启动亚马逊2.8.command"
    check_name = "环境检测.bat" if system == "windows" else "环境检测.command"
    doctor_name = "一键检测修复.bat" if system == "windows" else "一键检测修复.command"
    support_name = "导出支持包.bat" if system == "windows" else "导出支持包.command"
    readme = target_dir / "快速开始.txt"
    readme.write_text(
        "亚马逊 2.8 发行版\n\n"
        f"1. 解压后双击 `{launcher_name}`，不要直接在压缩包里运行\n"
        "2. 启动脚本会先自动检测环境，通过后启动本地服务并打开浏览器\n"
        "3. 第一次使用请进入设置页配置 AI Key 和 Amazon SP-API 账号\n"
        f"4. 如果打不开，先双击 `{doctor_name}`；只想快速看环境可双击 `{check_name}`\n"
        "5. 运行数据会写到当前目录下的 input/output/logs/.env/accounts.json\n\n"
        "防呆说明：\n"
        "- 客户电脑不需要安装 Python、pip 或 requirements.txt\n"
        "- 如果中文文件名打不开，可改双击 Start-Amazon-2.8.command / Start-Amazon-2.8.bat\n"
        "- 端口 5000 被占用时，程序会自动尝试 5001、5002...\n"
        "- Doctor 会自动补齐 .env、accounts.json、input/output/logs，并检查 AI 中转域名和 Amazon 网络\n"
        "- 如果客服要排查，双击导出支持包，会生成已脱敏的 logs/support-bundle-*.zip\n"
        "- Excel 写回前会自动备份到 backups，可双击 Open-Backups / 打开备份目录 查看\n"
        "- 不要删除 .env 和 accounts.json，它们保存本机配置\n\n"
        "附带文件：\n"
        "- 客户先看这里.txt\n"
        "- Read-Me-First.txt\n"
        f"- {doctor_name}\n"
        "- Doctor.command / Doctor.bat\n"
        f"- {support_name}\n"
        "- Support-Bundle.command / Support-Bundle.bat\n"
        f"- {check_name}\n"
        "- Env-Check.command / Env-Check.bat\n"
        "- release-manifest.json\n"
        "- dependency-inventory.json\n"
        "- Open-Output.command / Open-Output.bat\n"
        "- Open-Backups.command / Open-Backups.bat\n"
        "- README.md\n"
        "- docs/商家前端使用说明_飞书图文版.md\n"
        "- docs/客户部署说明.md\n"
        "- docs/RELEASE_CHECKLIST.md\n"
        "- .env.example\n"
        "- 亚马逊商品采集模板_v1.0.xlsx\n",
        encoding="utf-8",
    )


def _write_customer_readme(target_dir: Path) -> None:
    system = platform.system().lower()
    launcher_name = "启动亚马逊2.8.bat" if system == "windows" else "启动亚马逊2.8.command"
    check_name = "环境检测.bat" if system == "windows" else "环境检测.command"
    doctor_name = "一键检测修复.bat" if system == "windows" else "一键检测修复.command"
    support_name = "导出支持包.bat" if system == "windows" else "导出支持包.command"
    (target_dir / "客户先看这里.txt").write_text(
        "亚马逊 2.8 客户使用说明（只看这一份也可以）\n"
        "================================================\n\n"
        "第一步：先解压\n"
        "- 不要在压缩包里直接运行。\n"
        "- 解压后进入文件夹，再双击启动文件。\n\n"
        f"第二步：双击 {launcher_name}\n"
        "- 打开的黑色窗口不要关，关掉软件就停止了。\n"
        "- 正常情况下浏览器会自动打开工作台。\n"
        "- 如果浏览器没打开，就复制黑色窗口里的 http://127.0.0.1:端口 到浏览器。\n\n"
        "第三步：第一次打开先点左侧“设置”\n"
        "1. API 中转域名固定填：https://api.kk666.best\n"
        "2. 文字模型默认：gemini-3.1-flash-lite-preview\n"
        "3. 图片模型默认：gemini-3.1-flash-image-preview\n"
        "4. 客户只需要粘贴服务商给的文字 Key、图片 Key；不懂就点网页里的“一键恢复推荐”。\n"
        "5. 点“添加亚马逊账号”，填 Seller ID、LWA Client ID、LWA Client Secret、Refresh Token。\n\n"
        "第四步：回到“工作台”处理商品\n"
        "推荐顺序固定：先生成示例 Excel 测一下 -> 导入真实 Excel -> AI 文案/生图 -> 校验/缺项诊断 -> Amazon 预览 -> 正式提交。\n"
        "不要跳过 Amazon 预览直接正式提交。\n\n"
        "遇到问题怎么办\n"
        f"- 优先双击 {doctor_name}，它会自动补齐基础文件并生成 logs/doctor-report.txt。\n"
        f"- 如果客服让你发排查文件，双击 {support_name}，把 logs 里的最新 support-bundle-*.zip 发过去。\n"
        "- Excel 写回前会自动备份到 backups，误操作时让客服找最新备份。\n"
        f"- 只想快速检查可双击 {check_name}，或在设置页点“一键自检”。\n"
        "- 把自检结果截图发给技术支持。\n"
        "- 不要删除 .env、accounts.json、input、output、logs，这些是本机配置和运行数据。\n",
        encoding="utf-8",
    )


def _write_ascii_fallbacks(target_dir: Path) -> None:
    system = platform.system().lower()
    pairs = (
        ("启动亚马逊2.8.bat", "Start-Amazon-2.8.bat"),
        ("启动亚马逊2.8.command", "Start-Amazon-2.8.command"),
        ("一键检测修复.bat", "Doctor.bat"),
        ("一键检测修复.command", "Doctor.command"),
        ("导出支持包.bat", "Support-Bundle.bat"),
        ("导出支持包.command", "Support-Bundle.command"),
        ("打开输出目录.bat", "Open-Output.bat"),
        ("打开输出目录.command", "Open-Output.command"),
        ("打开备份目录.bat", "Open-Backups.bat"),
        ("打开备份目录.command", "Open-Backups.command"),
        ("环境检测.bat", "Env-Check.bat"),
        ("环境检测.command", "Env-Check.command"),
        ("客户先看这里.txt", "Read-Me-First.txt"),
    )
    for src_name, dest_name in pairs:
        src = target_dir / src_name
        if not src.exists():
            continue
        dest = target_dir / dest_name
        shutil.copy2(src, dest)
        if system != "windows" and dest.suffix == ".command":
            dest.chmod(0o755)


def _write_release_manifest(target_dir: Path, package_name: str) -> None:
    system = platform.system().lower()
    executable = target_dir / ("AmazonListingTool.exe" if system == "windows" else "AmazonListingTool")
    root_files = [
        path.name for path in sorted(target_dir.iterdir(), key=lambda item: item.name)
        if path.is_file()
    ]
    if "release-manifest.json" not in root_files:
        root_files.append("release-manifest.json")
    manifest = {
        "app": "AmazonListingTool",
        "version": _project_version(),
        "package": package_name,
        "platform": _platform_suffix(),
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "python": platform.python_version(),
        "executable": executable.name,
        "executable_sha256": _sha256(executable) if executable.exists() else "",
        "root_files": root_files,
        "doctor": "Doctor.bat" if system == "windows" else "Doctor.command",
        "start": "Start-Amazon-2.8.bat" if system == "windows" else "Start-Amazon-2.8.command",
        "env_check": "Env-Check.bat" if system == "windows" else "Env-Check.command",
        "support_bundle": "Support-Bundle.bat" if system == "windows" else "Support-Bundle.command",
        "open_output": "Open-Output.bat" if system == "windows" else "Open-Output.command",
        "open_backups": "Open-Backups.bat" if system == "windows" else "Open-Backups.command",
    }
    (target_dir / "release-manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _write_dependency_inventory(target_dir: Path) -> None:
    requirements = _iter_requirements((PROJECT_ROOT / "requirements.txt",))
    inventory = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "python": platform.python_version(),
        "requirements": [],
    }
    for package in requirements:
        try:
            installed_version = importlib_metadata.version(package)
        except importlib_metadata.PackageNotFoundError:
            installed_version = ""
        inventory["requirements"].append({
            "package": package,
            "import_name": _import_name_for_package(package),
            "installed_version": installed_version,
        })
    (target_dir / "dependency-inventory.json").write_text(
        json.dumps(inventory, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _copy_customer_files(target_dir: Path) -> None:
    docs_dir = target_dir / "docs"
    docs_dir.mkdir(exist_ok=True)
    for src, dest in (
        (PROJECT_ROOT / "README.md", target_dir / "README.md"),
        (PROJECT_ROOT / "VERSION", target_dir / "VERSION"),
        (PROJECT_ROOT / ".env.example", target_dir / ".env.example"),
        (PROJECT_ROOT / "亚马逊商品采集模板_v1.0.xlsx", target_dir / "亚马逊商品采集模板_v1.0.xlsx"),
        (PROJECT_ROOT / "docs" / "商家前端使用说明_飞书图文版.md", docs_dir / "商家前端使用说明_飞书图文版.md"),
        (PROJECT_ROOT / "docs" / "客户部署说明.md", docs_dir / "客户部署说明.md"),
        (PROJECT_ROOT / "docs" / "RELEASE_CHECKLIST.md", docs_dir / "RELEASE_CHECKLIST.md"),
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
    _write_doctor_launcher(package_dir)
    _write_support_bundle_launcher(package_dir)
    _write_folder_launchers(package_dir)
    _write_release_readme(package_dir)
    _write_customer_readme(package_dir)
    _write_dependency_inventory(package_dir)
    _write_ascii_fallbacks(package_dir)
    _write_release_manifest(package_dir, package_name)
    _remove_macos_metadata(package_dir)

    zip_base = RELEASE_DIR / package_name
    archive = shutil.make_archive(str(zip_base), "zip", root_dir=RELEASE_DIR, base_dir=package_name)
    archive_path = Path(archive)
    checksum_path = archive_path.with_suffix(archive_path.suffix + ".sha256")
    checksum_path.write_text(f"{_sha256(archive_path)}  {archive_path.name}\n", encoding="utf-8")

    print(f"发行目录: {package_dir}")
    print(f"压缩包: {archive}")
    print(f"SHA256: {checksum_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
