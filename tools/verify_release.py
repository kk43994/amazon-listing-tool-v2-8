"""Verify customer release folders contain the expected handoff files."""
from __future__ import annotations

import argparse
import json
import platform
from pathlib import Path


def _expected_files(system: str) -> set[str]:
    if system == "windows":
        return {
            "AmazonListingTool.exe",
            "启动亚马逊2.8.bat",
            "Start-Amazon-2.8.bat",
            "一键检测修复.bat",
            "Doctor.bat",
            "导出支持包.bat",
            "Support-Bundle.bat",
            "打开输出目录.bat",
            "Open-Output.bat",
            "打开备份目录.bat",
            "Open-Backups.bat",
            "环境检测.bat",
            "Env-Check.bat",
        }
    return {
        "AmazonListingTool",
        "启动亚马逊2.8.command",
        "Start-Amazon-2.8.command",
        "一键检测修复.command",
        "Doctor.command",
            "导出支持包.command",
            "Support-Bundle.command",
            "打开输出目录.command",
            "Open-Output.command",
            "打开备份目录.command",
            "Open-Backups.command",
            "环境检测.command",
        "Env-Check.command",
    }


def verify_release_dir(package_dir: Path, system: str | None = None) -> list[str]:
    system = (system or platform.system()).lower()
    expected = _expected_files(system)
    expected.update({
        "客户先看这里.txt",
        "Read-Me-First.txt",
        "快速开始.txt",
        ".env.example",
        "README.md",
        "VERSION",
        "release-manifest.json",
        "dependency-inventory.json",
        "亚马逊商品采集模板_v1.0.xlsx",
    })
    errors: list[str] = []
    for name in sorted(expected):
        if not (package_dir / name).exists():
            errors.append(f"缺少根目录文件: {name}")
    for name in ("商家前端使用说明_飞书图文版.md", "客户部署说明.md", "RELEASE_CHECKLIST.md"):
        if not (package_dir / "docs" / name).exists():
            errors.append(f"缺少 docs/{name}")

    manifest_path = package_dir / "release-manifest.json"
    if manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception as exc:
            errors.append(f"release-manifest.json 无法解析: {exc}")
        else:
            root_files = set(manifest.get("root_files") or [])
            for name in ("release-manifest.json", "dependency-inventory.json", "Read-Me-First.txt"):
                if name not in root_files:
                    errors.append(f"manifest root_files 未记录: {name}")
            if not manifest.get("version"):
                errors.append("manifest 缺少 version")
            if not manifest.get("executable_sha256"):
                errors.append("manifest 缺少 executable_sha256")
    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Verify AmazonListingTool release folder")
    parser.add_argument("package_dir", type=Path)
    parser.add_argument("--system", choices=("windows", "darwin", "linux"), default=None)
    args = parser.parse_args()

    errors = verify_release_dir(args.package_dir, system=args.system)
    if errors:
        for error in errors:
            print(f"[ERROR] {error}")
        return 1
    print(f"Release verification OK: {args.package_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
