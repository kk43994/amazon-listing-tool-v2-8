"""
运行时路径辅助

区分两类路径：
- 资源路径：模板、静态文件、默认配置等，打包后位于只读资源目录
- 运行路径：.env、accounts.json、input/output/logs 等，打包后应写到可写目录
"""
from __future__ import annotations

import os
import sys
from pathlib import Path


def is_frozen_app() -> bool:
    return bool(getattr(sys, "frozen", False))


def get_repo_root() -> Path:
    return Path(__file__).resolve().parent.parent


def _looks_like_resource_root(path: Path) -> bool:
    return path.exists() and (
        (path / "web").exists()
        or (path / "config").exists()
        or (path / ".env.example").exists()
    )


def get_resource_root() -> Path:
    if is_frozen_app():
        meipass = getattr(sys, "_MEIPASS", "")
        candidates = []
        if meipass:
            candidates.append(Path(meipass))
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / "_internal")
        candidates.append(exe_dir)
        for candidate in candidates:
            if _looks_like_resource_root(candidate):
                return candidate
        return exe_dir
    return get_repo_root()


def get_runtime_root() -> Path:
    if is_frozen_app():
        return Path(sys.executable).resolve().parent
    return get_repo_root()


def resource_path(*parts: str) -> str:
    return str(get_resource_root().joinpath(*parts))


def runtime_path(*parts: str) -> str:
    return str(get_runtime_root().joinpath(*parts))
