"""
共享工具函数
"""
import os
import re
import logging
from typing import List, Dict

logger = logging.getLogger(__name__)


def sanitize_filename(filename: str, fallback: str = 'upload.xlsx') -> str:
    """
    清理文件名，保留 Unicode 字符（如中文），去除路径遍历和危险字符。
    与 werkzeug.secure_filename 不同，本函数不会剥离非 ASCII 字符。
    """
    name = str(filename or '').strip()
    # 去掉路径分隔符和路径遍历
    name = name.replace('\\', '/').split('/')[-1]
    name = re.sub(r'\.\.+', '_', name)
    # 去掉控制字符和 shell 特殊字符，保留 Unicode 字母/数字/常见符号
    name = re.sub(r'[\x00-\x1f\x7f<>:"|?*]', '_', name)
    name = name.strip('. ')
    return name or fallback


def filter_rows(data: List[Dict], rows: str) -> List[Dict]:
    """
    根据行范围规格过滤数据列表。

    支持格式:
      - "3"     → 只取第3行
      - "1-10"  → 取第1到第10行
    行号从1开始。无效规格返回全部数据并记录警告。
    """
    if not rows or not rows.strip():
        return data

    rows = rows.strip()
    total = len(data)
    try:
        if '-' in rows:
            parts = rows.split('-', 1)
            start = int(parts[0]) - 1
            end = int(parts[1])
        else:
            start = int(rows) - 1
            end = start + 1

        if start < 0:
            start = 0
        if end > total:
            end = total
        if start >= end:
            logger.warning("行范围无效 (start >= end): %s，使用全部数据", rows)
            return data
        return data[start:end]
    except (ValueError, IndexError):
        logger.warning("行范围解析失败: %s，使用全部数据", rows)
        return data
