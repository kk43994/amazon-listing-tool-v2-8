"""
Amazon 搜索关键词处理工具。

根据 Amazon 官方规则：
- 字节限制因站点不同: US/EU=250, JP=500, IN=200
- 空格和标点不计入字节限制
- 超限会导致整条搜索词全部失效（不是截断）
- 不需要重复标题/五点中已有的词
"""
import re
import string
from typing import List

# 空格 + 常见标点，Amazon 明确不计入字节限制
_EXCLUDED_CHARS = set(string.whitespace + '.,;:!?\'"/-')


def count_search_term_bytes(text: str) -> int:
    """计算搜索词的有效字节数（仅统计单词本身，排除空格和标点）。"""
    total = 0
    for word in str(text or '').split():
        stripped = word.strip(string.punctuation)
        if stripped:
            total += len(stripped.encode('utf-8'))
    return total


def dedup_search_terms(search_terms: str, title: str, bullets: List[str] = None) -> str:
    """移除搜索词中已出现在标题或五点描述中的单词（大小写不敏感）。"""
    existing_words = set()
    for source in [title] + (bullets or []):
        for word in re.findall(r'[A-Za-z0-9\u00C0-\u024F\u3000-\u9FFF]+', str(source or '').lower()):
            if len(word) > 1:  # 跳过单字符（冠词 a 等太短不值得去重）
                existing_words.add(word)

    result = []
    for word in str(search_terms or '').split():
        cleaned = word.strip(string.punctuation).lower()
        if cleaned and cleaned not in existing_words:
            result.append(word)

    return ' '.join(result)


def truncate_search_terms(search_terms: str, byte_limit: int) -> str:
    """按字节限制截断搜索词（使用 Amazon 的计算规则：仅统计单词字节）。"""
    words = str(search_terms or '').split()
    result = []
    total = 0
    for word in words:
        stripped = word.strip(string.punctuation)
        if not stripped:
            continue
        word_bytes = len(stripped.encode('utf-8'))
        if total + word_bytes > byte_limit:
            break
        result.append(word)
        total += word_bytes
    return ' '.join(result)
