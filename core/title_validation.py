"""
Amazon 商品标题合规校验。

根据 Amazon 官方规则 (2025.01 生效):
- 大多数品类标题不超过 200 字符
- 同一单词最多出现 2 次（介词/冠词/连词除外）
- 禁用字符: ! $ ? _ { } ^ ¬ ¦
- 不允许全大写单词（品牌名除外）
"""
import re
from typing import Dict, List, Tuple

# Amazon 明确禁止的标题字符（除非是品牌名的一部分）
BANNED_CHARS = set('!$?_{}^~\u00ac\u00a6')

# 不受同词重复限制的虚词（介词/冠词/连词），覆盖主要站点语言
EXEMPT_WORDS = {
    # English
    'the', 'a', 'an', 'of', 'for', 'in', 'on', 'at', 'to', 'by',
    'with', 'and', 'or', 'is', 'it', 'as', 'from', 'that', 'this',
    'but', 'not', 'be', 'are', 'was', 'were', 'no', 'nor', 'so', 'yet',
    # German
    'der', 'die', 'das', 'ein', 'eine', 'und', 'oder', 'mit', 'für',
    'auf', 'aus', 'bei', 'nach', 'von', 'zu', 'im', 'am', 'um',
    # French
    'le', 'la', 'les', 'un', 'une', 'des', 'du', 'de', 'et', 'ou',
    'en', 'au', 'aux', 'par', 'sur', 'dans', 'pour', 'avec',
    # Spanish
    'el', 'los', 'las', 'un', 'una', 'unos', 'unas', 'y', 'o',
    'del', 'al', 'con', 'por', 'para', 'sin',
    # Italian
    'il', 'lo', 'gli', 'i', 'e', 'ed', 'di', 'da', 'con', 'per',
    # Japanese particles (romaji, unlikely in titles but for completeness)
    'no', 'ni', 'wo', 'wa', 'ga', 'to', 'de', 'he', 'mo',
}


def find_banned_characters(title: str, brand: str = '') -> List[Dict]:
    """找出标题中 Amazon 禁止的字符。品牌名中的字符豁免。"""
    title = str(title or '')
    brand = str(brand or '').strip()

    # 标记品牌名在标题中的位置范围（豁免区间）
    exempt_ranges = []
    if brand:
        start = title.lower().find(brand.lower())
        if start >= 0:
            exempt_ranges.append((start, start + len(brand)))

    violations = []
    for i, ch in enumerate(title):
        if ch in BANNED_CHARS:
            if not any(start <= i < end for start, end in exempt_ranges):
                violations.append({'char': ch, 'position': i})

    return violations


def find_duplicate_words(title: str) -> List[Dict]:
    """找出标题中重复超过 2 次的非虚词。"""
    words = re.findall(r'[A-Za-z0-9\u00C0-\u024F]+', str(title or '').lower())
    counts = {}
    for w in words:
        if w not in EXEMPT_WORDS and len(w) > 1:
            counts[w] = counts.get(w, 0) + 1

    return [{'word': w, 'count': c} for w, c in counts.items() if c > 2]


def fix_title(title: str, brand: str = '') -> Tuple[str, List[str]]:
    """自动修复标题：删除禁用字符、去除多余重复词。返回 (修复后标题, 修改记录)。"""
    changes = []
    result = str(title or '')
    brand = str(brand or '').strip()

    # 1. 删除禁用字符（品牌名中的豁免）
    banned = find_banned_characters(result, brand)
    if banned:
        chars_to_remove = set(v['char'] for v in banned)
        cleaned = []
        for i, ch in enumerate(result):
            if ch in chars_to_remove and any(v['position'] == i for v in banned):
                continue
            cleaned.append(ch)
        result = ''.join(cleaned)
        result = re.sub(r'  +', ' ', result).strip()
        changes.append(f"删除禁用字符: {', '.join(repr(c) for c in sorted(chars_to_remove))}")

    # 2. 去除第 3 次及以上出现的重复词
    duplicates = find_duplicate_words(result)
    for dup in duplicates:
        word = dup['word']
        seen = 0
        tokens = re.split(r'(\s+)', result)
        new_tokens = []
        for token in tokens:
            if token.lower() == word:
                seen += 1
                if seen > 2:
                    continue
            new_tokens.append(token)
        result = ''.join(new_tokens)
        result = re.sub(r'  +', ' ', result).strip()
        changes.append(f"去除重复词 '{word}'（出现{dup['count']}次，保留2次）")

    return result, changes


def validate_title(title: str, brand: str = '') -> Dict:
    """综合校验标题合规性。"""
    banned = find_banned_characters(title, brand)
    duplicates = find_duplicate_words(title)
    is_valid = not banned and not duplicates

    result = {
        'valid': is_valid,
        'banned_chars': banned,
        'duplicate_words': duplicates,
        'suggested_fix': None,
        'fix_changes': [],
    }

    if not is_valid:
        fixed, changes = fix_title(title, brand)
        result['suggested_fix'] = fixed
        result['fix_changes'] = changes

    return result
