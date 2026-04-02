"""
Amazon Bullet Points (Key Product Features) 合规校验。

根据 Amazon 官方规则:
- 单条不超过 500 字符（部分品类 256）
- 禁止包含价格、运费、公司信息、退款承诺
- 禁止 emoji 和特殊符号
"""
import re
from typing import Dict, List

# Amazon Bullet Points 中禁止出现的内容模式
_PROHIBITED_PATTERNS = [
    (re.compile(r'\$\s*\d+[\d,.]*', re.IGNORECASE), '包含价格信息'),
    (re.compile(r'\b(free\s+shipping|fast\s+shipping|ships?\s+free)\b', re.IGNORECASE), '包含运费信息'),
    (re.compile(r'\b(money[\s-]*back|full\s+refund|satisfaction\s+guarant|100%\s+guarant)\b', re.IGNORECASE), '包含退款/保证承诺'),
    (re.compile(r'\b(our\s+company|our\s+brand|we\s+are|about\s+us|founded\s+in)\b', re.IGNORECASE), '包含公司信息'),
    (re.compile(r'\b(contact\s+us|customer\s+service|email\s+us|call\s+us)\b', re.IGNORECASE), '包含联系方式'),
]

# emoji 范围（Unicode emoji blocks）
_EMOJI_PATTERN = re.compile(
    '[\U0001F600-\U0001F64F'   # emoticons
    '\U0001F300-\U0001F5FF'    # symbols & pictographs
    '\U0001F680-\U0001F6FF'    # transport & map
    '\U0001F1E0-\U0001F1FF'    # flags
    '\U00002702-\U000027B0'    # dingbats
    '\U0000FE00-\U0000FE0F'    # variation selectors
    '\U0001F900-\U0001F9FF'    # supplemental symbols
    '\U0001FA00-\U0001FA6F'    # chess symbols
    '\U0001FA70-\U0001FAFF'    # symbols extended-A
    '\U00002600-\U000026FF'    # misc symbols
    ']'
)


def validate_bullet(text: str, index: int = 0, max_chars: int = 500) -> Dict:
    """校验单条 bullet point 的合规性。"""
    text = str(text or '').strip()
    issues = []

    if not text:
        return {'valid': True, 'issues': []}

    # 长度
    if len(text) > max_chars:
        issues.append({
            'level': 'error',
            'message': f'卖点{index}超过{max_chars}字符({len(text)}字符)',
        })

    # 禁止内容
    for pattern, description in _PROHIBITED_PATTERNS:
        match = pattern.search(text)
        if match:
            issues.append({
                'level': 'warning',
                'message': f'卖点{index}{description}: "{match.group()}"',
            })

    # emoji
    emojis = _EMOJI_PATTERN.findall(text)
    if emojis:
        issues.append({
            'level': 'warning',
            'message': f'卖点{index}包含emoji: {"".join(emojis[:5])}',
        })

    return {
        'valid': not any(i['level'] == 'error' for i in issues),
        'issues': issues,
    }


def validate_bullets(bullets: List[str], max_chars: int = 500) -> Dict:
    """校验全部 bullet points。"""
    all_issues = []
    all_valid = True

    for i, bp in enumerate(bullets, 1):
        result = validate_bullet(bp, index=i, max_chars=max_chars)
        all_issues.extend(result['issues'])
        if not result['valid']:
            all_valid = False

    return {'valid': all_valid, 'issues': all_issues}


def clean_bullet(text: str) -> str:
    """清理单条 bullet point 中的 emoji。"""
    return _EMOJI_PATTERN.sub('', str(text or '')).strip()
