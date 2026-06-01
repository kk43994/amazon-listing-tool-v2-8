"""Seller-facing notes for common Amazon SP-API issue messages."""
import re
from typing import Any, Dict, Iterable, List


_CODE_NOTES = {
    'InvalidInput': '参数无效：Amazon 拒绝了本次请求，但没有指出具体字段。优先检查商品类型是否适用于当前站点/账号、必填字段是否齐全、图片 URL 是否可访问、商品编码或免码设置是否正确。',
    'MISSING_REQUIRED_ATTRIBUTE': '缺少必填属性：请补齐 Amazon 标记的字段后重新预览。',
    'INVALID_ATTRIBUTE_VALUE': '字段值无效：当前填写的值不符合 Amazon 当前类目的格式或枚举要求。',
    'ATTRIBUTE_VALUE_TOO_LONG': '字段内容过长：请缩短该字段内容后重新预览。',
    'ATTRIBUTE_VALUE_TOO_SHORT': '字段内容过短：请补充更完整的字段内容。',
    'DUPLICATE_VALUE': '字段值重复：请检查该字段是否与其他商品或其他字段重复。',
    'INVALID_UPC': 'UPC 条码无效：请确认 UPC 是真实有效的 12 位条码。',
    'INVALID_EAN': 'EAN 条码无效：请确认 EAN 是真实有效的 13 位条码。',
    'INVALID_GTIN': 'GTIN 无效：请确认商品编码来自有效来源，或改用免码流程。',
    'PRODUCT_IDENTIFIER_NOT_FOUND': '商品编码未被 Amazon 识别：请确认条码是否在 GS1 等官方库中有效，或确认是否应申请/使用免码。',
    'INVALID_IMAGE_URL': '图片 URL 无效：请确认图片链接可公开访问，且返回真实图片文件。',
    'IMAGE_TOO_SMALL': '图片尺寸过小：主图建议至少 1000x1000 像素。',
    'INVALID_PRICE': '价格无效：请确认价格是大于 0 的数字，不要带货币符号。',
    'PRICE_TOO_LOW': '价格过低：当前价格低于 Amazon 允许范围，请确认小数点和币种。',
    'PRICE_TOO_HIGH': '价格过高：当前价格高于 Amazon 允许范围，请确认小数点和币种。',
    'MISSING_PARENT_SKU': '缺少父体 SKU：子体商品必须填写 parent_sku 并关联到有效父体。',
    'INVALID_VARIATION_THEME': '变体主题无效：当前类目不接受这个 variation_theme，请按官方模板可选值调整。',
    'UNAUTHORIZED': '账号无权限：当前账号可能没有该类目、品牌或站点的上架权限。',
    'BRAND_NOT_APPROVED': '品牌未授权：请先确认品牌备案、授权或可售状态。',
    '100239': '标题和主图不匹配：Amazon 认为标题描述的商品与主图不是同一个商品。请换成与标题一致的主图，或把标题改成主图实际展示的商品。',
}


_MESSAGE_NOTES = (
    (
        re.compile(r'unable to retrieve media content|image file type is(?:n\'t| not) supported|invalid image url', re.I),
        'Amazon 无法读取图片：请确认图片链接可以在无登录状态下打开，文件是 JPG/PNG 等支持格式，并且不会返回 403、404 或网页内容。',
    ),
    (
        re.compile(r'title and the main image .*same product|main image .*same product', re.I),
        _CODE_NOTES['100239'],
    ),
    (
        re.compile(r'required but missing|missing required', re.I),
        '缺少必填字段：请补齐提示中提到的 Amazon 字段后重新预览。',
    ),
    (
        re.compile(r'not a valid value|invalid value|value .* invalid', re.I),
        '字段值不符合类目规则：请按官方模板或字段下拉选项重新选择。',
    ),
    (
        re.compile(r'externally assigned product identifier|external product id|product identifier', re.I),
        '商品身份信息不足或无效：请补真实 UPC/EAN/GTIN，或确认该商品是否符合免码条件。',
    ),
    (
        re.compile(r'merchant suggested asin|merchant_suggested_asin', re.I),
        'Amazon 没有拿到可匹配的 ASIN：如果是跟卖已有商品，请补 ASIN；如果是新建商品，请确认编码/免码路径。',
    ),
    (
        re.compile(r'attribute .* does not belong|does not belong to product type', re.I),
        '当前类目不接受这个字段：通常是商品类型选错，或把其他类目的字段带进了当前商品。',
    ),
    (
        re.compile(r'dangerous goods|hazmat|supplier_declared_dg_hz_regulation', re.I),
        '危化品/合规字段需要确认：如果商品不是危险品，通常填写 not_applicable；如果含电池或液体，需要按实际情况补充。',
    ),
    (
        re.compile(r'invalid parameters provided', re.I),
        _CODE_NOTES['InvalidInput'],
    ),
)


def _clean_text(value: Any) -> str:
    return str(value or '').strip()


def _clean_list(value: Any) -> List[str]:
    if value is None:
        return []
    if isinstance(value, str):
        values: Iterable[Any] = [value]
    elif isinstance(value, Iterable):
        values = value
    else:
        values = [value]
    return [_clean_text(item) for item in values if _clean_text(item)]


def explain_amazon_issue(issue: Dict[str, Any] | None, fallback_message: str = '') -> Dict[str, str]:
    """Return English source text plus a Chinese seller note for an Amazon issue."""
    issue = issue or {}
    code = _clean_text(issue.get('code'))
    message_en = _clean_text(issue.get('message') or issue.get('details') or fallback_message)
    attrs = _clean_list(issue.get('attributeNames') or issue.get('attributes'))

    note = ''
    for pattern, candidate in _MESSAGE_NOTES:
        if pattern.search(message_en):
            note = candidate
            break
    if not note:
        note = _CODE_NOTES.get(code, '')

    hint = _clean_text(issue.get('hint'))
    fix = _clean_text(issue.get('fix'))
    extras = [item for item in (hint, fix) if item and item not in note]
    if not note and extras:
        note = '；'.join(extras)
    elif note and extras:
        note = f"{note} {'；'.join(extras)}"

    if note and attrs and '相关字段' not in note:
        note = f"{note} 相关字段：{'、'.join(attrs)}。"

    display_message = ''
    if message_en and note:
        display_message = f"{message_en}（{note}）"
    elif message_en:
        display_message = message_en
    elif note:
        display_message = note

    return {
        'message_en': message_en,
        'message_zh': note,
        'seller_hint': note,
        'display_message': display_message,
    }


def annotate_amazon_issue(issue: Dict[str, Any], fallback_message: str = '') -> Dict[str, Any]:
    """Mutate and return an issue with seller-facing bilingual fields."""
    if not isinstance(issue, dict):
        return issue
    note = explain_amazon_issue(issue, fallback_message=fallback_message)
    if note['message_en']:
        issue.setdefault('message_en', note['message_en'])
    if note['message_zh']:
        issue.setdefault('message_zh', note['message_zh'])
        issue.setdefault('seller_hint', note['seller_hint'])
    if note['display_message']:
        issue.setdefault('display_message', note['display_message'])
    return issue
