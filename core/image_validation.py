"""
Amazon 商品图片合规校验。

根据 Amazon 官方规则:
- 主图最长边 >= 1000px（启用缩放功能）
- 主图背景纯白 RGB(255,255,255)
- 产品需占据画面 >= 85%
- 推荐 JPEG 格式
"""
import logging
from typing import Dict, List, Optional, Tuple

from PIL import Image

logger = logging.getLogger(__name__)

MIN_LONGEST_SIDE = 1000
WHITE_THRESHOLD = 245  # 四角像素平均值高于此阈值视为白底
WHITE_CORNER_SAMPLE = 10  # 每个角采样 10x10 像素区域


def validate_image(image: Image.Image, is_main: bool = True) -> Dict:
    """校验图片是否符合 Amazon 要求。返回 {valid, issues, dimensions, format_hint}。"""
    issues: List[Dict] = []
    width, height = image.size
    longest = max(width, height)

    # 尺寸检查
    if longest < MIN_LONGEST_SIDE:
        issues.append({
            'level': 'error',
            'message': f'图片最长边 {longest}px，低于 Amazon 最低要求 {MIN_LONGEST_SIDE}px',
        })

    # 白底检查（仅主图）
    if is_main:
        bg_result = check_white_background(image)
        if not bg_result['is_white']:
            issues.append({
                'level': 'warning',
                'message': f'主图背景可能不是纯白 (四角平均亮度 {bg_result["avg_brightness"]:.0f}/255)',
            })

    # 格式建议
    fmt = (image.format or '').upper()
    format_hint = ''
    if fmt and fmt not in ('JPEG', 'JPG'):
        format_hint = f'当前格式 {fmt}，Amazon 推荐 JPEG'

    return {
        'valid': not any(i['level'] == 'error' for i in issues),
        'issues': issues,
        'dimensions': {'width': width, 'height': height},
        'format_hint': format_hint,
    }


def check_white_background(image: Image.Image) -> Dict:
    """通过采样四角像素判断背景是否为纯白。"""
    img = image.convert('RGB')
    width, height = img.size

    if width < WHITE_CORNER_SAMPLE * 2 or height < WHITE_CORNER_SAMPLE * 2:
        return {'is_white': True, 'avg_brightness': 255.0}

    s = WHITE_CORNER_SAMPLE
    corners = [
        img.crop((0, 0, s, s)),               # 左上
        img.crop((width - s, 0, width, s)),     # 右上
        img.crop((0, height - s, s, height)),   # 左下
        img.crop((width - s, height - s, width, height)),  # 右下
    ]

    total_brightness = 0
    pixel_count = 0
    for corner in corners:
        pixels = corner.tobytes()
        for i in range(0, len(pixels), 3):
            r, g, b = pixels[i], pixels[i + 1], pixels[i + 2]
            total_brightness += (r + g + b) / 3
            pixel_count += 1

    avg = total_brightness / pixel_count if pixel_count else 0

    return {
        'is_white': avg >= WHITE_THRESHOLD,
        'avg_brightness': avg,
    }


def ensure_jpeg(image: Image.Image) -> Image.Image:
    """确保图片为 RGB 模式（适合保存为 JPEG）。"""
    if image.mode in ('RGBA', 'LA', 'P'):
        background = Image.new('RGB', image.size, (255, 255, 255))
        if image.mode == 'P':
            image = image.convert('RGBA')
        background.paste(image, mask=image.split()[-1] if 'A' in image.mode else None)
        return background
    if image.mode != 'RGB':
        return image.convert('RGB')
    return image
