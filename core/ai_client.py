"""
统一 AI 客户端。

支持两类协议：
1. OpenAI 兼容接口
2. Gemini generateContent 接口

文本与图片配置完全独立，可分别设置 base_url / endpoint / api_key / model。
"""
import base64
import logging
import math
import re
from typing import Optional

import httpx
from openai import OpenAI

from config import get_config

logger = logging.getLogger(__name__)

GEMINI_PROTOCOL = "gemini_generate_content"
OPENAI_RESPONSES_PROTOCOL = "openai_responses"
OPENAI_TEXT_PROTOCOL = "openai_chat_completions"
OPENAI_IMAGE_PROTOCOL = "openai_images"


def _build_endpoint_url(base_url: str, endpoint_template: str, model: str) -> str:
    base = str(base_url or "").rstrip("/")
    path = str(endpoint_template or "").strip()
    if not path:
        raise ValueError("未配置接口路径")
    # 容错: 用户把具体模型名写成了 {gemini-...} 时，按 {model} 处理
    path = re.sub(r"\{([^{}]+)\}", lambda m: "{model}" if m.group(1) != "model" else m.group(0), path)
    if "{model}" in path:
        path = path.format(model=model)
    if not path.startswith("/"):
        path = "/" + path
    return f"{base}{path}"


def _build_openai_client(api_key: str, base_url: str) -> OpenAI:
    cfg = get_config()
    return OpenAI(
        api_key=api_key,
        base_url=base_url,
        timeout=cfg.OPENAI_TIMEOUT,
        max_retries=cfg.OPENAI_MAX_RETRIES,
    )


def _image_endpoint_url(base_url: str, endpoint_template: str, model: str, operation: str) -> str:
    endpoint = str(endpoint_template or "").strip()
    if operation == "edit":
        if "/images/edits" not in endpoint.lower():
            if "/images/generations" in endpoint.lower():
                endpoint = re.sub(r"/images/generations\b", "/images/edits", endpoint, flags=re.IGNORECASE)
            else:
                endpoint = "/v1/images/edits"
    elif not endpoint:
        endpoint = "/v1/images/generations"
    return _build_endpoint_url(base_url, endpoint, model)


def _build_gemini_headers(api_key: str) -> dict:
    return {
        "Authorization": f"Bearer {api_key}",
        "x-goog-api-key": api_key,
        "Content-Type": "application/json",
        "Accept": "application/json",
    }


def _extract_text_content(response_data: dict) -> str:
    """兼容 Gemini 与 OpenAI 风格响应，提取文本内容。"""
    candidates = response_data.get("candidates", [])
    if candidates:
        texts = []
        for candidate in candidates:
            content = candidate.get("content", {})
            for part in content.get("parts", []) or []:
                text = part.get("text")
                if text:
                    texts.append(str(text).strip())
        if texts:
            return "\n".join(t for t in texts if t).strip()

    choices = response_data.get("choices", [])
    if choices:
        content = choices[0].get("message", {}).get("content", "")
        if isinstance(content, str):
            return content.strip()

    return ""


def _gemini_generate_content(
    *,
    api_key: str,
    base_url: str,
    endpoint_template: str,
    model: str,
    contents: list,
    generation_config: Optional[dict] = None,
) -> dict:
    cfg = get_config()
    url = _build_endpoint_url(base_url, endpoint_template, model)
    payload = {"contents": contents}
    if generation_config:
        payload["generationConfig"] = generation_config

    with httpx.Client(timeout=cfg.OPENAI_TIMEOUT) as client:
        response = client.post(url, headers=_build_gemini_headers(api_key), json=payload)
        response.raise_for_status()
        return response.json()


def _extract_responses_text(response_data: dict) -> str:
    """提取 Responses API 的文本输出，兼容代理返回的常见结构。"""
    output_text = response_data.get("output_text")
    if isinstance(output_text, str) and output_text.strip():
        return output_text.strip()

    texts = []
    for item in response_data.get("output", []) or []:
        for content in item.get("content", []) or []:
            text = content.get("text")
            if text:
                texts.append(str(text).strip())
    if texts:
        return "\n".join(text for text in texts if text).strip()

    return _extract_text_content(response_data)


def _openai_responses(
    *,
    api_key: str,
    base_url: str,
    endpoint_template: str,
    model: str,
    system_prompt: str,
    prompt: str,
    temperature: float,
    max_tokens: int,
) -> dict:
    cfg = get_config()
    url = _build_endpoint_url(base_url, endpoint_template, model)
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    payload = {
        "model": model,
        "input": [
            {
                "role": "system",
                "content": [{"type": "input_text", "text": system_prompt}],
            },
            {
                "role": "user",
                "content": [{"type": "input_text", "text": prompt}],
            },
        ],
        "temperature": temperature,
        "max_output_tokens": max_tokens,
    }

    with httpx.Client(timeout=cfg.OPENAI_TIMEOUT) as client:
        response = client.post(url, headers=headers, json=payload)
        response.raise_for_status()
        return response.json()


AMAZON_SYSTEM_PROMPT = (
    "You are an expert Amazon product listing copywriter and SEO specialist. "
    "You produce content that is original, conversion-focused, and compliant with Amazon's listing policies. "
    "Always return ONLY the requested content — no explanations, no markdown code blocks, no extra commentary. "
    "Follow all character/byte limits precisely. "
    "When given a target marketplace language, write entirely in that language."
)


def ai_text(
    prompt: str,
    temperature: float = 0.7,
    max_tokens: int = 2000,
    raise_on_error: bool = False,
    system_prompt: str = "",
) -> str:
    """
    AI 文本生成。

    根据配置自动走 OpenAI Responses / chat.completions 或 Gemini generateContent。
    system_prompt 为空时使用默认 Amazon 专家角色。
    """
    cfg = get_config()
    system = system_prompt or AMAZON_SYSTEM_PROMPT
    try:
        if cfg.AI_TEXT_PROTOCOL == OPENAI_RESPONSES_PROTOCOL:
            result = _openai_responses(
                api_key=cfg.AI_TEXT_API_KEY,
                base_url=cfg.AI_TEXT_API_BASE,
                endpoint_template=cfg.AI_TEXT_ENDPOINT_TEMPLATE,
                model=cfg.AI_TEXT_MODEL,
                system_prompt=system,
                prompt=prompt,
                temperature=temperature,
                max_tokens=max_tokens,
            )
            return _extract_responses_text(result)

        if cfg.AI_TEXT_PROTOCOL == GEMINI_PROTOCOL:
            combined_prompt = f"{system}\n\n{prompt}"
            result = _gemini_generate_content(
                api_key=cfg.AI_TEXT_API_KEY,
                base_url=cfg.AI_TEXT_API_BASE,
                endpoint_template=cfg.AI_TEXT_ENDPOINT_TEMPLATE,
                model=cfg.AI_TEXT_MODEL,
                contents=[{"parts": [{"text": combined_prompt}]}],
                generation_config={
                    "temperature": temperature,
                    "maxOutputTokens": max_tokens,
                },
            )
            return _extract_text_content(result)

        client = _build_openai_client(cfg.AI_TEXT_API_KEY, cfg.AI_TEXT_API_BASE)
        response = client.chat.completions.create(
            model=cfg.AI_TEXT_MODEL,
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": prompt},
            ],
            temperature=temperature,
            max_tokens=max_tokens,
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        logger.error(f"AI文本生成失败: {e}")
        if raise_on_error:
            raise
        return ""


def ai_image_edit(image_data: bytes, prompt: str, size: str = "auto") -> Optional[str]:
    """
    AI 图片编辑。

    对 Gemini generateContent 而言，这里会发送「文字提示 + 原图 inline_data」。
    """
    return _ai_image_edit(image_data, prompt, size=size, raise_on_error=False)


def _ai_image_edit(image_data: bytes, prompt: str, size: str = "auto",
                   reference_image_data: Optional[bytes] = None,
                   raise_on_error: bool = False) -> Optional[str]:
    cfg = get_config()

    if cfg.AI_IMAGE_PROTOCOL == GEMINI_PROTOCOL:
        encoded = base64.b64encode(image_data).decode("utf-8")
        generation_config = {"responseModalities": ["IMAGE"]}
        if size and size != "auto":
            aspect_ratio = _size_to_aspect_ratio(size)
            if aspect_ratio:
                generation_config["imageConfig"] = {"aspectRatio": aspect_ratio}

        parts = [
            {"text": prompt},
            {"inline_data": {"mime_type": "image/png", "data": encoded}},
        ]
        # 附带参考图（用于风格/背景参考）
        if reference_image_data:
            ref_encoded = base64.b64encode(reference_image_data).decode("utf-8")
            parts.insert(1, {"text": "Reference image for the desired background/style:"})
            parts.insert(2, {"inline_data": {"mime_type": "image/png", "data": ref_encoded}})
            parts.insert(3, {"text": "Now edit the product image above to match this reference style. Keep the product itself unchanged."})

        try:
            result = _gemini_generate_content(
                api_key=cfg.AI_IMAGE_API_KEY,
                base_url=cfg.AI_IMAGE_API_BASE,
                endpoint_template=cfg.AI_IMAGE_ENDPOINT_TEMPLATE,
                model=cfg.AI_IMAGE_MODEL,
                contents=[{"parts": parts}],
                generation_config=generation_config,
            )
            extracted = _extract_image_base64(result)
            if not extracted and raise_on_error:
                raise ValueError("AI图片编辑未返回图片数据")
            return extracted
        except httpx.HTTPStatusError as e:
            body = e.response.text[:500] if e.response else ""
            logger.error(f"AI图片编辑HTTP错误 {e.response.status_code}: {body}")
            if raise_on_error:
                raise
            return None
        except Exception as e:
            logger.error(f"AI图片编辑失败: {e}")
            if raise_on_error:
                raise
            return None

    url = _image_endpoint_url(
        cfg.AI_IMAGE_API_BASE,
        cfg.AI_IMAGE_ENDPOINT_TEMPLATE,
        cfg.AI_IMAGE_MODEL,
        "edit",
    )
    headers = {
        "Authorization": f"Bearer {cfg.AI_IMAGE_API_KEY}",
        "Accept": "application/json",
    }
    # 构建 multipart files — 支持多张图片输入
    files_list = [
        ("image[]", ("product.png", image_data, "image/png")),
    ]
    if reference_image_data:
        files_list.append(("image[]", ("reference.png", reference_image_data, "image/png")))
        prompt = f"{prompt}\nUse the second image as a style/background reference. Keep the product from the first image unchanged."

    form_data = {
        "prompt": prompt,
        "model": cfg.AI_IMAGE_MODEL,
        "n": "1",
        "size": size,
    }

    try:
        with httpx.Client(timeout=cfg.OPENAI_TIMEOUT) as client:
            response = client.post(url, headers=headers, files=files_list, data=form_data)
            response.raise_for_status()
            result = response.json()
            extracted = _extract_image_base64(result)
            if not extracted and raise_on_error:
                raise ValueError("AI图片编辑未返回图片数据")
            return extracted
    except httpx.HTTPStatusError as e:
        body = e.response.text[:500] if e.response else ""
        logger.error(f"AI图片编辑HTTP错误 {e.response.status_code}: {body}")
        if raise_on_error:
            raise
        return None
    except Exception as e:
        logger.error(f"AI图片编辑失败: {e}")
        if raise_on_error:
            raise
        return None


def ai_image_edit_url(image_url: str, prompt: str, size: str = "auto",
                      reference_image_url: str = '',
                      raise_on_error: bool = False) -> Optional[str]:
    """先下载图片（和可选的参考图），再走图片编辑。"""
    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
            ),
        }
        with httpx.Client(timeout=30, follow_redirects=True) as client:
            resp = client.get(image_url, headers=headers)
            resp.raise_for_status()
            if len(resp.content) < 1000:
                logger.warning(f"下载的图片数据过小 ({len(resp.content)} bytes)")
                if raise_on_error:
                    raise ValueError(f"下载的图片数据过小 ({len(resp.content)} bytes)")
                return None

            # 下载参考图（如有）
            reference_data = None
            if reference_image_url and reference_image_url.strip().startswith('http'):
                try:
                    ref_resp = client.get(reference_image_url.strip(), headers=headers)
                    ref_resp.raise_for_status()
                    if len(ref_resp.content) >= 1000:
                        reference_data = ref_resp.content
                        logger.info(f"  📎 已下载参考图 ({len(reference_data)} bytes)")
                    else:
                        logger.warning(f"参考图数据过小，已忽略 ({len(ref_resp.content)} bytes)")
                except Exception as ref_err:
                    logger.warning(f"参考图下载失败，继续不使用参考图: {ref_err}")

            return _ai_image_edit(resp.content, prompt, size=size,
                                  reference_image_data=reference_data,
                                  raise_on_error=raise_on_error)
    except Exception as e:
        logger.error(f"AI图片编辑(URL模式)失败: {e}")
        if raise_on_error:
            raise
        return None


def ai_image_generate(prompt: str, size: str = "1024x1024") -> Optional[str]:
    """纯文本生成图片。"""
    return _ai_image_generate(prompt, size=size, raise_on_error=False)


def _ai_image_generate(prompt: str, size: str = "1024x1024", raise_on_error: bool = False) -> Optional[str]:
    """纯文本生成图片。"""
    cfg = get_config()

    if cfg.AI_IMAGE_PROTOCOL == GEMINI_PROTOCOL:
        generation_config = {"responseModalities": ["IMAGE"]}
        aspect_ratio = _size_to_aspect_ratio(size)
        if aspect_ratio:
            generation_config["imageConfig"] = {"aspectRatio": aspect_ratio}

        try:
            result = _gemini_generate_content(
                api_key=cfg.AI_IMAGE_API_KEY,
                base_url=cfg.AI_IMAGE_API_BASE,
                endpoint_template=cfg.AI_IMAGE_ENDPOINT_TEMPLATE,
                model=cfg.AI_IMAGE_MODEL,
                contents=[{"parts": [{"text": prompt}]}],
                generation_config=generation_config,
            )
            extracted = _extract_image_base64(result)
            if not extracted and raise_on_error:
                raise ValueError("AI图片生成未返回图片数据")
            return extracted
        except httpx.HTTPStatusError as e:
            body = e.response.text[:500] if e.response else ""
            logger.error(f"AI图片生成HTTP错误 {e.response.status_code}: {body}")
            if raise_on_error:
                raise
            return None
        except Exception as e:
            logger.error(f"AI图片生成失败: {e}")
            if raise_on_error:
                raise
            return None

    url = _image_endpoint_url(
        cfg.AI_IMAGE_API_BASE,
        cfg.AI_IMAGE_ENDPOINT_TEMPLATE,
        cfg.AI_IMAGE_MODEL,
        "generate",
    )
    headers = {
        "Authorization": f"Bearer {cfg.AI_IMAGE_API_KEY}",
        "Content-Type": "application/json",
        "Accept": "application/json",
    }
    body = {
        "prompt": prompt,
        "model": cfg.AI_IMAGE_MODEL,
        "n": 1,
        "size": size,
    }

    try:
        with httpx.Client(timeout=cfg.OPENAI_TIMEOUT) as client:
            response = client.post(url, headers=headers, json=body)
            response.raise_for_status()
            result = response.json()
            extracted = _extract_image_base64(result)
            if not extracted and raise_on_error:
                raise ValueError("AI图片生成未返回图片数据")
            return extracted
    except httpx.HTTPStatusError as e:
        body_text = e.response.text[:500] if e.response else ""
        logger.error(f"AI图片生成HTTP错误 {e.response.status_code}: {body_text}")
        if raise_on_error:
            raise
        return None
    except Exception as e:
        logger.error(f"AI图片生成失败: {e}")
        if raise_on_error:
            raise
        return None


def _size_to_aspect_ratio(size: str) -> Optional[str]:
    if not size or size == "auto":
        return None

    match = re.match(r"^\s*(\d+)\s*x\s*(\d+)\s*$", str(size), re.IGNORECASE)
    if not match:
        return None

    width = int(match.group(1))
    height = int(match.group(2))
    if width <= 0 or height <= 0:
        return None

    divisor = math.gcd(width, height) or 1
    return f"{width // divisor}:{height // divisor}"


# ===== 响应解析 =====

def _extract_image_base64(response_data: dict) -> Optional[str]:
    """
    从响应中提取 base64 图片。

    兼容：
    1. Gemini generateContent candidates.parts.inlineData / inline_data
    2. 代理 choices.message.content 中的 data URI / base64
    3. 标准 OpenAI Images API data[0].b64_json / url
    """
    candidates = response_data.get("candidates", [])
    for candidate in candidates:
        content = candidate.get("content", {})
        for part in content.get("parts", []) or []:
            inline = part.get("inlineData") or part.get("inline_data")
            if isinstance(inline, dict):
                data = inline.get("data")
                if data:
                    return data

            text = part.get("text")
            if text:
                b64 = _parse_base64_content(text)
                if b64:
                    return b64

    choices = response_data.get("choices", [])
    if choices:
        content = choices[0].get("message", {}).get("content", "")
        if content:
            b64 = _parse_base64_content(content)
            if b64:
                return b64

    data = response_data.get("data", [])
    if data:
        item = data[0]
        if item.get("b64_json"):
            return item["b64_json"]
        if item.get("url"):
            try:
                with httpx.Client(timeout=30) as client:
                    resp = client.get(item["url"])
                    resp.raise_for_status()
                    return base64.b64encode(resp.content).decode("utf-8")
            except Exception as e:
                logger.error(f"下载生成图片失败: {e}")
                return None

    logger.warning(f"无法从响应中提取图片数据, keys={list(response_data.keys())}")
    return None


def _parse_base64_content(content: str) -> Optional[str]:
    """从字符串中解析 base64 图片数据。"""
    if not content:
        return None

    match = re.search(r"data:image/\w+;base64,([A-Za-z0-9+/=\s]+)", content)
    if match:
        return match.group(1).replace("\n", "").replace("\r", "").replace(" ", "")

    match = re.search(r"!\[.*?\]\(data:image/\w+;base64,([A-Za-z0-9+/=\s]+)\)", content)
    if match:
        return match.group(1).replace("\n", "").replace("\r", "").replace(" ", "")

    stripped = content.strip()
    if len(stripped) > 1000:
        clean = stripped.replace("\n", "").replace("\r", "")
        if re.match(r"^[A-Za-z0-9+/=]+$", clean):
            return clean

    return None
