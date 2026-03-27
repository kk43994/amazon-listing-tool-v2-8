from core.ai_client import _build_endpoint_url, _extract_image_base64, _extract_text_content


def test_extract_text_content_from_gemini_response():
    response = {
        'candidates': [
            {
                'content': {
                    'parts': [
                        {'text': '第一段'},
                        {'text': '第二段'},
                    ]
                }
            }
        ]
    }

    assert _extract_text_content(response) == '第一段\n第二段'


def test_extract_image_base64_from_gemini_response():
    response = {
        'candidates': [
            {
                'content': {
                    'parts': [
                        {'inlineData': {'mimeType': 'image/png', 'data': 'ZmFrZS1pbWFnZQ=='}},
                    ]
                }
            }
        ]
    }

    assert _extract_image_base64(response) == 'ZmFrZS1pbWFnZQ=='


def test_extract_image_base64_from_gemini_snake_case_response():
    response = {
        'candidates': [
            {
                'content': {
                    'parts': [
                        {'inline_data': {'mime_type': 'image/png', 'data': 'c25ha2UtY2FzZQ=='}},
                    ]
                }
            }
        ]
    }

    assert _extract_image_base64(response) == 'c25ha2UtY2FzZQ=='


def test_build_endpoint_url_tolerates_braced_model_name():
    url = _build_endpoint_url(
        'https://api.kk666.online',
        '/v1beta/models/{gemini-3.1-flash-lite-preview}:generateContent',
        'gemini-3.1-flash-lite-preview',
    )

    assert url == 'https://api.kk666.online/v1beta/models/gemini-3.1-flash-lite-preview:generateContent'
