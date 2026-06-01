from amazon.listings import ListingsAPI


class _FakeAuth:
    def get_headers(self):
        return {'Authorization': 'Bearer test'}


class _FakeResponse:
    def __init__(self, status_code, payload):
        self.status_code = status_code
        self._payload = payload

    def json(self):
        return self._payload


def test_put_listings_item_uses_identifiers_only_in_preview(monkeypatch):
    captured = {}

    def fake_put(url, params=None, headers=None, json=None, timeout=None):
        captured['params'] = dict(params or {})
        return _FakeResponse(200, {'status': 'ACCEPTED', 'issues': []})

    monkeypatch.setattr('amazon.listings.requests.put', fake_put)

    api = ListingsAPI(auth=_FakeAuth(), seller_id='SELLER', marketplace_id='ATVPDKIKX0DER')
    product = {'sku': 'SKU-1', 'product_type': 'PRODUCT'}

    api.put_listings_item('SKU-1', product, preview=True)
    assert captured['params']['includedData'] == 'issues,identifiers'

    api.put_listings_item('SKU-1', product, preview=False)
    assert captured['params']['includedData'] == 'issues'


def test_put_listings_item_backfills_asin_from_listing_lookup(monkeypatch):
    def fake_put(url, params=None, headers=None, json=None, timeout=None):
        return _FakeResponse(200, {'status': 'ACCEPTED', 'issues': [], 'submissionId': 'SUB-001'})

    monkeypatch.setattr('amazon.listings.requests.put', fake_put)
    monkeypatch.setattr(
        ListingsAPI,
        'get_listings_item',
        lambda self, sku: {'summaries': [{'asin': 'B00BACKFILL123'}]},
    )

    api = ListingsAPI(auth=_FakeAuth(), seller_id='SELLER', marketplace_id='ATVPDKIKX0DER')
    result = api.put_listings_item('SKU-1', {'sku': 'SKU-1', 'product_type': 'PRODUCT'}, preview=False)

    assert result['asin'] == 'B00BACKFILL123'


def test_put_listings_item_flattens_top_level_errors(monkeypatch):
    def fake_put(url, params=None, headers=None, json=None, timeout=None):
        return _FakeResponse(400, {
            'errors': [{
                'code': 'InvalidInput',
                'message': 'Unable to Retrieve Media Content',
                'attributeNames': ['main_product_image_locator'],
            }]
        })

    monkeypatch.setattr('amazon.listings.requests.put', fake_put)

    api = ListingsAPI(auth=_FakeAuth(), seller_id='SELLER', marketplace_id='ATVPDKIKX0DER')
    result = api.put_listings_item('SKU-1', {'sku': 'SKU-1', 'product_type': 'PRODUCT'}, preview=True)

    assert result['status'] == 'ERROR'
    assert result['http_status'] == 400
    assert result['issues'][0]['code'] == 'InvalidInput'
    assert result['issues'][0]['severity'] == 'ERROR'
    assert result['issues'][0]['message_en'] == 'Unable to Retrieve Media Content'
    assert 'Amazon 无法读取图片' in result['issues'][0]['message_zh']
    assert 'Unable to Retrieve Media Content（' in result['issues'][0]['display_message']


def test_put_listings_item_adds_chinese_note_for_generic_invalid_input(monkeypatch):
    def fake_put(url, params=None, headers=None, json=None, timeout=None):
        return _FakeResponse(400, {
            'errors': [{
                'code': 'InvalidInput',
                'message': 'Invalid parameters provided.',
            }]
        })

    monkeypatch.setattr('amazon.listings.requests.put', fake_put)

    api = ListingsAPI(auth=_FakeAuth(), seller_id='SELLER', marketplace_id='ATVPDKIKX0DER')
    result = api.put_listings_item('SKU-1', {'sku': 'SKU-1', 'product_type': 'PRODUCT'}, preview=True)

    issue = result['issues'][0]
    assert issue['message'] == 'Invalid parameters provided.'
    assert issue['message_en'] == 'Invalid parameters provided.'
    assert '参数无效' in issue['message_zh']
    assert issue['display_message'].startswith('Invalid parameters provided.（参数无效')


def test_put_listings_item_adds_note_for_title_image_mismatch(monkeypatch):
    def fake_put(url, params=None, headers=None, json=None, timeout=None):
        return _FakeResponse(200, {
            'status': 'INVALID',
            'issues': [{
                'severity': 'ERROR',
                'code': '100239',
                'message': "The title and the Main image that you provided on this SKU don't seem to represent the same product.",
                'attributeNames': ['item_name', 'main_product_image_locator'],
            }]
        })

    monkeypatch.setattr('amazon.listings.requests.put', fake_put)

    api = ListingsAPI(auth=_FakeAuth(), seller_id='SELLER', marketplace_id='ATVPDKIKX0DER')
    result = api.put_listings_item('SKU-1', {'sku': 'SKU-1', 'product_type': 'PRODUCT'}, preview=True)

    issue = result['issues'][0]
    assert issue['code'] == '100239'
    assert '标题和主图不匹配' in issue['message_zh']
    assert 'item_name' in issue['message_zh']


def test_submit_listings_blocks_when_preview_returns_error_status(monkeypatch):
    api = ListingsAPI(auth=_FakeAuth(), seller_id='SELLER', marketplace_id='ATVPDKIKX0DER')
    calls = []

    monkeypatch.setattr(
        api.mapper,
        'validate_required_fields',
        lambda product: {'valid': True, 'errors': [], 'warnings': []},
    )

    def fake_put(sku, product, preview=False):
        calls.append((sku, preview))
        if preview:
            return {'status': 'ERROR', 'issues': [{'severity': 'ERROR', 'message': 'timeout'}]}
        return {'status': 'ACCEPTED', 'issues': []}

    monkeypatch.setattr(api, 'put_listings_item', fake_put)

    results = api.submit_listings([{'sku': 'SKU-1', 'title': 'Demo', 'product_type': 'PRODUCT'}], preview_first=True, delay=0)

    assert calls == [('SKU-1', True)]
    assert results[0]['status'] == 'PREVIEW_ERROR'


def test_submit_listings_uses_resolved_asin_when_submit_response_has_none(monkeypatch):
    api = ListingsAPI(auth=_FakeAuth(), seller_id='SELLER', marketplace_id='ATVPDKIKX0DER')

    monkeypatch.setattr(
        api.mapper,
        'validate_required_fields',
        lambda product: {'valid': True, 'errors': [], 'warnings': []},
    )
    monkeypatch.setattr(
        api,
        'put_listings_item',
        lambda sku, product, preview=False: (
            {'status': 'VALID', 'issues': []}
            if preview
            else {'status': 'ACCEPTED', 'issues': [], 'submissionId': 'SUB-001'}
        ),
    )
    monkeypatch.setattr(
        api,
        'resolve_submission_asin',
        lambda sku, submit_result=None, max_attempts=2, delay=0.5: 'B00RESOLVED123',
    )

    results = api.submit_listings([{'sku': 'SKU-1', 'title': 'Demo', 'product_type': 'PRODUCT'}], preview_first=True, delay=0)

    assert results[0]['status'] == 'ACCEPTED'
    assert results[0]['asin'] == 'B00RESOLVED123'


def test_submit_listings_skips_duplicate_skus(monkeypatch):
    api = ListingsAPI(auth=_FakeAuth(), seller_id='SELLER', marketplace_id='ATVPDKIKX0DER')
    calls = []

    monkeypatch.setattr(
        api.mapper,
        'validate_required_fields',
        lambda product: {'valid': True, 'errors': [], 'warnings': []},
    )

    def fake_put(sku, product, preview=False):
        calls.append((sku, preview))
        if preview:
            return {'status': 'VALID', 'issues': []}
        return {'status': 'ACCEPTED', 'issues': [], 'submissionId': f'SUB-{sku}'}

    monkeypatch.setattr(api, 'put_listings_item', fake_put)
    monkeypatch.setattr(api, 'resolve_submission_asin', lambda *args, **kwargs: '')

    results = api.submit_listings([
        {'sku': 'SKU-1', 'title': 'Demo 1', 'product_type': 'PRODUCT'},
        {'sku': 'SKU-1', 'title': 'Demo 1 Duplicate', 'product_type': 'PRODUCT'},
    ], preview_first=True, delay=0)

    assert calls == [('SKU-1', True), ('SKU-1', False)]
    assert results[0]['status'] == 'ACCEPTED'
    assert results[1]['status'] == 'SKIPPED'
    assert '重复 SKU' in results[1]['issues'][0]['message']
