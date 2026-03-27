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
