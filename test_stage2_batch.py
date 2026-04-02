from amazon.feeds import FeedsAPI
from amazon.mapper import FieldMapper
from stage2_pipeline import Stage2Pipeline


def test_feeds_api_uses_marketplace_region():
    feeds = FeedsAPI(auth=None, marketplace_id='A1F83G8C2ARO7P')
    assert feeds.base_url == 'https://sellingpartnerapi-eu.amazon.com'


def test_feeds_api_requests_issue_report():
    feeds = FeedsAPI(auth=None)
    payload = feeds._build_json_feed(
        items=[{'sku': 'SKU-1', 'attributes': {}, 'product_type': 'PRODUCT'}],
        seller_id='SELLER-1',
    )

    assert payload['header']['report']['includedData'] == ['issues']


def test_stage2_batch_uses_item_results(monkeypatch):
    pipeline = Stage2Pipeline()

    monkeypatch.setattr(
        pipeline.mapper,
        'validate_required_fields',
        lambda product: {'valid': True, 'errors': [], 'warnings': []},
    )
    monkeypatch.setattr(
        pipeline.listings,
        'put_listings_item',
        lambda sku, product, preview=False: {'status': 'VALID', 'issues': []},
    )
    monkeypatch.setattr(
        pipeline.mapper,
        'build_put_body',
        lambda product: {
            'productType': product.get('product_type', 'PRODUCT'),
            'requirements': 'LISTING',
            'attributes': {'item_name': [{'value': product['sku'], 'marketplace_id': 'ATVPDKIKX0DER'}]},
        },
    )
    monkeypatch.setattr(
        pipeline.feeds,
        'submit_and_wait',
        lambda **kwargs: {
            'status': 'DONE',
            'feed_id': 'feed-1',
            'item_results': [
                {'sku': 'SKU-1', 'status': 'ACCEPTED', 'issues': []},
                {'sku': 'SKU-2', 'status': 'INVALID', 'issues': [{'message': 'missing field'}]},
            ],
        },
    )

    results = pipeline._submit_batch([
        {'sku': 'SKU-1', 'product_type': 'PRODUCT'},
        {'sku': 'SKU-2', 'product_type': 'PRODUCT'},
    ])

    assert [result['sku'] for result in results] == ['SKU-1', 'SKU-2']
    assert pipeline.stats['submitted'] == 2
    assert pipeline.stats['success'] == 1
    assert pipeline.stats['failed'] == 1


def test_stage2_batch_done_without_item_results_is_not_marked_success(monkeypatch):
    pipeline = Stage2Pipeline()

    monkeypatch.setattr(
        pipeline.mapper,
        'validate_required_fields',
        lambda product: {'valid': True, 'errors': [], 'warnings': []},
    )
    monkeypatch.setattr(
        pipeline.listings,
        'put_listings_item',
        lambda sku, product, preview=False: {'status': 'VALID', 'issues': []},
    )
    monkeypatch.setattr(
        pipeline.mapper,
        'build_put_body',
        lambda product: {
            'productType': product.get('product_type', 'PRODUCT'),
            'requirements': 'LISTING',
            'attributes': {'item_name': [{'value': product['sku'], 'marketplace_id': 'ATVPDKIKX0DER'}]},
        },
    )
    monkeypatch.setattr(
        pipeline.feeds,
        'submit_and_wait',
        lambda **kwargs: {
            'status': 'DONE',
            'feed_id': 'feed-2',
            'item_results': [],
        },
    )

    results = pipeline._submit_batch([
        {'sku': 'SKU-1', 'product_type': 'PRODUCT'},
        {'sku': 'SKU-2', 'product_type': 'PRODUCT'},
    ])

    assert [result['status'] for result in results] == ['UNKNOWN', 'UNKNOWN']
    assert pipeline.stats['success'] == 0
    assert pipeline.stats['failed'] == 2


def test_stage2_batch_uses_mapper_requirements_and_product_type(monkeypatch):
    pipeline = Stage2Pipeline()
    captured = {}

    monkeypatch.setattr(
        pipeline.mapper,
        'validate_required_fields',
        lambda product: {'valid': True, 'errors': [], 'warnings': []},
    )
    monkeypatch.setattr(
        pipeline.listings,
        'put_listings_item',
        lambda sku, product, preview=False: {'status': 'VALID', 'issues': []},
    )
    monkeypatch.setattr(
        pipeline.mapper,
        'build_put_body',
        lambda product: {
            'productType': 'SHIRT',
            'requirements': 'LISTING_PRODUCT_ONLY',
            'attributes': {'item_name': [{'value': product['sku'], 'marketplace_id': 'ATVPDKIKX0DER'}]},
        },
    )

    def fake_submit_and_wait(**kwargs):
        captured['items'] = kwargs['items']
        return {
            'status': 'DONE',
            'feed_id': 'feed-parent',
            'item_results': [{'sku': 'PARENT-1', 'status': 'ACCEPTED', 'issues': []}],
        }

    monkeypatch.setattr(pipeline.feeds, 'submit_and_wait', fake_submit_and_wait)

    results = pipeline._submit_batch([{
        'sku': 'PARENT-1',
        'product_type': 'SHIRT',
        'parentage_level': 'parent',
        'variation_theme': 'COLOR_NAME',
    }])

    assert results[0]['status'] == 'ACCEPTED'
    assert captured['items'][0]['product_type'] == 'SHIRT'
    assert captured['items'][0]['requirements'] == 'LISTING_PRODUCT_ONLY'


def test_stage2_batch_filters_invalid_rows_before_feed(monkeypatch):
    pipeline = Stage2Pipeline()

    def fake_validate(product):
        if product['sku'] == 'BAD-1':
            return {'valid': False, 'errors': ['缺少标题'], 'warnings': []}
        return {'valid': True, 'errors': [], 'warnings': []}

    monkeypatch.setattr(pipeline.mapper, 'validate_required_fields', fake_validate)
    monkeypatch.setattr(
        pipeline.listings,
        'put_listings_item',
        lambda sku, product, preview=False: {'status': 'VALID', 'issues': []},
    )
    monkeypatch.setattr(
        pipeline.mapper,
        'build_put_body',
        lambda product: {
            'productType': product.get('product_type', 'PRODUCT'),
            'requirements': 'LISTING',
            'attributes': {'item_name': [{'value': product['sku'], 'marketplace_id': 'ATVPDKIKX0DER'}]},
        },
    )

    captured = {}

    def fake_submit_and_wait(**kwargs):
        captured['items'] = kwargs['items']
        return {
            'status': 'DONE',
            'feed_id': 'feed-valid',
            'item_results': [{'sku': 'GOOD-1', 'status': 'ACCEPTED', 'issues': []}],
        }

    monkeypatch.setattr(pipeline.feeds, 'submit_and_wait', fake_submit_and_wait)

    results = pipeline._submit_batch([
        {'sku': 'BAD-1', 'product_type': 'PRODUCT'},
        {'sku': 'GOOD-1', 'product_type': 'PRODUCT'},
    ])

    assert captured['items'] == [{
        'sku': 'GOOD-1',
        'product_type': 'PRODUCT',
        'requirements': 'LISTING',
        'attributes': {'item_name': [{'value': 'GOOD-1', 'marketplace_id': 'ATVPDKIKX0DER'}]},
    }]
    assert results[0]['sku'] == 'BAD-1'
    assert results[0]['status'] == 'VALIDATION_ERROR'
    assert results[1]['sku'] == 'GOOD-1'
    assert results[1]['status'] == 'ACCEPTED'


def test_stage2_individual_blocks_submit_when_preview_fails(monkeypatch):
    pipeline = Stage2Pipeline()
    calls = []

    monkeypatch.setattr(
        pipeline.mapper,
        'validate_required_fields',
        lambda product: {'valid': True, 'errors': [], 'warnings': []},
    )

    def fake_put(sku, product, preview=False):
        calls.append((sku, preview))
        if preview:
            return {'status': 'INVALID', 'issues': [{'severity': 'ERROR', 'message': '缺少字段'}]}
        return {'status': 'ACCEPTED', 'issues': []}

    monkeypatch.setattr(pipeline.listings, 'put_listings_item', fake_put)

    results = pipeline._submit_individual([{'sku': 'SKU-1', 'product_type': 'PRODUCT'}], preview_before_submit=True)

    assert calls == [('SKU-1', True)]
    assert results[0]['status'] == 'PREVIEW_INVALID'


def test_stage2_individual_counts_accepted_with_warnings_as_success(monkeypatch):
    pipeline = Stage2Pipeline()

    monkeypatch.setattr(
        pipeline.mapper,
        'validate_required_fields',
        lambda product: {'valid': True, 'errors': [], 'warnings': []},
    )
    monkeypatch.setattr('stage2_pipeline.time.sleep', lambda *_args: None)
    monkeypatch.setattr(
        pipeline.listings,
        'put_listings_item',
        lambda sku, product, preview=False: (
            {'status': 'VALID', 'issues': []}
            if preview
            else {'status': 'ACCEPTED_WITH_WARNINGS', 'issues': [{'severity': 'WARNING', 'message': 'minor'}]}
        ),
    )

    results = pipeline._submit_individual([
        {'sku': 'SKU-1', 'title': 'Demo', 'product_type': 'PRODUCT'},
    ])

    assert results[0]['status'] == 'ACCEPTED_WITH_WARNINGS'
    assert pipeline.stats['success'] == 1
    assert pipeline.stats['failed'] == 0


def test_field_mapper_accepts_s3_media_locator():
    mapper = FieldMapper()
    attrs = mapper.build_listing_attributes({
        'title': 'Demo Product',
        'brand': 'Demo Brand',
        'product_type': 'PRODUCT',
        'price': '19.99',
        'currency': 'USD',
        'quantity': '3',
        'main_image_url': 's3://demo-bucket/amazon28/US/SKU-1/main.jpg',
        'other_image_1': 's3://demo-bucket/amazon28/US/SKU-1/sub2.jpg',
    })

    assert attrs['main_product_image_locator'][0]['media_location'].startswith('s3://demo-bucket/')
    assert attrs['other_product_image_locator_1'][0]['media_location'].startswith('s3://demo-bucket/')
