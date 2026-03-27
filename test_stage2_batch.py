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
        'build_listing_attributes',
        lambda product: {'item_name': [{'value': product['sku'], 'marketplace_id': 'ATVPDKIKX0DER'}]},
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
        'build_listing_attributes',
        lambda product: {'item_name': [{'value': product['sku'], 'marketplace_id': 'ATVPDKIKX0DER'}]},
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
