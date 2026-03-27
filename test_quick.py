from amazon.mapper import FieldMapper
from config import reload_config
import config as config_module
from web.app import _resolve_ai_public_image_url, _resolve_ai_status_from_result


def test_mapper_keeps_shipping_and_compliance_fields(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    row = {
        'SKU': 'SKU-1',
        'item_name': 'Demo Title',
        'main_image_url': 'https://example.com/original.jpg',
        'item_weight': '1.2',
        'item_weight_unit': 'pounds',
        'item_length': '10',
        'item_width': '20',
        'item_height': '30',
        'package_weight': '2.5',
        'package_weight_unit': 'ounces',
        'package_length': '11',
        'package_width': '21',
        'package_height': '31',
        'batteries_required': 'Yes',
        'batteries_included': 'No',
        'battery_type': 'AA',
    }
    col_map = {
        'sku': 'SKU',
        'title': 'item_name',
        'image_url': 'main_image_url',
        'weight': 'item_weight',
        'item_weight_unit': 'item_weight_unit',
        'item_length': 'item_length',
        'item_width': 'item_width',
        'item_height': 'item_height',
        'package_weight': 'package_weight',
        'package_weight_unit': 'package_weight_unit',
        'package_length': 'package_length',
        'package_width': 'package_width',
        'package_height': 'package_height',
        'batteries_required': 'batteries_required',
        'batteries_included': 'batteries_included',
        'battery_type': 'battery_type',
    }

    product = mapper.map_excel_row(row, col_map)

    assert product['weight'] == '1.2'
    assert product['weight_unit'] == 'pounds'
    assert product['item_length'] == '10'
    assert product['package_length'] == '11'
    assert product['batteries_required'] == 'Yes'
    assert product['battery_type'] == 'AA'

    attrs = mapper.build_listing_attributes(product)
    assert attrs['item_weight'][0]['unit'] == 'pounds'
    assert attrs['item_dimensions'][0]['length']['value'] == 10.0
    assert attrs['item_package_dimensions'][0]['length']['value'] == 11.0
    assert attrs['item_package_weight'][0]['unit'] == 'ounces'
    assert attrs['batteries_required'][0]['value'] is True
    assert attrs['batteries_included'][0]['value'] is False


def test_mapper_promotes_local_ai_image_when_public_base_is_configured(monkeypatch):
    monkeypatch.setenv('OUTPUT_IMAGE_PUBLIC_BASE', 'https://cdn.example.com/amazon-images')
    reload_config()

    mapper = FieldMapper()
    row = {
        'SKU': 'SKU-2',
        'item_name': 'Demo Title',
        'main_image_url': 'https://example.com/original.jpg',
        'AI主图路径': r'C:\tmp\main_2.jpg',
    }
    col_map = {
        'sku': 'SKU',
        'title': 'item_name',
        'image_url': 'main_image_url',
    }

    product = mapper.map_excel_row(row, col_map)

    assert product['main_image_url'] == 'https://cdn.example.com/amazon-images/main_2.jpg'
    assert product['main_image_source'] == 'ai_public'


def test_mapper_skips_invalid_dimension_values(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    attrs = mapper.build_listing_attributes({
        'sku': 'SKU-3',
        'title': 'Demo Title',
        'item_length': 'abc',
        'item_width': '20',
        'package_height': 'bad',
        'package_width': '5',
    })

    assert attrs['item_dimensions'][0]['width']['value'] == 20.0
    assert 'length' not in attrs['item_dimensions'][0]
    assert attrs['item_package_dimensions'][0]['width']['value'] == 5.0
    assert 'height' not in attrs['item_package_dimensions'][0]


def test_mapper_preserves_dynamic_schema_fields_and_battery_alias(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    row = {
        'SKU': 'SKU-4',
        'item_name': 'Demo Headphones',
        'main_image_url': 'https://example.com/headphones.jpg',
        'are_batteries_included': 'Yes',
        'connectivity_technology': 'Bluetooth',
        'headphones_form_factor': 'On Ear',
    }
    col_map = {
        'sku': 'SKU',
        'title': 'item_name',
        'image_url': 'main_image_url',
    }

    product = mapper.map_excel_row(row, col_map)

    assert product['batteries_included'] == 'Yes'
    assert product['connectivity_technology'] == 'Bluetooth'
    assert product['headphones_form_factor'] == 'On Ear'

    attrs = mapper.build_listing_attributes(product)
    assert attrs['batteries_included'][0]['value'] is True
    assert attrs['connectivity_technology'][0]['value'] == 'Bluetooth'
    assert attrs['headphones_form_factor'][0]['value'] == 'On Ear'


def test_mapper_uses_schema_required_fields_when_available(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    import amazon.schema_manager as schema_manager

    monkeypatch.setattr(
        schema_manager,
        'fetch_schema',
        lambda product_type, marketplace='US', sp_client=None: {
            'product_type': product_type,
            'marketplace': marketplace,
        },
    )
    monkeypatch.setattr(
        schema_manager,
        'parse_schema',
        lambda schema_data: {
            'required_fields': [
                {
                    'name': 'connectivity_technology',
                    'title': 'Connectivity Technology',
                    'group': 'product_details',
                },
                {
                    'name': 'headphones_form_factor',
                    'title': 'Headphones Form Factor',
                    'group': 'product_details',
                },
            ],
            'optional_fields': [],
            'enum_fields': {},
            'field_groups': {},
        },
    )

    mapper = FieldMapper()
    product = {
        'sku': 'SKU-5',
        'title': 'Demo Headphones',
        'price': '29.99',
        'product_type': 'HEADPHONES',
        'main_image_url': 'https://example.com/headphones.jpg',
    }

    validation = mapper.validate_required_fields(product)

    assert validation['valid'] is False
    assert any('connectivity_technology' in msg for msg in validation['errors'])
    assert any('headphones_form_factor' in msg for msg in validation['errors'])
    assert validation['schema_required_missing'] == [
        {
            'name': 'connectivity_technology',
            'title': 'Connectivity Technology',
            'group': 'product_details',
        },
        {
            'name': 'headphones_form_factor',
            'title': 'Headphones Form Factor',
            'group': 'product_details',
        },
    ]


def test_mapper_excludes_helper_alias_fields_from_attributes(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    row = {
        'SKU': 'SKU-6',
        'item_name': 'Demo Bottle',
        'main_image_url': 'https://example.com/bottle.jpg',
        'brand_name': 'Helper Brand',
        'color_name': 'Black',
        'product_id_type': 'UPC',
        'standard_price': '19.99',
        'source_url': 'https://source.example.com/item',
        'connectivity_technology': 'Bluetooth',
    }
    col_map = {
        'sku': 'SKU',
        'title': 'item_name',
        'image_url': 'main_image_url',
    }

    product = mapper.map_excel_row(row, col_map)
    attrs = mapper.build_listing_attributes(product)

    assert 'brand_name' not in product
    assert 'color_name' not in product
    assert 'product_id_type' not in product
    assert 'standard_price' not in product
    assert 'source_url' not in product
    assert attrs['connectivity_technology'][0]['value'] == 'Bluetooth'
    assert 'brand_name' not in attrs
    assert 'color_name' not in attrs
    assert 'product_id_type' not in attrs
    assert 'standard_price' not in attrs
    assert 'source_url' not in attrs


def test_web_helpers_distinguish_public_ai_image_and_failed_status():
    assert _resolve_ai_public_image_url({'AI主图URL': 'https://cdn.example.com/a.jpg'}) == 'https://cdn.example.com/a.jpg'
    assert _resolve_ai_public_image_url({'AI主图URL': '/api/output-image/a.jpg'}) == ''

    assert _resolve_ai_status_from_result({'ai_title': 'New title'}, 'rewrite') == 'completed'
    assert _resolve_ai_status_from_result({}, 'rewrite') == 'failed'
    assert _resolve_ai_status_from_result({'ai_main_image': '/api/output-image/a.jpg'}, 'image') == 'completed'
    assert _resolve_ai_status_from_result({}, 'image') == 'failed'


def test_mapper_blocks_internal_code_submission_and_ignores_exemption_identifier(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()

    internal_product = {
        'sku': 'SKU-7',
        'title': 'Demo Product',
        'price': '19.99',
        'main_image_url': 'https://example.com/a.jpg',
        'upc': '00890123400029',
        'product_identity_mode': 'internal_code',
        'external_product_id_type': 'UPC',
    }
    internal_validation = mapper.validate_required_fields(internal_product)
    assert internal_validation['valid'] is False
    assert any('内部码' in msg for msg in internal_validation['errors'])

    exemption_product = {
        'sku': 'SKU-8',
        'title': 'Demo Product',
        'price': '19.99',
        'main_image_url': 'https://example.com/a.jpg',
        'upc': '00890123400029',
        'product_identity_mode': 'gtin_exemption',
        'external_product_id_type': 'UPC',
    }
    attrs = mapper.build_listing_attributes(exemption_product)
    exemption_validation = mapper.validate_required_fields(exemption_product)
    assert 'externally_assigned_product_identifier' not in attrs
    assert exemption_validation['valid'] is True
    assert any('忽略已填写的 UPC/EAN/GTIN' in msg for msg in exemption_validation['warnings'])


def test_mapper_requires_valid_real_gtin_when_identity_mode_is_real(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    product = {
        'sku': 'SKU-9',
        'title': 'Demo Product',
        'price': '19.99',
        'main_image_url': 'https://example.com/a.jpg',
        'upc': 'ABC123',
        'product_identity_mode': 'real_gtin',
        'external_product_id_type': 'UPC',
    }

    validation = mapper.validate_required_fields(product)

    assert validation['valid'] is False
    assert any('必须为纯数字' in msg for msg in validation['errors'])


def test_reload_config_initializes_instance(monkeypatch):
    original_instance = config_module._config_instance
    try:
        config_module._config_instance = None

        cfg = reload_config()

        assert cfg is not None
        assert cfg is config_module._config_instance
    finally:
        config_module._config_instance = original_instance
        if original_instance is not None:
            original_instance.reload()


def test_config_supports_separate_text_and_image_settings(monkeypatch):
    monkeypatch.setattr(config_module, 'load_dotenv', lambda override=True: None)
    monkeypatch.setenv('AI_API_KEY', '')
    monkeypatch.setenv('AI_API_BASE', '')
    monkeypatch.setenv('AI_TEXT_API_KEY', 'text-key')
    monkeypatch.setenv('AI_TEXT_API_BASE', 'api.kk666.online')
    monkeypatch.setenv('AI_TEXT_ENDPOINT_TEMPLATE', '/v1beta/models/{model}:generateContent')
    monkeypatch.setenv('AI_TEXT_MODEL', 'gemini-3.1-flash-lite-preview')
    monkeypatch.setenv('AI_IMAGE_API_KEY', 'image-key')
    monkeypatch.setenv('AI_IMAGE_API_BASE', 'api.kk666.online')
    monkeypatch.setenv('AI_IMAGE_ENDPOINT_TEMPLATE', '/v1beta/models/{model}:generateContent')
    monkeypatch.setenv('AI_IMAGE_MODEL', 'gemini-3.1-flash-image-preview')

    cfg = config_module.Config()

    assert cfg.AI_TEXT_API_KEY == 'text-key'
    assert cfg.AI_IMAGE_API_KEY == 'image-key'
    assert cfg.AI_TEXT_API_BASE == 'https://api.kk666.online'
    assert cfg.AI_IMAGE_API_BASE == 'https://api.kk666.online'
    assert cfg.AI_TEXT_PROTOCOL == 'gemini_generate_content'
    assert cfg.AI_IMAGE_PROTOCOL == 'gemini_generate_content'
