from amazon.mapper import FieldMapper, SEARCH_TERM_BYTE_LIMIT
from config import reload_config
import config as config_module
from core.search_term_utils import count_search_term_bytes, dedup_search_terms, truncate_search_terms
from core.template_service import evaluate_template_families, extract_source_keyword, validate_product_type_candidate
from web.app import _prepare_submission_products, _resolve_ai_public_image_url, _resolve_ai_status_from_result


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


def test_mapper_prefers_ai_secondary_image_when_public_base_is_configured(monkeypatch):
    monkeypatch.setenv('OUTPUT_IMAGE_PUBLIC_BASE', 'https://cdn.example.com/amazon-images')
    reload_config()

    mapper = FieldMapper()
    row = {
        'SKU': 'SKU-2',
        'item_name': 'Demo Title',
        'other_image_url_1': 'https://example.com/original-sub.jpg',
        'AI副图2路径': r'C:\tmp\sub_2.jpg',
    }
    col_map = {
        'sku': 'SKU',
        'title': 'item_name',
        'image_2': 'other_image_url_1',
    }

    product = mapper.map_excel_row(row, col_map)

    assert product['other_image_1'] == 'https://cdn.example.com/amazon-images/sub_2.jpg'


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


def test_mapper_builds_item_depth_width_height(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    attrs = mapper.build_listing_attributes({
        'sku': 'SKU-3A',
        'title': 'Demo Rack',
        'item_length': '12.6',
        'item_width': '9.4',
        'item_height': '11.0',
        'dimension_unit': 'inches',
    })

    assert attrs['item_depth_width_height'][0]['depth']['value'] == 12.6
    assert attrs['item_depth_width_height'][0]['width']['unit'] == 'inches'
    assert attrs['item_depth_width_height'][0]['height']['value'] == 11.0
    assert 'item_width_height' not in attrs


def test_mapper_sets_gtin_exemption_boolean(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    attrs = mapper.build_listing_attributes({
        'sku': 'SKU-3B',
        'title': 'Demo Rack',
        'product_identity_mode': 'gtin_exemption',
        'supplier_declared_has_product_identifier_exemption': 'True',
    })

    assert attrs['supplier_declared_has_product_identifier_exemption'][0]['value'] is True


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


def test_mapper_builds_variant_relationship_attributes_for_parent_and_child(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()

    parent_body = mapper.build_put_body({
        'sku': 'PARENT-1',
        'title': 'Demo Parent',
        'product_type': 'SHIRT',
        'brand': 'Demo',
        'parentage_level': 'parent',
        'variation_theme': 'COLOR_NAME',
    })
    child_attrs = mapper.build_listing_attributes({
        'sku': 'CHILD-RED',
        'title': 'Demo Child Red',
        'product_type': 'SHIRT',
        'brand': 'Demo',
        'parentage_level': 'child',
        'parent_sku': 'PARENT-1',
        'variation_theme': 'COLOR_NAME',
        'color': 'Red',
        'price': '19.99',
        'quantity': '5',
    })

    assert parent_body['requirements'] == 'LISTING_PRODUCT_ONLY'
    assert parent_body['attributes']['parentage_level'][0]['value'] == 'parent'
    assert parent_body['attributes']['variation_theme'][0]['name'] == 'COLOR_NAME'
    assert 'child_parent_sku_relationship' not in parent_body['attributes']
    assert 'purchasable_offer' not in parent_body['attributes']
    assert 'fulfillment_availability' not in parent_body['attributes']
    assert 'condition_type' not in parent_body['attributes']

    assert child_attrs['parentage_level'][0]['value'] == 'child'
    assert child_attrs['child_parent_sku_relationship'][0]['parent_sku'] == 'PARENT-1'
    assert child_attrs['variation_theme'][0]['name'] == 'COLOR_NAME'


def test_mapper_builds_pants_specific_bottoms_size_and_closure(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    attrs = mapper.build_listing_attributes({
        'sku': 'PANTS-1',
        'title': 'Demo Pants',
        'product_type': 'PANTS',
        'brand': 'Generic',
        'bottoms_size': 'M',
        'closure': 'drawstring',
        'rise': 'mid_rise',
    })

    assert attrs['bottoms_size'][0]['size_system'] == 'as1'
    assert attrs['bottoms_size'][0]['size_class'] == 'alpha'
    assert attrs['bottoms_size'][0]['size'] == 'm'
    assert attrs['closure'][0]['type'][0]['value'] == 'Drawstring'
    assert attrs['rise'][0]['style'][0]['value'] == 'Mid Rise'


def test_mapper_variant_validation_relaxes_gtin_for_parent_but_enforces_child_relationship(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    monkeypatch.setattr(FieldMapper, '_load_schema_fields', lambda self, product_type: None)

    parent_validation = mapper.validate_required_fields({
        'sku': 'PARENT-2',
        'title': 'Demo Parent',
        'product_type': 'SHIRT',
        'parentage_level': 'parent',
        'variation_theme': 'SIZE_NAME',
    })
    child_validation = mapper.validate_required_fields({
        'sku': 'CHILD-2',
        'title': 'Demo Child',
        'product_type': 'SHIRT',
        'parentage_level': 'child',
        'variation_theme': 'SIZE_NAME',
        'price': '21.99',
        'main_image_url': 'https://example.com/child.jpg',
    })

    assert parent_validation['valid'] is True
    assert not any('GTIN免码' in msg or '商品标识' in msg for msg in parent_validation['errors'])
    assert child_validation['valid'] is False
    assert any('parent_sku' in msg for msg in child_validation['errors'])
    assert any('颜色或尺寸' in msg for msg in child_validation['errors'])


def test_mapper_variant_validation_accepts_pants_bottoms_size(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    monkeypatch.setattr(FieldMapper, '_load_schema_fields', lambda self, product_type: None)

    validation = mapper.validate_required_fields({
        'sku': 'PANTS-CHILD',
        'title': 'Demo Pants Child',
        'product_type': 'PANTS',
        'parentage_level': 'child',
        'parent_sku': 'PANTS-PARENT',
        'variation_theme': 'SIZE_NAME',
        'bottoms_size': 'L',
        'price': '21.99',
        'product_identity_mode': 'gtin_exemption',
    })

    assert validation['valid'] is True
    assert not any('颜色或尺寸' in msg for msg in validation['errors'])


def test_prepare_submission_products_auto_includes_parent_and_blocks_missing_parent(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()

    mapper = FieldMapper()
    col_map = {
        'sku': 'SKU',
        'title': 'item_name',
        'product_type': 'product_type',
        'parentage_level': 'parentage_level',
        'parent_sku': 'parent_sku',
        'variation_theme': 'variation_theme',
        'color': 'color',
    }
    rows = [
        {
            'SKU': 'PARENT-3',
            'item_name': 'Parent Shirt',
            'product_type': 'SHIRT',
            'parentage_level': 'parent',
            'variation_theme': 'COLOR_NAME',
        },
        {
            'SKU': 'CHILD-3-RED',
            'item_name': 'Child Shirt Red',
            'product_type': 'SHIRT',
            'parentage_level': 'child',
            'parent_sku': 'PARENT-3',
            'variation_theme': 'COLOR_NAME',
            'color': 'Red',
        },
        {
            'SKU': 'CHILD-ORPHAN',
            'item_name': 'Orphan Child',
            'product_type': 'SHIRT',
            'parentage_level': 'child',
            'parent_sku': 'MISSING-PARENT',
            'variation_theme': 'COLOR_NAME',
            'color': 'Blue',
        },
    ]

    ordered, blockers = _prepare_submission_products(rows, col_map, ['CHILD-3-RED', 'CHILD-ORPHAN'], mapper)

    assert [product['sku'] for product in ordered] == ['PARENT-3', 'CHILD-3-RED']
    assert any(item['sku'] == 'CHILD-ORPHAN' and '父体' in item['message'] for item in blockers)


def test_evaluate_template_families_detects_theme_mismatch_and_duplicate_children():
    template_definition = {
        'variation_mode': 'variation',
        'product_type': 'SHIRT',
    }
    col_map = {
        'sku': 'SKU',
        'product_type': 'product_type',
        'parentage_level': 'parentage_level',
        'parent_sku': 'parent_sku',
        'variation_theme': 'variation_theme',
        'color': 'color',
        'size': 'size',
    }
    rows = [
        {'SKU': 'PARENT-A', 'product_type': 'SHIRT', 'parentage_level': 'parent', 'variation_theme': 'COLOR_SIZE'},
        {'SKU': 'CHILD-A-RED-S', 'product_type': 'SHIRT', 'parentage_level': 'child', 'parent_sku': 'PARENT-A', 'variation_theme': 'COLOR_NAME', 'color': 'Red', 'size': 'S'},
        {'SKU': 'CHILD-A-RED-S-2', 'product_type': 'SHIRT', 'parentage_level': 'child', 'parent_sku': 'PARENT-A', 'variation_theme': 'COLOR_SIZE', 'color': 'Red', 'size': 'S'},
    ]

    issues = evaluate_template_families(rows, template_definition, col_map=col_map)

    assert any('variation_theme' in msg for msg in issues['PARENT-A'])
    assert any('variation_theme' in msg for msg in issues['CHILD-A-RED-S'])
    assert any('变体组合重复' in msg for msg in issues['CHILD-A-RED-S'])
    assert any('变体组合重复' in msg for msg in issues['CHILD-A-RED-S-2'])


def test_evaluate_template_families_accepts_pants_bottoms_size():
    template_definition = {
        'variation_mode': 'variation',
        'product_type': 'PANTS',
    }
    col_map = {
        'sku': 'SKU',
        'product_type': 'product_type',
        'parentage_level': 'parentage_level',
        'parent_sku': 'parent_sku',
        'variation_theme': 'variation_theme',
        'bottoms_size': 'bottoms_size',
    }
    rows = [
        {'SKU': 'PARENT-PANTS', 'product_type': 'PANTS', 'parentage_level': 'parent', 'variation_theme': 'SIZE_NAME'},
        {'SKU': 'CHILD-PANTS-L', 'product_type': 'PANTS', 'parentage_level': 'child', 'parent_sku': 'PARENT-PANTS', 'variation_theme': 'SIZE_NAME', 'bottoms_size': 'L'},
    ]

    issues = evaluate_template_families(rows, template_definition, col_map=col_map)

    assert issues == {}


def test_extract_source_keyword_does_not_treat_numeric_slug_as_valid_title():
    info = extract_source_keyword(source_url='https://detail.1688.com/offer/968357855919.html')
    assert info['query'] == ''
    assert info['warning']


def test_validate_product_type_candidate_blocks_numeric_ids():
    result = validate_product_type_candidate(product_type='968357855919', marketplace='US', variation_mode='single')
    assert result['usable'] is False
    assert '商品ID' in result['message']


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


# ===== Group 1: 搜索词合规测试 =====

def test_count_search_term_bytes_excludes_spaces_and_punctuation():
    assert count_search_term_bytes("hello world foo") == 13  # 5+5+3, 空格不计入
    assert count_search_term_bytes("hello, world; foo.") == 13  # 标点也不计入
    assert count_search_term_bytes("") == 0
    assert count_search_term_bytes("  ") == 0


def test_count_search_term_bytes_multibyte():
    assert count_search_term_bytes("café") == 5  # c=1, a=1, f=1, é=2
    assert count_search_term_bytes("東京") == 6  # 每个 CJK 字符 3 字节


def test_dedup_removes_title_words():
    result = dedup_search_terms(
        "portable water bottle travel hydration",
        title="Portable Stainless Steel Water Bottle",
        bullets=[],
    )
    # "portable", "water", "bottle" 已在标题中
    assert "portable" not in result.lower()
    assert "water" not in result.lower()
    assert "bottle" not in result.lower()
    assert "travel" in result.lower()
    assert "hydration" in result.lower()


def test_dedup_removes_bullet_words():
    result = dedup_search_terms(
        "durable lightweight compact",
        title="Steel Bottle",
        bullets=["Durable construction", "Lightweight design"],
    )
    assert "durable" not in result.lower()
    assert "lightweight" not in result.lower()
    assert "compact" in result.lower()


def test_dedup_case_insensitive():
    result = dedup_search_terms("PREMIUM quality", title="Premium Quality Product")
    assert "premium" not in result.lower()
    assert "quality" not in result.lower()


def test_truncate_respects_byte_limit():
    terms = "aaa bbb ccc ddd eee fff ggg"  # 各 3 字节
    result = truncate_search_terms(terms, 10)
    # 10 字节能放 3 个词 (3+3+3=9)，第 4 个会超限
    assert count_search_term_bytes(result) <= 10
    assert len(result.split()) == 3


def test_truncate_multibyte_limit():
    terms = "café résumé naïve"  # café=5, résumé=8, naïve=6
    result = truncate_search_terms(terms, 13)
    assert count_search_term_bytes(result) <= 13
    assert "café" in result  # 5 <= 13
    assert "résumé" in result  # 5+8=13 <= 13
    assert "naïve" not in result  # 5+8+6=19 > 13


def test_mapper_search_term_jp_500_limit(monkeypatch):
    """日本站搜索词限制 500 字节，不是 250。"""
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()
    mapper = FieldMapper('A1VC38T7YXB528')  # JP
    product = {
        'sku': 'TEST-JP',
        'title': 'Test Product',
        'brand': 'TestBrand',
        'keywords': 'a ' * 260,  # ~260 字节单词，超过 250 但低于 500
    }
    result = mapper.validate_required_fields(product)
    # JP 限制 500，260 字节应该通过
    kw_errors = [e for e in result['errors'] if '搜索词' in e]
    assert not kw_errors


def test_mapper_search_term_over_limit_is_error(monkeypatch):
    """搜索词超限应该是 error（valid=False），不是 warning。"""
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()
    mapper = FieldMapper('ATVPDKIKX0DER')  # US, limit=250
    product = {
        'sku': 'TEST-US',
        'title': 'Test Product',
        'brand': 'TestBrand',
        'keywords': 'abcdefghij ' * 30,  # 每词 10 字节 * 30 = 300 字节，超过 250
    }
    result = mapper.validate_required_fields(product)
    assert not result['valid']
    assert any('搜索词' in e and '失效' in e for e in result['errors'])


# ===== Group 2: 标题合规测试 =====

from core.title_validation import find_banned_characters, find_duplicate_words, fix_title, validate_title


def test_find_banned_chars_detects_exclamation():
    result = find_banned_characters("Amazing Product! Buy Now$")
    chars = [v['char'] for v in result]
    assert '!' in chars
    assert '$' in chars


def test_find_banned_chars_allows_brand():
    result = find_banned_characters("Ca$h Brand Premium Bottle", brand="Ca$h Brand")
    # $ 在品牌名范围内，应被豁免
    assert len(result) == 0


def test_find_banned_chars_empty():
    assert find_banned_characters("") == []
    assert find_banned_characters("Normal Title Here") == []


def test_find_duplicate_words_detects_triple():
    dupes = find_duplicate_words("Premium Steel Premium Water Premium Bottle")
    words = [d['word'] for d in dupes]
    assert 'premium' in words
    assert dupes[0]['count'] == 3


def test_find_duplicate_words_allows_double():
    dupes = find_duplicate_words("Premium Water Premium Bottle")
    # 恰好 2 次，不应报错
    assert len(dupes) == 0


def test_find_duplicate_words_exempts_prepositions():
    dupes = find_duplicate_words("Bag for Travel for Work for Gym")
    # "for" 是介词，豁免
    assert len(dupes) == 0


def test_fix_title_removes_banned_chars():
    fixed, changes = fix_title("Amazing! Product$ Here")
    assert '!' not in fixed
    assert '$' not in fixed
    assert len(changes) > 0


def test_fix_title_removes_excess_duplicates():
    fixed, changes = fix_title("Blue Steel Blue Water Blue Bottle")
    # "blue" 出现 3 次，应去掉 1 次
    assert fixed.lower().count('blue') == 2
    assert len(changes) > 0


def test_validate_title_combined():
    result = validate_title("Premium! Premium Premium Bottle$")
    assert not result['valid']
    assert len(result['banned_chars']) == 2  # ! and $
    assert len(result['duplicate_words']) == 1  # premium x3
    assert result['suggested_fix'] is not None


def test_mapper_validates_title_banned_chars(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()
    mapper = FieldMapper('ATVPDKIKX0DER')
    product = {
        'sku': 'TEST',
        'title': 'Great Product! Buy Now$',
        'brand': 'TestBrand',
    }
    result = mapper.validate_required_fields(product)
    assert not result['valid']
    assert any('禁止字符' in e for e in result['errors'])


def test_mapper_validates_title_duplicate_words(monkeypatch):
    monkeypatch.delenv('OUTPUT_IMAGE_PUBLIC_BASE', raising=False)
    reload_config()
    mapper = FieldMapper('ATVPDKIKX0DER')
    product = {
        'sku': 'TEST',
        'title': 'Steel Steel Steel Bottle',
        'brand': 'TestBrand',
    }
    result = mapper.validate_required_fields(product)
    assert any('steel' in w and '3' in w for w in result['warnings'])
