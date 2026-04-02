"""
Amazon SP-API 字段映射器 V2
按照SP-API实际格式构建attributes JSON
支持Product Type动态Schema
"""
import logging
import os
import re
from typing import Dict, List, Optional
from urllib.parse import quote

from config import get_config

logger = logging.getLogger(__name__)

# Amazon US marketplace
US_MARKETPLACE = "ATVPDKIKX0DER"

# 各站点Marketplace ID
MARKETPLACE_IDS = {
    'US': 'ATVPDKIKX0DER',
    'CA': 'A2EUQ1WTGCTBG2',
    'MX': 'A1AM78C64UM0Y8',
    'UK': 'A1F83G8C2ARO7P',
    'DE': 'A1PA6795UKMFR9',
    'FR': 'A13V1IB3VIYZZH',
    'IT': 'APJ6JRA9NG5V4',
    'ES': 'A1RKKUPIHCS9HS',
    'JP': 'A1VC38T7YXB528',
    'AU': 'A39IBJ37TRP1C6',
}

# SP-API endpoints per region
ENDPOINTS = {
    'NA': 'https://sellingpartnerapi-na.amazon.com',
    'EU': 'https://sellingpartnerapi-eu.amazon.com',
    'FE': 'https://sellingpartnerapi-fe.amazon.com',
}

MARKETPLACE_REGION = {
    'ATVPDKIKX0DER': 'NA', 'A2EUQ1WTGCTBG2': 'NA', 'A1AM78C64UM0Y8': 'NA',
    'A1F83G8C2ARO7P': 'EU', 'A1PA6795UKMFR9': 'EU', 'A13V1IB3VIYZZH': 'EU',
    'APJ6JRA9NG5V4': 'EU', 'A1RKKUPIHCS9HS': 'EU',
    'A1VC38T7YXB528': 'FE', 'A39IBJ37TRP1C6': 'FE',
}

MARKETPLACE_CODE_BY_ID = {marketplace_id: code for code, marketplace_id in MARKETPLACE_IDS.items()}

# 各站点默认币种
MARKETPLACE_CURRENCY = {
    'ATVPDKIKX0DER': 'USD',   # US
    'A2EUQ1WTGCTBG2': 'CAD',  # CA
    'A1AM78C64UM0Y8': 'MXN',  # MX
    'A1F83G8C2ARO7P': 'GBP',  # UK
    'A1PA6795UKMFR9': 'EUR',  # DE
    'A13V1IB3VIYZZH': 'EUR',  # FR
    'APJ6JRA9NG5V4': 'EUR',   # IT
    'A1RKKUPIHCS9HS': 'EUR',  # ES
    'A1VC38T7YXB528': 'JPY',  # JP
    'A39IBJ37TRP1C6': 'AUD',  # AU
}

# 各站点主要语言（用于 AI 提示词语言选择和 language_tag）
MARKETPLACE_LANGUAGE = {
    'ATVPDKIKX0DER': 'en_US',
    'A2EUQ1WTGCTBG2': 'en_CA',
    'A1AM78C64UM0Y8': 'es_MX',
    'A1F83G8C2ARO7P': 'en_GB',
    'A1PA6795UKMFR9': 'de_DE',
    'A13V1IB3VIYZZH': 'fr_FR',
    'APJ6JRA9NG5V4': 'it_IT',
    'A1RKKUPIHCS9HS': 'es_ES',
    'A1VC38T7YXB528': 'ja_JP',
    'A39IBJ37TRP1C6': 'en_AU',
    'A21TJRUUN4KGV': 'en_IN',
}

# 搜索词字节限制（仅统计单词字节，空格和标点不计入）
# 超限会导致整条搜索词全部失效
SEARCH_TERM_BYTE_LIMIT = {
    'ATVPDKIKX0DER': 250,   # US
    'A2EUQ1WTGCTBG2': 250,  # CA
    'A1AM78C64UM0Y8': 250,  # MX
    'A1F83G8C2ARO7P': 250,  # UK
    'A1PA6795UKMFR9': 250,  # DE
    'A13V1IB3VIYZZH': 250,  # FR
    'APJ6JRA9NG5V4': 250,   # IT
    'A1RKKUPIHCS9HS': 250,  # ES
    'A1VC38T7YXB528': 500,  # JP
    'A39IBJ37TRP1C6': 250,  # AU
    'A21TJRUUN4KGV': 200,   # IN
}

GENERIC_ATTRIBUTE_EXCLUDE = {
    'sku', 'title', 'brand', 'description', 'keywords', 'price', 'list_price', 'currency',
    'quantity', 'condition_type', 'fulfillment_channel', 'product_type', 'item_type',
    'upc', 'ean', 'gtin', 'asin', 'color', 'size', 'material', 'department',
    'target_gender', 'age_range', 'manufacturer', 'model_number', 'model_name',
    'country_of_origin', 'weight', 'weight_unit', 'item_length', 'item_width',
    'item_height', 'dimension_unit', 'package_weight', 'package_weight_unit',
    'package_length', 'package_width', 'package_height', 'batteries_required',
    'batteries_included', 'battery_type', 'number_of_batteries',
    'battery_cell_composition', 'lithium_battery_packaging',
    'lithium_battery_energy_content', 'lithium_battery_weight',
    'hazmat_declaration', 'supplier_declared_dg_hz_regulation',
    'supplier_declared_has_product_identifier_exemption',
    'parent_sku', 'parentage_level', 'variation_theme',
    'capacity', 'special_feature', 'included_components', 'care_instructions',
    'unit_count', 'number_of_items', 'main_image_url', 'main_image_source',
    'ai_main_image_path',
    'brand_name', 'color_name', 'size_name', 'material_type', 'generic_keywords',
    'item_name', 'product_description', 'item_type_keyword', 'product_id',
    'product_id_type', 'external_product_id_type', 'product_identity_mode',
    'standard_price', 'parent_child', 'relationship_type',
    'submission_id', 'submission_route',
    'other_image_url_1', 'other_image_url_2', 'other_image_url_3', 'other_image_url_4',
    'other_image_url_5', 'other_image_url_6', 'other_image_url_7', 'other_image_url_8',
    'item_weight', 'item_weight_unit_of_measure', 'item_dimensions_unit_of_measure',
    'target_audience_keywords', 'source_asin', 'source_color', 'source_marketplace',
    'source_parent_asin', 'source_price_usd', 'source_size', 'source_url',
}

GENERIC_BOOL_FIELDS = {
    'assembly_required', 'batteries_included', 'batteries_required', 'inflatable',
    'is_dishwasher_safe', 'is_electric', 'is_foldable', 'is_gift_message_available',
    'is_gift_wrap_available', 'is_hypoallergenic', 'is_microwave_safe',
    'is_organic', 'is_oven_safe',
}

NUMERIC_HINT_FIELDS = {
    'amperage', 'map_price', 'number_of_players', 'page_yield', 'servings_per_container',
    'voltage', 'wattage',
}


class FieldMapper:
    """SP-API字段映射器 V2"""

    def __init__(self, marketplace_id: str = US_MARKETPLACE):
        self.config = get_config()
        self.marketplace_id = marketplace_id
        self.region = MARKETPLACE_REGION.get(marketplace_id, 'NA')
        self.endpoint = ENDPOINTS.get(self.region, ENDPOINTS['NA'])
        self.default_currency = MARKETPLACE_CURRENCY.get(marketplace_id, 'USD')
        self.language_tag = MARKETPLACE_LANGUAGE.get(marketplace_id, 'en_US')
        self._schema_fields_cache = {}

    def build_listing_attributes(self, product: Dict) -> Dict:
        """
        从标准化产品数据构建SP-API attributes JSON

        按照Amazon SP-API putListingsItem的attributes格式:
        {"attribute_name": [{"value": "xxx", "marketplace_id": "ATVPDKIKX0DER"}]}
        """
        attrs = {}
        mp = self.marketplace_id
        parentage_level = self._normalize_parentage_level(product.get('parentage_level'))
        variation_theme = self._clean_value(product.get('variation_theme'))
        parent_sku = self._clean_value(product.get('parent_sku'))
        is_parent = parentage_level == 'parent'
        is_child = parentage_level == 'child'

        # === Product Identity ===
        self._set_text_attr(attrs, 'item_name', product.get('title'), mp)
        self._set_text_attr(attrs, 'brand', product.get('brand'), mp)
        self._set_text_attr(attrs, 'manufacturer', product.get('manufacturer'), mp)
        self._set_text_attr(attrs, 'model_number', product.get('model_number'), mp)
        self._set_text_attr(attrs, 'model_name', product.get('model_name'), mp)
        self._set_text_attr(attrs, 'item_type_keyword', product.get('item_type'), mp)

        # 产品标识(UPC/EAN/GTIN)
        identity_mode = self._resolve_product_identity_mode(product)
        if not is_parent and identity_mode == 'real_gtin' and product.get('upc'):
            attrs['externally_assigned_product_identifier'] = [{
                'type': 'upc',
                'value': str(product['upc']),
                'marketplace_id': mp,
            }]
        elif not is_parent and identity_mode == 'real_gtin' and product.get('ean'):
            attrs['externally_assigned_product_identifier'] = [{
                'type': 'ean',
                'value': str(product['ean']),
                'marketplace_id': mp,
            }]
        elif not is_parent and identity_mode == 'real_gtin' and product.get('gtin'):
            attrs['externally_assigned_product_identifier'] = [{
                'type': 'gtin',
                'value': str(product['gtin']),
                'marketplace_id': mp,
            }]

        # ASIN匹配(跟卖)
        if product.get('asin'):
            attrs['merchant_suggested_asin'] = [{
                'value': product['asin'],
                'marketplace_id': mp,
            }]

        if parentage_level:
            attrs['parentage_level'] = [{
                'value': parentage_level,
                'marketplace_id': mp,
            }]
            if is_child and parent_sku:
                attrs['child_parent_sku_relationship'] = [{
                    'child_relationship_type': 'variation',
                    'parent_sku': parent_sku,
                    'marketplace_id': mp,
                }]
        if variation_theme:
            attrs['variation_theme'] = [{
                'name': variation_theme,
                'marketplace_id': mp,
            }]

        # === Product Details ===
        self._set_text_attr(attrs, 'product_description', product.get('description'), mp)

        # Bullet Points (5条独立)
        bullets = self._extract_bullets(product)
        if bullets:
            attrs['bullet_point'] = [
                {'value': bp, 'marketplace_id': mp}
                for bp in bullets
            ]

        # 搜索关键词
        if product.get('keywords'):
            attrs['generic_keyword'] = [{
                'value': product['keywords'],
                'marketplace_id': mp,
            }]

        product_type = str(product.get('product_type', '') or '').strip().upper()

        # 属性
        self._set_text_attr(attrs, 'color', product.get('color'), mp, lang='en_US')
        if product_type == 'PANTS' and product.get('bottoms_size'):
            size_text = str(product.get('bottoms_size') or product.get('size') or '').strip()
            if size_text:
                attrs['bottoms_size'] = [{
                    'marketplace_id': mp,
                    'size_system': 'as1',
                    'size_class': 'alpha',
                    'size': self._normalize_bottoms_size_value(size_text),
                }]
        else:
            self._set_text_attr(attrs, 'size', product.get('size'), mp, lang='en_US')
        self._set_text_attr(attrs, 'material', product.get('material'), mp, lang='en_US')
        self._set_text_attr(attrs, 'department', product.get('department'), mp, lang='en_US')
        self._set_text_attr(attrs, 'target_gender', product.get('target_gender'), mp)
        self._set_text_attr(attrs, 'age_range_description', product.get('age_range_description') or product.get('age_range'), mp)
        self._set_text_attr(attrs, 'special_feature', product.get('special_feature'), mp, lang='en_US')
        self._set_text_attr(attrs, 'target_audience_keyword', product.get('target_audience'), mp)
        self._set_text_attr(attrs, 'subject_keyword', product.get('subject_keywords'), mp)
        self._set_text_attr(attrs, 'care_instructions', product.get('care_instructions'), mp, lang='en_US')
        self._set_text_attr(attrs, 'included_components', product.get('included_components'), mp, lang='en_US')

        if product.get('number_of_items'):
            try:
                attrs['number_of_items'] = [{
                    'value': int(float(product['number_of_items'])),
                    'marketplace_id': mp,
                }]
            except (ValueError, TypeError):
                pass

        if product.get('unit_count'):
            unit_count = self._parse_measurement(product['unit_count'], fallback_unit='count')
            if unit_count:
                attrs['unit_count'] = [{
                    'value': int(unit_count['value']) if float(unit_count['value']).is_integer() else unit_count['value'],
                    'type': {
                        'language_tag': 'en_US',
                        'value': 'Count' if unit_count['unit'] == 'count' else str(unit_count['unit']).replace('_', ' ').title(),
                    },
                    'marketplace_id': mp,
                }]

        if product.get('capacity'):
            capacity = self._parse_measurement(product['capacity'], fallback_unit='fluid_ounces')
            if capacity:
                attrs['capacity'] = [{**capacity, 'marketplace_id': mp}]

        # === Images ===
        if self._is_media_locator(product.get('main_image_url')):
            attrs['main_product_image_locator'] = [{
                'media_location': product['main_image_url'],
                'marketplace_id': mp,
            }]
        for i in range(1, 9):
            img_key = f'other_image_{i}'
            if self._is_media_locator(product.get(img_key)):
                attrs[f'other_product_image_locator_{i}'] = [{
                    'media_location': product[img_key],
                    'marketplace_id': mp,
                }]

        # === Offer ===
        if not is_parent and product.get('price'):
            try:
                price_val = float(product['price'])
                currency = product.get('currency') or self.default_currency
                attrs['purchasable_offer'] = [{
                    'marketplace_id': mp,
                    'currency': currency,
                    'our_price': [{
                        'schedule': [{
                            'value_with_tax': price_val
                        }]
                    }],
                }]
            except (ValueError, TypeError):
                logger.warning(f"  ⚠️ 无效价格: {product.get('price')}")

        if not is_parent and product.get('list_price'):
            try:
                attrs['list_price'] = [{
                    'value': float(product['list_price']),
                    'currency': product.get('currency') or self.default_currency,
                    'marketplace_id': mp,
                }]
            except (ValueError, TypeError):
                logger.warning(f"  ⚠️ 无效划线价: {product.get('list_price')}")

        # 库存
        if not is_parent and product.get('quantity'):
            try:
                qty = int(product['quantity'])
                channel = product.get('fulfillment_channel', 'DEFAULT')
                attrs['fulfillment_availability'] = [{
                    'fulfillment_channel_code': channel,
                    'quantity': qty,
                    'marketplace_id': mp,
                }]
            except (ValueError, TypeError):
                pass

        # 商品状态
        if not is_parent:
            condition = self._normalize_condition_type(product.get('condition_type', 'new_new'))
            attrs['condition_type'] = [{
                'value': condition,
                'marketplace_id': mp,
            }]

        # === Shipping ===
        dims = self._build_dimensions(product)
        if dims:
            attrs['item_dimensions'] = [dims]
        depth_width_height = self._build_item_depth_width_height(product)
        if depth_width_height:
            attrs['item_depth_width_height'] = [depth_width_height]

        pkg_dims = self._build_package_dimensions(product)
        if pkg_dims:
            attrs['item_package_dimensions'] = [pkg_dims]

        if product.get('weight'):
            try:
                attrs['item_weight'] = [{
                    'unit': product.get('weight_unit', 'grams'),
                    'value': float(product['weight']),
                    'marketplace_id': mp,
                }]
            except (ValueError, TypeError):
                pass

        package_weight = product.get('package_weight') or product.get('weight')
        package_weight_unit = product.get('package_weight_unit') or product.get('weight_unit', 'grams')
        if package_weight:
            try:
                attrs['item_package_weight'] = [{
                    'unit': package_weight_unit,
                    'value': float(package_weight),
                    'marketplace_id': mp,
                }]
            except (ValueError, TypeError):
                pass

        # === Safety & Compliance ===
        self._set_text_attr(attrs, 'country_of_origin', product.get('country_of_origin'), mp)

        if product.get('batteries_required') is not None:
            val = self._coerce_bool(product['batteries_required'])
            attrs['batteries_required'] = [{
                'value': val,
                'marketplace_id': mp,
            }]

        if product.get('batteries_included') is not None:
            val = self._coerce_bool(product['batteries_included'])
            attrs['batteries_included'] = [{
                'value': val,
                'marketplace_id': mp,
            }]

        if product.get('supplier_declared_has_product_identifier_exemption') is not None:
            val = self._coerce_bool(product['supplier_declared_has_product_identifier_exemption'])
            attrs['supplier_declared_has_product_identifier_exemption'] = [{
                'value': val,
                'marketplace_id': mp,
            }]

        declared_dg_regulation = (
            self._clean_value(product.get('supplier_declared_dg_hz_regulation'))
            or self._clean_value(product.get('hazmat_declaration'))
        )
        if declared_dg_regulation:
            attrs['supplier_declared_dg_hz_regulation'] = [{
                'value': declared_dg_regulation,
                'marketplace_id': mp,
            }]
        elif not self._coerce_bool(product.get('batteries_required')):
            attrs['supplier_declared_dg_hz_regulation'] = [{
                'value': 'not_applicable',
                'marketplace_id': mp,
            }]

        closure_text = self._clean_value(product.get('closure'))
        if closure_text:
            attrs['closure'] = [{
                'marketplace_id': mp,
                'type': [{
                    'language_tag': 'en_US',
                    'value': self._humanize_enum_label(closure_text),
                }],
            }]

        rise_text = self._clean_value(product.get('rise'))
        if rise_text:
            attrs['rise'] = [{
                'marketplace_id': mp,
                'style': [{
                    'language_tag': 'en_US',
                    'value': self._humanize_enum_label(rise_text),
                }],
            }]

        self._apply_generic_attributes(attrs, product, mp)

        return attrs

    def build_put_body(self, product: Dict) -> Dict:
        """
        构建完整的putListingsItem请求体

        Returns:
            {
                "productType": "PRODUCT",
                "requirements": "LISTING",
                "attributes": {...}
            }
        """
        product_type = product.get('product_type', 'PRODUCT')
        parentage_level = self._normalize_parentage_level(product.get('parentage_level'))

        # 决定requirements类型
        if parentage_level == 'parent':
            requirements = 'LISTING_PRODUCT_ONLY'
        elif product.get('asin') and not product.get('title'):
            # 有ASIN没标题 = 跟卖(offer only)
            requirements = 'LISTING_OFFER_ONLY'
        else:
            requirements = 'LISTING'

        return {
            'productType': product_type,
            'requirements': requirements,
            'attributes': self.build_listing_attributes(product),
        }

    def validate_required_fields(self, product: Dict, schema_fields: Dict = None) -> Dict:
        """
        验证SP-API必填字段

        Returns:
            {
                'valid': bool,
                'errors': [必填缺失],
                'warnings': [推荐缺失],
                'info': [可选缺失]
            }
        """
        result = {
            'valid': True,
            'errors': [],
            'warnings': [],
            'info': [],
            'schema_required_missing': [],
            'product_identity_mode': '',
        }

        # 必填字段
        required = {
            'sku': 'SKU(卖家商品编号)',
            'title': '商品标题(item_name)',
        }
        parentage_level = self._normalize_parentage_level(product.get('parentage_level'))
        is_parent = parentage_level == 'parent'
        is_child = parentage_level == 'child'

        # 上架必填(有offer时)
        if not is_parent and product.get('price'):
            required['price'] = '价格'

        for field, label in required.items():
            if not product.get(field):
                result['errors'].append(f"缺少必填字段: {label}")
                result['valid'] = False

        # 强烈推荐
        recommended = {
            'brand': '品牌(brand)',
            'description': '商品描述',
            'keywords': '搜索关键词',
        }
        for field, label in recommended.items():
            if not product.get(field):
                result['warnings'].append(f"建议填写: {label}")

        if not is_parent:
            identity_validation = self._validate_product_identity(product)
            result['product_identity_mode'] = identity_validation['mode']
            if identity_validation['mode_missing']:
                result['warnings'].append(identity_validation['mode_missing'])
            for message in identity_validation['errors']:
                result['errors'].append(message)
                result['valid'] = False
            for message in identity_validation['warnings']:
                result['warnings'].append(message)
            for message in identity_validation['info']:
                result['info'].append(message)

        if parentage_level and parentage_level not in ('parent', 'child'):
            result['errors'].append("变体商品的 parentage_level 必须是 parent 或 child")
            result['valid'] = False
        if (product.get('parent_sku') or product.get('variation_theme')) and not parentage_level:
            result['errors'].append("填写 parent_sku 或 variation_theme 时，必须同时声明 parentage_level")
            result['valid'] = False
        if parentage_level:
            if not product.get('variation_theme'):
                result['errors'].append("变体商品必须填写 variation_theme")
                result['valid'] = False
            if is_child and not product.get('parent_sku'):
                result['errors'].append("子体商品必须填写 parent_sku")
                result['valid'] = False
            if is_child and not self._has_variation_dimension(product):
                result['errors'].append("子体至少需要填写颜色或尺寸中的一个变体维度")
                result['valid'] = False
            if is_parent and product.get('parent_sku'):
                result['warnings'].append("父体通常不需要填写 parent_sku，当前值会在提交时忽略")

        # 检查标题
        if product.get('title'):
            if len(product['title']) > 200:
                result['errors'].append(f"标题超过200字符({len(product['title'])}字符)")
                result['valid'] = False
            from core.title_validation import find_banned_characters, find_duplicate_words
            brand = product.get('brand', '')
            banned = find_banned_characters(product['title'], brand=brand)
            for item in banned:
                result['errors'].append(
                    f"标题含有Amazon禁止字符 '{item['char']}' (位置{item['position']})"
                )
                result['valid'] = False
            for dup in find_duplicate_words(product['title']):
                result['warnings'].append(
                    f"标题中 '{dup['word']}' 出现{dup['count']}次(Amazon规定最多2次)"
                )

        # 检查bullet points（长度 + 禁止内容 + emoji）
        bullets = self._extract_bullets(product)
        from core.bullet_validation import validate_bullets
        bp_result = validate_bullets(bullets)
        for issue in bp_result['issues']:
            if issue['level'] == 'error':
                result['errors'].append(issue['message'])
                result['valid'] = False
            else:
                result['warnings'].append(issue['message'])

        # 检查搜索词字节数（Amazon 规则：仅统计单词字节，空格和标点不计入；超限整条失效）
        if product.get('keywords'):
            from core.search_term_utils import count_search_term_bytes
            kw_bytes = count_search_term_bytes(str(product['keywords']))
            kw_limit = SEARCH_TERM_BYTE_LIMIT.get(self.marketplace_id, 250)
            if kw_bytes > kw_limit:
                result['errors'].append(
                    f"搜索词超过{kw_limit}字节限制({kw_bytes}字节)，将导致整条搜索词失效"
                )
                result['valid'] = False

        # 检查币种
        user_currency = product.get('currency', '').strip()
        if user_currency and user_currency != self.default_currency:
            marketplace_code = MARKETPLACE_CODE_BY_ID.get(self.marketplace_id, self.marketplace_id)
            result['warnings'].append(
                f"当前站点({marketplace_code})默认币种为 {self.default_currency}，"
                f"您填写的是 {user_currency}，请确认是否正确"
            )

        # 检查图片
        main_image_url = product.get('main_image_url')
        if not main_image_url:
            result['warnings'].append("缺少主图URL")
        elif not self._is_media_locator(main_image_url):
            result['errors'].append("主图媒体地址必须是 http(s) 或 s3:// 地址")
            result['valid'] = False
        elif product.get('main_image_source') == 'original_fallback_local_ai':
            result['warnings'].append(
                "AI主图仅保存在本地，当前回退为原始远程图；如需提交 AI 图，请启用媒体存储并完成上传"
            )

        # 检查合规字段
        if self._coerce_bool(product.get('batteries_required')):
            if not product.get('battery_type'):
                result['warnings'].append("标记需要电池但未填写电池类型(battery_type)")
            if not product.get('batteries_included'):
                result['warnings'].append("标记需要电池但未说明是否含电池(batteries_included)")

        schema_fields = schema_fields or self._load_schema_fields(product.get('product_type', ''))
        if schema_fields:
            attrs = self.build_listing_attributes(product)
            for field in schema_fields.get('required_fields', []):
                attr_name = str(field.get('name', '')).strip()
                if not attr_name or attr_name in attrs:
                    continue
                title = field.get('title') or attr_name
                group = field.get('group') or 'other'
                result['schema_required_missing'].append({
                    'name': attr_name,
                    'title': title,
                    'group': group,
                })
                result['errors'].append(
                    f"缺少类目必填字段: {title} ({attr_name})"
                )
                result['valid'] = False

        return result

    def _normalize_parentage_level(self, value) -> str:
        text = self._clean_value(value).lower()
        if text in ('parent', 'child'):
            return text
        return text

    def _variation_size_value(self, product: Dict) -> str:
        return self._clean_value(product.get('size') or product.get('bottoms_size'))

    def _has_variation_dimension(self, product: Dict) -> bool:
        return bool(self._clean_value(product.get('color')) or self._variation_size_value(product))

    def map_excel_row(self, row: Dict, col_map: Dict) -> Dict:
        """
        将Excel行数据映射为标准产品数据

        优先使用AI生成的值，原始值作为fallback
        """
        product = {}

        def pick_value(*sources: str) -> str:
            for source in sources:
                value = self._get_row_value(row, col_map, source)
                if value:
                    return value
            return ""

        # 映射规则: (标准字段, [可能的来源列，优先级从高到低])
        field_sources = {
            'sku': ['sku'],
            'title': ['AI标题', 'title'],  # AI优先
            'brand': ['brand'],
            'description': ['AI商品描述', 'description'],  # AI优先
            'keywords': ['AI搜索关键词', 'keywords'],  # AI优先
            'price': ['price'],
            'list_price': ['list_price'],
            'currency': ['currency'],
            'quantity': ['quantity'],
            'condition_type': ['condition_type'],
            'fulfillment_channel': ['fulfillment_channel'],
            'product_type': ['product_type'],
            'item_type': ['item_type_keyword'],
            'upc': ['upc'],
            'ean': ['ean'],
            'gtin': ['gtin'],
            'external_product_id_type': ['external_product_id_type', 'product_id_type'],
            'asin': ['asin', 'source_asin'],
            'product_identity_mode': ['product_identity_mode'],
            'color': ['color'],
            'size': ['size'],
            'material': ['material'],
            'department': ['department'],
            'target_gender': ['target_gender'],
            'target_audience': ['AI目标受众', 'target_audience_keywords', 'target_audience'],
            'subject_keywords': ['AI主题关键词', 'subject_keywords'],
            'age_range': ['age_range'],
            'age_range_description': ['age_range_description', 'age_range'],
            'manufacturer': ['manufacturer'],
            'model_number': ['model_number', 'part_number', 'model_name'],
            'model_name': ['model_name', 'model_number', 'part_number'],
            'country_of_origin': ['country_of_origin'],
            'weight': ['weight'],
            'weight_unit': ['item_weight_unit', 'item_weight_unit_of_measure'],
            'item_length': ['item_length'],
            'item_width': ['item_width'],
            'item_height': ['item_height'],
            'dimension_unit': ['dimension_unit', 'item_dimensions_unit_of_measure'],
            'package_weight': ['package_weight'],
            'package_weight_unit': ['package_weight_unit'],
            'package_length': ['package_length'],
            'package_width': ['package_width'],
            'package_height': ['package_height'],
            'batteries_required': ['batteries_required', 'are_batteries_required'],
            'batteries_included': ['batteries_included', 'are_batteries_included'],
            'battery_type': ['battery_type'],
            'number_of_batteries': ['number_of_batteries'],
            'battery_cell_composition': ['battery_cell_composition'],
            'lithium_battery_packaging': ['lithium_battery_packaging'],
            'lithium_battery_energy_content': ['lithium_battery_energy_content'],
            'lithium_battery_weight': ['lithium_battery_weight'],
            'supplier_declared_dg_hz_regulation': ['supplier_declared_dg_hz_regulation', 'hazmat_declaration'],
            'hazmat_declaration': ['hazmat_declaration', 'supplier_declared_dg_hz_regulation'],
            'supplier_declared_has_product_identifier_exemption': ['supplier_declared_has_product_identifier_exemption'],
            'parent_sku': ['parent_sku'],
            'parentage_level': ['parentage_level'],
            'variation_theme': ['variation_theme'],
            'bottoms_size': ['bottoms_size', 'size'],
            'capacity': ['capacity'],
            'special_feature': ['AI特殊功能', 'special_feature'],
            'included_components': ['included_components'],
            'care_instructions': ['care_instructions'],
            'unit_count': ['unit_count'],
            'number_of_items': ['number_of_items'],
        }

        for standard_field, sources in field_sources.items():
            value = pick_value(*sources)
            if value:
                product[standard_field] = value

        if str(product.get('product_type', '') or '').strip().upper() == 'PANTS':
            if not product.get('bottoms_size') and product.get('size'):
                product['bottoms_size'] = product['size']

        main_image_url, main_image_source, ai_main_image_path = self._resolve_main_image(row, col_map)
        if main_image_url:
            product['main_image_url'] = main_image_url
        if main_image_source:
            product['main_image_source'] = main_image_source
        if ai_main_image_path:
            product['ai_main_image_path'] = ai_main_image_path

        # Bullet Points特殊处理
        for i in range(1, 6):
            ai_key = f'AI卖点{i}'
            if row.get(ai_key):
                product[f'bullet_point_{i}'] = row[ai_key]
            else:
                col_name = col_map.get(f'bullet_point_{i}')
                if col_name and row.get(col_name):
                    product[f'bullet_point_{i}'] = str(row[col_name]).strip()

        # 副图
        for i in range(2, 10):
            image_url = self._resolve_additional_image(row, col_map, i)
            if image_url:
                product[f'other_image_{i-1}'] = image_url

        self._collect_dynamic_fields(row, product)

        return product

    def get_display_fields(self, row: Dict) -> Dict:
        """
        从Excel行提取前端表格需要显示的字段

        统一从多种可能的列名中取值，返回标准化的展示字段。
        适用于前端表格渲染、数据预览等场景。

        Args:
            row: Excel原始行数据(dict)

        Returns:
            标准化的展示字段字典
        """
        return {
            'sku': row.get('sku', '') or row.get('SKU', '') or row.get('seller_sku', '') or '',
            'title': row.get('title', '') or row.get('item_name', '') or '',
            'main_image_url': (row.get('main_image_url', '') or row.get('main_image', '')
                               or row.get('image_url', '') or ''),
            'price': row.get('price', '') or row.get('standard_price', '') or '',
            'brand': row.get('brand', '') or row.get('brand_name', '') or '',
            'parent_sku': row.get('parent_sku', '') or '',
            'parentage_level': row.get('parentage_level', '') or row.get('relationship_type', '') or '',
            'variation_theme': row.get('variation_theme', '') or '',
            'color': row.get('color', '') or row.get('color_name', '') or '',
            'size': row.get('size', '') or row.get('size_name', '') or '',
            'ai_title': row.get('AI_title', '') or row.get('ai_title', '') or row.get('AI标题', '') or '',
            'ai_status': row.get('ai_status', '') or '',
            'submit_status': row.get('submit_status', '') or '',
        }

    # ===== 内部方法 =====

    def _set_text_attr(self, attrs: Dict, name: str, value, mp: str,
                       lang: str = None):
        """设置文本属性"""
        if not value:
            return
        entry = {'value': str(value).strip(), 'marketplace_id': mp}
        if lang:
            entry['language_tag'] = lang
        attrs[name] = [entry]

    def _extract_bullets(self, product: Dict) -> List[str]:
        """提取Bullet Points列表"""
        bullets = []
        # 方式1: 独立字段
        for i in range(1, 6):
            bp = product.get(f'bullet_point_{i}', '')
            if bp:
                bullets.append(str(bp).strip())

        # 方式2: 合并字段
        if not bullets and product.get('bullet_points'):
            raw = str(product['bullet_points'])
            for line in raw.split('\n'):
                line = line.strip().lstrip('•-* ')
                if line:
                    bullets.append(line)

        return bullets[:5]

    def _apply_generic_attributes(self, attrs: Dict, product: Dict, mp: str):
        """将已映射但未显式处理的类目字段按简单属性透传到 attributes。"""
        for name, raw_value in product.items():
            if name in attrs or name in GENERIC_ATTRIBUTE_EXCLUDE:
                continue
            if name.startswith('bullet_point_') or name.startswith('other_image_'):
                continue
            if not re.match(r'^[a-z][a-z0-9_]*$', str(name or '')):
                continue

            entry = self._build_generic_attr_entry(name, raw_value, mp)
            if entry is not None:
                attrs[name] = [entry]

    def _build_generic_attr_entry(self, name: str, value, marketplace_id: str) -> Optional[Dict]:
        text = self._clean_value(value)
        if not text:
            return None

        entry = {'marketplace_id': marketplace_id}
        if name in GENERIC_BOOL_FIELDS or name.startswith(('is_', 'are_', 'has_')):
            entry['value'] = self._coerce_bool(value)
            return entry

        numeric_value = self._coerce_number(value)
        if numeric_value is not None and name in NUMERIC_HINT_FIELDS:
            entry['value'] = numeric_value
            return entry

        entry['value'] = text
        return entry

    def _get_row_value(self, row: Dict, col_map: Dict, source: str) -> str:
        """从AI字段、标准字段或列映射中提取值"""
        direct_value = row.get(source)
        cleaned = self._clean_value(direct_value)
        if cleaned:
            return cleaned

        col_name = col_map.get(source)
        if col_name:
            mapped_value = row.get(col_name)
            cleaned = self._clean_value(mapped_value)
            if cleaned:
                return cleaned

        return ""

    def _collect_dynamic_fields(self, row: Dict, product: Dict):
        """保留 Excel 中已有、但当前显式映射尚未覆盖的 snake_case 类目字段。"""
        for raw_name, raw_value in row.items():
            if str(raw_name).startswith('_'):
                continue

            field_name = self._normalize_dynamic_field_name(raw_name)
            if not field_name or field_name in product:
                continue
            if field_name in GENERIC_ATTRIBUTE_EXCLUDE:
                continue

            clean_value = self._clean_value(raw_value)
            if clean_value:
                product[field_name] = clean_value

    def _normalize_dynamic_field_name(self, raw_name) -> str:
        text = self._clean_value(raw_name)
        if text.startswith(('source_', 'ai_', 'listing_check_', 'validation_', 'preview_', 'submit_', 'template_')):
            return ""
        alias_map = {
            'are_batteries_included': 'batteries_included',
            'are_batteries_required': 'batteries_required',
        }
        normalized = alias_map.get(text, text)
        if not re.match(r'^[a-z][a-z0-9_]*$', normalized):
            return ""
        return normalized

    def _load_schema_fields(self, product_type: str) -> Optional[Dict]:
        product_type = self._clean_value(product_type)
        if not product_type:
            return None

        marketplace_code = MARKETPLACE_CODE_BY_ID.get(self.marketplace_id, 'US')
        cache_key = (product_type, marketplace_code)
        if cache_key in self._schema_fields_cache:
            return self._schema_fields_cache[cache_key]

        try:
            from amazon.schema_manager import fetch_schema, parse_schema
            schema_fields = parse_schema(fetch_schema(product_type, marketplace_code))
        except Exception as exc:
            logger.warning(f"  ⚠️ 加载 Schema 失败，已回退到基础校验: {product_type} ({exc})")
            schema_fields = None

        self._schema_fields_cache[cache_key] = schema_fields
        return schema_fields

    def _resolve_product_identity_mode(self, product: Dict) -> str:
        explicit = self._clean_value(product.get('product_identity_mode')).lower()
        aliases = {
            'real_gtin': 'real_gtin',
            'gtin': 'real_gtin',
            'real': 'real_gtin',
            'barcode': 'real_gtin',
            'real_barcode': 'real_gtin',
            '真实gtin': 'real_gtin',
            '真实条码': 'real_gtin',
            '真实编码': 'real_gtin',
            'gtin_exemption': 'gtin_exemption',
            'exemption': 'gtin_exemption',
            'exempt': 'gtin_exemption',
            '免码': 'gtin_exemption',
            '免gtin': 'gtin_exemption',
            '内部码': 'internal_code',
            'internal_code': 'internal_code',
            'internal': 'internal_code',
            'erp': 'internal_code',
            'erp_code': 'internal_code',
        }
        if explicit:
            return aliases.get(explicit, explicit)

        if self._clean_value(product.get('upc') or product.get('ean') or product.get('gtin')):
            return 'real_gtin'
        return ''

    def _validate_product_identity(self, product: Dict) -> Dict:
        mode = self._resolve_product_identity_mode(product)
        code = self._clean_value(product.get('upc') or product.get('ean') or product.get('gtin'))
        explicit_type = self._clean_value(product.get('external_product_id_type')).upper()
        identifier_type = explicit_type or ('UPC' if product.get('upc') else 'EAN' if product.get('ean') else 'GTIN' if product.get('gtin') else '')

        result = {
            'mode': mode,
            'mode_missing': '',
            'errors': [],
            'warnings': [],
            'info': [],
        }

        if not mode:
            result['mode_missing'] = "未声明商品标识模式，当前会按“真实GTIN”或“待确认”推断，建议明确选择：真实GTIN / GTIN免码 / 内部码"
            if not code:
                result['errors'].append("缺少商品标识信息：请填写真实 UPC/EAN/GTIN，或将商品标识模式改为 GTIN免码")
            return result

        if mode == 'internal_code':
            result['errors'].append("当前商品标识模式为“内部码”，不能直接用于亚马逊正式新建；请改为真实GTIN或GTIN免码")
            if code:
                result['warnings'].append("已填写的 UPC/EAN/GTIN 将被视为内部码，不应直接提交给亚马逊")
            return result

        if mode == 'gtin_exemption':
            if code:
                result['warnings'].append("当前商品标识模式为 GTIN免码，提交时会忽略已填写的 UPC/EAN/GTIN")
            result['info'].append("GTIN免码模式下，请确保对应店铺/品牌/类目已获得 Amazon GTIN exemption")
            return result

        if mode == 'real_gtin':
            if not code:
                result['errors'].append("商品标识模式为“真实GTIN”时，必须填写真实 UPC/EAN/GTIN")
                return result

            format_error = self._validate_identifier_format(code, identifier_type or 'GTIN')
            if format_error:
                result['errors'].append(format_error)
            elif not explicit_type:
                result['warnings'].append("未填写产品ID类型，当前按检测到的条码字段自动推断")
            return result

        result['warnings'].append(f"未知的商品标识模式: {mode}")
        return result

    def _validate_identifier_format(self, code: str, identifier_type: str) -> str:
        compact = re.sub(r'[\s-]+', '', self._clean_value(code))
        if not compact:
            return "条码为空"
        if not compact.isdigit():
            return f"{identifier_type} 必须为纯数字，当前值: {code}"

        rules = {
            'UPC': {12},
            'EAN': {13},
            'GTIN': {8, 12, 13, 14},
            'ISBN': {10, 13},
        }
        expected_lengths = rules.get(identifier_type.upper(), {8, 12, 13, 14})
        if len(compact) not in expected_lengths:
            allowed = '/'.join(str(length) for length in sorted(expected_lengths))
            return f"{identifier_type} 长度不合法，应为 {allowed} 位数字，当前为 {len(compact)} 位"
        return ""

    def _resolve_main_image(self, row: Dict, col_map: Dict):
        """优先使用AI公开图；本地路径仅在配置公开基地址后转换为可提交URL。"""
        original_url = self._get_row_value(row, col_map, 'image_url')
        ai_public_url = self._clean_value(row.get('AI主图URL'))
        ai_main_image_path = self._clean_value(row.get('AI主图路径'))

        if self._is_media_locator(ai_public_url):
            return ai_public_url, 'ai_public', ai_main_image_path

        if ai_main_image_path:
            public_url = self._local_image_to_public_url(ai_main_image_path)
            if public_url:
                return public_url, 'ai_public', ai_main_image_path
            if original_url:
                return original_url, 'original_fallback_local_ai', ai_main_image_path
            return ai_main_image_path, 'local_only_ai', ai_main_image_path

        if original_url:
            source = 'original' if self._is_http_url(original_url) else 'invalid_original'
            return original_url, source, ""

        return "", "missing", ai_main_image_path

    def _resolve_additional_image(self, row: Dict, col_map: Dict, slot: int) -> str:
        """优先使用 AI 副图公开地址，回退到原始副图列。"""
        ai_public_url = self._clean_value(row.get(f'AI副图{slot}URL'))
        if self._is_media_locator(ai_public_url):
            return ai_public_url

        ai_local_path = self._clean_value(row.get(f'AI副图{slot}路径'))
        if ai_local_path:
            public_url = self._local_image_to_public_url(ai_local_path)
            if public_url:
                return public_url

        for source in (f'image_{slot}', f'other_image_url_{slot-1}'):
            original_url = self._get_row_value(row, col_map, source)
            if original_url:
                return original_url

        return ""

    def _local_image_to_public_url(self, local_path: str) -> str:
        """将本地输出图路径转换为公开URL。"""
        public_base = getattr(self.config, 'OUTPUT_IMAGE_PUBLIC_BASE', '') or ''
        if not public_base:
            return ""

        filename = os.path.basename(str(local_path).replace('\\', os.sep))
        if not filename:
            return ""

        quoted_filename = quote(filename)
        if '{filename}' in public_base:
            return public_base.format(filename=quoted_filename)
        return f"{public_base.rstrip('/')}/{quoted_filename}"

    def _clean_value(self, value) -> str:
        if value is None:
            return ""
        text = str(value).strip()
        return text

    def _humanize_enum_label(self, value) -> str:
        text = self._clean_value(value)
        if not text:
            return text
        return re.sub(r'[\s_\-]+', ' ', text).title()

    def _normalize_bottoms_size_value(self, value: str) -> str:
        text = self._clean_value(value).lower()
        mapping = {
            's': 's',
            'small': 's',
            'm': 'm',
            'medium': 'm',
            'l': 'l',
            'large': 'l',
            'xl': 'x_l',
            'x-large': 'x_l',
            'x large': 'x_l',
            'extra large': 'x_l',
            'xxl': 'xx_l',
            '2xl': '2x_l',
            'xxxl': '3x_l',
        }
        return mapping.get(text, value)

    def _coerce_bool(self, value) -> bool:
        return str(value).strip().lower() in ('true', 'yes', '1', '是')

    def _coerce_number(self, value):
        text = self._clean_value(value)
        if not re.match(r'^-?\d+(?:\.\d+)?$', text):
            return None
        try:
            number = float(text)
        except (ValueError, TypeError):
            return None
        return int(number) if number.is_integer() else number

    def _is_http_url(self, value) -> bool:
        text = self._clean_value(value).lower()
        return text.startswith('http://') or text.startswith('https://')

    def _is_media_locator(self, value) -> bool:
        text = self._clean_value(value).lower()
        return text.startswith('s3://') or self._is_http_url(text)

    def _normalize_condition_type(self, value) -> str:
        text = self._clean_value(value).lower()
        mapping = {
            'new': 'new_new',
            'new_new': 'new_new',
            'used_like_new': 'used_like_new',
            'used_very_good': 'used_very_good',
            'used_good': 'used_good',
            'used_acceptable': 'used_acceptable',
        }
        return mapping.get(text, text or 'new_new')

    def _parse_measurement(self, raw_value, fallback_unit: str = 'count') -> Optional[Dict]:
        text = self._clean_value(raw_value)
        if not text:
            return None

        match = re.search(r'(-?\d+(?:\.\d+)?)\s*([A-Za-z_ ]+)?', text)
        if not match:
            return None

        try:
            value = float(match.group(1))
        except (ValueError, TypeError):
            return None

        unit = (match.group(2) or '').strip().lower().replace(' ', '_') or fallback_unit
        aliases = {
            'fluid_ounce': 'fluid_ounces',
            'fluid_ounces': 'fluid_ounces',
            'ounce': 'ounces',
            'ounces': 'ounces',
            'oz': 'ounces',
            'count': 'count',
            'ct': 'count',
        }
        unit = aliases.get(unit, unit)
        return {'value': value, 'unit': unit}

    def _build_dimensions(self, product: Dict) -> Optional[Dict]:
        """构建商品尺寸"""
        length = product.get('item_length')
        width = product.get('item_width')
        height = product.get('item_height')

        if not any([length, width, height]):
            return None

        dims = {'marketplace_id': self.marketplace_id}
        unit = product.get('dimension_unit', 'centimeters')
        if length:
            entry = self._build_measurement_value(length, unit, 'item_length')
            if entry:
                dims['length'] = entry
        if width:
            entry = self._build_measurement_value(width, unit, 'item_width')
            if entry:
                dims['width'] = entry
        if height:
            entry = self._build_measurement_value(height, unit, 'item_height')
            if entry:
                dims['height'] = entry
        return dims if len(dims) > 1 else None

    def _build_item_width_height(self, product: Dict) -> Optional[Dict]:
        """构建 Item Dimensions W x H。"""
        width = product.get('item_width')
        height = product.get('item_height')

        if not width or not height:
            return None

        unit = product.get('dimension_unit', 'inches')
        width_entry = self._build_measurement_value(width, unit, 'item_width')
        height_entry = self._build_measurement_value(height, unit, 'item_height')
        if not width_entry or not height_entry:
            return None

        return {
            'marketplace_id': self.marketplace_id,
            'width': width_entry,
            'height': height_entry,
        }

    def _build_item_depth_width_height(self, product: Dict) -> Optional[Dict]:
        """构建 Item Dimensions D x W x H。Amazon 的 depth 对应这里的 length。"""
        depth = product.get('item_length')
        width = product.get('item_width')
        height = product.get('item_height')

        if not all([depth, width, height]):
            return None

        unit = product.get('dimension_unit', 'inches')
        depth_entry = self._build_measurement_value(depth, unit, 'item_length')
        width_entry = self._build_measurement_value(width, unit, 'item_width')
        height_entry = self._build_measurement_value(height, unit, 'item_height')
        if not depth_entry or not width_entry or not height_entry:
            return None

        return {
            'marketplace_id': self.marketplace_id,
            'depth': depth_entry,
            'width': width_entry,
            'height': height_entry,
        }

    def _build_package_dimensions(self, product: Dict) -> Optional[Dict]:
        """构建包装尺寸"""
        pl = product.get('package_length') or product.get('item_length')
        pw = product.get('package_width') or product.get('item_width')
        ph = product.get('package_height') or product.get('item_height')

        if not any([pl, pw, ph]):
            return None

        dims = {'marketplace_id': self.marketplace_id}
        unit = product.get('dimension_unit', 'centimeters')
        if pl:
            entry = self._build_measurement_value(pl, unit, 'package_length')
            if entry:
                dims['length'] = entry
        if pw:
            entry = self._build_measurement_value(pw, unit, 'package_width')
            if entry:
                dims['width'] = entry
        if ph:
            entry = self._build_measurement_value(ph, unit, 'package_height')
            if entry:
                dims['height'] = entry
        return dims if len(dims) > 1 else None

    def _build_measurement_value(self, raw_value, unit: str, field_name: str) -> Optional[Dict]:
        """将尺寸文本安全转换为 SP-API measurement 结构。"""
        try:
            return {'unit': unit, 'value': float(raw_value)}
        except (ValueError, TypeError):
            logger.warning(f"  ⚠️ 无效尺寸字段 {field_name}: {raw_value}")
            return None
