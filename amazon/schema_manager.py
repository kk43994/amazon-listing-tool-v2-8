"""
Schema 获取与解析
调用 ProductTypeDefinitions API -> 下载 JSON Schema -> 解析为字段清单 -> 本地缓存
"""
import os
import json
import time
import logging
import requests
from typing import Dict, List, Optional
from core.runtime_paths import runtime_path

logger = logging.getLogger(__name__)

CACHE_DIR = runtime_path('output', 'schema_cache')
CACHE_TTL = 7 * 24 * 3600  # 7 days


def _cache_path(product_type: str, marketplace: str) -> str:
    os.makedirs(CACHE_DIR, exist_ok=True)
    return os.path.join(CACHE_DIR, f'{product_type}_{marketplace}.json')


def search_product_types(keyword: str, sp_client=None) -> List[Dict]:
    """根据商品名搜索推荐的产品类型"""
    if sp_client is None:
        from amazon.sp_client import SPClient
        sp_client = SPClient()
    return sp_client.search_product_types(keyword)


def fetch_schema(product_type: str, marketplace: str = 'US', sp_client=None) -> Dict:
    """获取并缓存 JSON Schema"""
    # Check cache
    cp = _cache_path(product_type, marketplace)
    if os.path.exists(cp):
        mtime = os.path.getmtime(cp)
        if time.time() - mtime < CACHE_TTL:
            with open(cp, 'r', encoding='utf-8') as f:
                logger.info(f"Using cached schema: {product_type}")
                return json.load(f)

    # Fetch from API
    if sp_client is None:
        from amazon.sp_client import SPClient
        sp_client = SPClient()

    definition = sp_client.get_schema(product_type)

    # Download the actual JSON Schema from the schema link
    schema_link = definition.get('schema', {}).get('link', {}).get('resource')
    schema_json = None
    if schema_link:
        try:
            resp = requests.get(schema_link, timeout=30)
            resp.raise_for_status()
            schema_json = resp.json()
        except Exception as e:
            logger.warning(f"Failed to download schema JSON: {e}")

    result = {
        'product_type': product_type,
        'marketplace': marketplace,
        'property_groups': definition.get('propertyGroups', {}),
        'requirements': definition.get('requirements', []),
        'schema': schema_json,
        'fetched_at': time.time(),
    }

    # Cache
    with open(cp, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    logger.info(f"Schema cached: {cp}")

    return result


def parse_schema(schema_data: Dict) -> Dict:
    """
    解析 schema 为字段清单

    Returns:
        {
            'required_fields': [{'name', 'type', 'title', 'description', 'group'}],
            'optional_fields': [...],
            'enum_fields': {'field_name': ['val1', 'val2', ...]},
            'field_groups': {'group_name': {'title', 'description', 'fields': [...]}}
        }
    """
    schema_json = schema_data.get('schema') or {}
    property_groups = schema_data.get('property_groups') or {}
    properties = schema_json.get('properties', {})
    required_names = set(schema_json.get('required', []))

    # Build group -> fields mapping from propertyGroups
    field_to_group = {}
    for group_name, group_info in property_groups.items():
        for prop_name in group_info.get('propertyNames', []):
            field_to_group[prop_name] = group_name

    required_fields = []
    optional_fields = []
    enum_fields = {}
    field_groups = {}

    for name, prop in properties.items():
        # Resolve field info from array items or direct
        field_def = prop
        if prop.get('type') == 'array' and 'items' in prop:
            field_def = prop['items']
            # Could be nested further
            if '$ref' in field_def:
                continue  # Skip complex refs for now

        field_info = {
            'name': name,
            'type': _extract_type(field_def),
            'title': prop.get('title', name),
            'description': prop.get('description', ''),
            'group': field_to_group.get(name, 'other'),
        }

        # Check for enum values (could be in items.properties.value.enum)
        enum_vals = _extract_enum(field_def)
        if enum_vals:
            enum_fields[name] = enum_vals
            field_info['enum'] = True

        if name in required_names:
            required_fields.append(field_info)
        else:
            optional_fields.append(field_info)

        # Group
        grp = field_info['group']
        if grp not in field_groups:
            gi = property_groups.get(grp, {})
            field_groups[grp] = {
                'title': gi.get('title', grp),
                'description': gi.get('description', ''),
                'fields': [],
            }
        field_groups[grp]['fields'].append(field_info)

    return {
        'required_fields': required_fields,
        'optional_fields': optional_fields,
        'enum_fields': enum_fields,
        'field_groups': field_groups,
    }


def _extract_type(field_def: Dict) -> str:
    """Extract field type from schema definition"""
    if 'type' in field_def:
        return field_def['type']
    if 'properties' in field_def:
        # Look for a 'value' property
        val_prop = field_def.get('properties', {}).get('value', {})
        return val_prop.get('type', 'string')
    return 'string'


def _extract_enum(field_def: Dict) -> Optional[List[str]]:
    """Extract enum values from schema definition"""
    # Direct enum
    if 'enum' in field_def:
        return field_def['enum']
    # Nested in properties.value.enum
    props = field_def.get('properties', {})
    val_prop = props.get('value', {})
    if 'enum' in val_prop:
        return val_prop['enum']
    # In items
    items = field_def.get('items', {})
    if 'enum' in items:
        return items['enum']
    val_in_items = items.get('properties', {}).get('value', {})
    if 'enum' in val_in_items:
        return val_in_items['enum']
    return None
