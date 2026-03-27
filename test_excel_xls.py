from pathlib import Path
from uuid import uuid4

import xlwt

from core.excel.processor import ExcelProcessor


def _create_xls(path: Path):
    workbook = xlwt.Workbook()
    sheet = workbook.add_sheet('Sheet1')
    headers = ['SKU', 'item_name', 'main_image_url']
    values = ['SKU-1', 'Demo Product', 'https://example.com/demo.jpg']

    for col_idx, header in enumerate(headers):
        sheet.write(0, col_idx, header)
    for col_idx, value in enumerate(values):
        sheet.write(1, col_idx, value)

    workbook.save(str(path))


def test_excel_processor_reads_xls():
    output_dir = Path('output').resolve()
    path = output_dir / f'test_xls_{uuid4().hex[:8]}.xls'

    try:
        _create_xls(path)

        processor = ExcelProcessor()
        data = processor.read_input(str(path))
        col_map = processor.detect_columns()

        assert len(data) == 1
        assert data[0]['SKU'] == 'SKU-1'
        assert col_map['sku'] == 'SKU'
        assert col_map['title'] == 'item_name'
        assert col_map['image_url'] == 'main_image_url'
    finally:
        if path.exists():
            path.unlink()
