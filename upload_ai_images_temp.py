"""
临时匿名上传工具。

将 Excel 中 AI主图路径 指向的本地图片上传到 x0.at，
并把返回的 URL 写回 AI主图URL 列。
"""
import argparse
import os
import time

from config import get_config
from core.temp_image_host import populate_ai_image_urls


def find_latest_output_xlsx(output_dir: str) -> str:
    os.makedirs(output_dir, exist_ok=True)
    xlsx_files = [
        os.path.join(output_dir, name)
        for name in os.listdir(output_dir)
        if name.endswith((".xlsx", ".xls")) and not name.startswith(("~", "."))
    ]
    if not xlsx_files:
        return ""
    xlsx_files.sort(key=os.path.getmtime, reverse=True)
    return xlsx_files[0]


def main():
    parser = argparse.ArgumentParser(description="临时上传 AI 图片到 x0.at 并写回 Excel")
    parser.add_argument("--input", "-i", help="输入Excel路径，默认自动取 output/ 下最新文件")
    parser.add_argument("--output", "-o", help="输出Excel路径，默认生成一个新文件")
    parser.add_argument("--in-place", action="store_true", help="原地更新输入文件")
    parser.add_argument("--overwrite", action="store_true", help="即使 AI主图URL 已存在也重新上传覆盖")
    args = parser.parse_args()

    config = get_config()
    input_path = args.input or find_latest_output_xlsx(config.OUTPUT_DIR)
    if not input_path:
        raise SystemExit("未找到可处理的 Excel 文件，请使用 --input 指定路径")
    if not os.path.exists(input_path):
        raise SystemExit(f"输入文件不存在: {input_path}")

    if args.in_place:
        output_path = input_path
    elif args.output:
        output_path = args.output
    else:
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(config.OUTPUT_DIR, f"临时图床结果_{timestamp}.xlsx")

    stats = populate_ai_image_urls(
        input_path,
        output_path=output_path,
        only_missing=not args.overwrite,
    )

    print("临时图床处理完成")
    print(f"输入文件: {input_path}")
    print(f"输出文件: {output_path}")
    print(
        "统计: "
        f"uploaded={stats['uploaded']} "
        f"copied={stats['copied']} "
        f"skipped={stats['skipped']} "
        f"missing={stats['missing']} "
        f"failed={stats['failed']} "
        f"total={stats['total']}"
    )


if __name__ == "__main__":
    main()
