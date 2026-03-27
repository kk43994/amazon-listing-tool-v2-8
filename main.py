"""
亚马逊商品处理工具 - 主入口
用法:
    python main.py --stage 1 --input input/商品数据.xlsx
    python main.py --stage 1 --preview
    python main.py --stage 2 --input output/处理结果_xxx.xlsx
    python main.py --stage 2 --preview --rows 1-5
"""
import os
import sys
import argparse
import logging
from datetime import datetime

from config import get_config

# 配置日志
def setup_logging(verbose: bool = False):
    config = get_config()
    os.makedirs(config.LOGS_DIR, exist_ok=True)
    log_file = os.path.join(config.LOGS_DIR, f'{datetime.now():%Y%m%d_%H%M%S}.log')

    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_file, encoding='utf-8'),
        ]
    )
    return logging.getLogger(__name__)


def find_input_file(config, stage: int) -> str:
    """自动查找输入文件"""
    if stage == 1:
        search_dir = config.INPUT_DIR
    else:
        search_dir = config.OUTPUT_DIR

    os.makedirs(search_dir, exist_ok=True)
    xlsx_files = [f for f in os.listdir(search_dir)
                  if f.endswith(('.xlsx', '.xls')) and not f.startswith('~')]

    if not xlsx_files:
        return None

    # 按修改时间排序，取最新
    xlsx_files.sort(
        key=lambda f: os.path.getmtime(os.path.join(search_dir, f)),
        reverse=True
    )
    return os.path.join(search_dir, xlsx_files[0])


def main():
    parser = argparse.ArgumentParser(
        description='亚马逊商品处理与自动提交工具',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  %(prog)s --stage 1 --input input/商品数据.xlsx        第一阶段：AI内容处理
  %(prog)s --stage 1 --preview                          预览模式
  %(prog)s --stage 1 --no-images                        只生成文案，跳过图片
  %(prog)s --stage 1 --rows 1-5                         只处理前5条
  %(prog)s --stage 1 --resume                           断点续传
  %(prog)s --stage 2 --input output/结果.xlsx            第二阶段：SP-API提交
  %(prog)s --stage 2 --preview                          预览提交计划
  %(prog)s --stage 2 --mode batch                       批量Feed提交
        """
    )

    parser.add_argument('--stage', type=int, choices=[1, 2], default=1,
                       help='执行阶段: 1=AI内容处理, 2=SP-API提交')
    parser.add_argument('--input', '-i', type=str,
                       help='输入Excel文件路径 (默认自动检测)')
    parser.add_argument('--output', '-o', type=str, default=None,
                       help='输出文件路径')
    parser.add_argument('--rows', type=str, default=None,
                       help='处理行范围 (如 1-10 或 3)')
    parser.add_argument('--no-images', action='store_true',
                       help='[阶段1] 跳过图片处理')
    parser.add_argument('--no-text', action='store_true',
                       help='[阶段1] 跳过文案生成')
    parser.add_argument('--preview', action='store_true',
                       help='预览模式（不实际执行API调用）')
    parser.add_argument('--resume', action='store_true',
                       help='[阶段1] 从上次断点续传')
    parser.add_argument('--mode', type=str, choices=['individual', 'batch'],
                       default='individual',
                       help='[阶段2] 提交模式: individual=逐条, batch=批量Feed')
    parser.add_argument('--verbose', '-v', action='store_true',
                       help='详细日志')

    args = parser.parse_args()
    logger = setup_logging(args.verbose)
    config = get_config()

    logger.info("=" * 60)
    logger.info(f"🚀 亚马逊商品处理工具 v1.0")
    logger.info(f"   阶段: {args.stage}")
    logger.info(f"   模式: {'预览' if args.preview else '执行'}")
    logger.info("=" * 60)

    # 查找输入文件
    if not args.input:
        args.input = find_input_file(config, args.stage)
        if args.input:
            logger.info(f"📂 自动检测到: {args.input}")
        else:
            target_dir = config.INPUT_DIR if args.stage == 1 else config.OUTPUT_DIR
            logger.error(f"❌ 未找到Excel文件")
            logger.error(f"   请将文件放入 {target_dir}/ 目录")
            logger.error(f"   或使用 --input 指定路径")
            sys.exit(1)

    if not os.path.exists(args.input):
        logger.error(f"❌ 文件不存在: {args.input}")
        sys.exit(1)

    # ========== 第一阶段 ==========
    if args.stage == 1:
        # 检查API Key (预览模式不需要)
        if not args.preview:
            if not config.AI_API_KEY or config.AI_API_KEY == 'your-api-key-here':
                logger.error("❌ 请在 .env 文件中配置 AI_API_KEY")
                sys.exit(1)

        from stage1_pipeline import Stage1Pipeline
        pipeline = Stage1Pipeline()

        if args.preview:
            # 预览模式：只读取Excel，显示会处理哪些数据
            data = pipeline.excel.read_input(args.input)
            col_map = pipeline.excel.detect_columns()

            if args.rows:
                data = _filter_rows(data, args.rows)

            logger.info(f"\n📋 预览: 将处理 {len(data)} 条商品")
            logger.info(f"   图片处理: {'✅' if not args.no_images else '❌ 跳过'}")
            logger.info(f"   文案生成: {'✅' if not args.no_text else '❌ 跳过'}")
            logger.info(f"\n   列映射:")
            for field, col in col_map.items():
                if col:
                    logger.info(f"     {field} → {col}")

            for idx, item in enumerate(data[:5]):  # 只显示前5条
                title = item.get(col_map.get('title', ''), 'N/A')
                sku = item.get(col_map.get('sku', ''), 'N/A')
                logger.info(f"\n   #{idx+1} SKU={sku} 标题={str(title)[:50]}...")

            if len(data) > 5:
                logger.info(f"\n   ... 还有 {len(data)-5} 条")
        else:
            result = pipeline.run(
                input_file=args.input,
                output_file=args.output,
                process_images=not args.no_images,
                process_text=not args.no_text,
                rows=args.rows,
                resume=args.resume,
            )
            logger.info(f"\n✅ 第一阶段完成! 输出: {result}")

    # ========== 第二阶段 ==========
    elif args.stage == 2:
        from stage2_pipeline import Stage2Pipeline
        pipeline = Stage2Pipeline()
        result = pipeline.run(
            input_file=args.input,
            output_file=args.output,
            mode=args.mode,
            preview=args.preview,
            rows=args.rows,
        )
        if not args.preview:
            logger.info(f"\n✅ 第二阶段完成! 报告: {result}")


def _filter_rows(data, rows: str):
    """根据行范围过滤"""
    try:
        if '-' in rows:
            start, end = rows.split('-')
            return data[int(start)-1:int(end)]
        else:
            return [data[int(rows)-1]]
    except (ValueError, IndexError):
        return data


if __name__ == '__main__':
    main()
