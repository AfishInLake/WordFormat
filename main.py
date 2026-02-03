#! /usr/bin/env python
# @Time    : 2026/1/18 13:22
# @Author  : afish
# @File    : main.py
import argparse
import os
from pathlib import Path

from loguru import logger

# 导入核心函数
from src.set_style import auto_format_thesis_document
from src.set_tag import main as set_tag_main


def validate_file(path: str, file_type: str = "文件") -> str:
    """校验文件是否存在，且为文件类型（非文件夹）"""
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise argparse.ArgumentTypeError(f"{file_type}不存在: {abs_path}")
    if not os.path.isfile(abs_path):
        raise argparse.ArgumentTypeError(f"{file_type}路径非文件: {abs_path}")
    return abs_path


def get_json_path(docx_path: str, json_dir: str = "tmp/") -> str:
    """保留原函数（generate-json模式生成JSON时使用）"""
    docx_path = Path(docx_path)
    json_save_path = Path(os.path.join(json_dir, f"{docx_path.stem}.json"))
    json_save_path.parent.mkdir(parents=True, exist_ok=True)
    return str(json_save_path)


def create_common_parser(
    subparser, name: str, description: str
) -> argparse.ArgumentParser:
    """抽离公共参数【移除子命令--json参数，全局已指定】"""
    parser = subparser.add_parser(name=name, description=description, help=description)
    parser.add_argument(
        "--config",
        "-c",
        required=True,
        type=lambda x: validate_file(x, "配置文件"),
        help="格式配置YAML路径（必填），例如：example/undergrad_thesis.yaml",
    )
    parser.add_argument(
        "--output",
        "-o",
        default="output/",
        help="校验/格式化后文档保存目录（可选，默认：output/）",
    )
    return parser


if __name__ == "__main__":
    # 1. 创建参数解析器
    parser = argparse.ArgumentParser(
        description="学位论文格式自动校验工具（多模式控制）"
    )

    # 2. 全局参数【核心改造：移除--json-dir，新增全局--json/-jf（指定完整JSON路径）】
    parser.add_argument(
        "--docx",
        "-d",
        required=True,
        type=lambda x: validate_file(x, "Word文档"),
        help="待处理的Word文档路径（必填），例如：tmp/毕业设计说明书.docx",
    )
    parser.add_argument(
        "--json",
        "-jf",  # 全局指定JSON完整路径，短选项保持jf（符合使用习惯）
        required=True,
        type=lambda x: validate_file(x, "JSON文件"),
        help="JSON文件完整路径（必填），例如：output/毕业设计说明书.json",
    )
    # 保留--json-dir（仅generate-json模式使用，用于指定JSON生成目录，非必填）
    parser.add_argument(
        "--json-dir",
        "-j",
        default="tmp/",
        help="【仅generate-json模式有效】JSON文件生成目录（可选，默认：tmp/）",
    )

    # 3. 子命令解析器
    subparsers = parser.add_subparsers(
        dest="mode",
        required=True,
    )

    # 3.1 模式1：仅生成JSON
    parser_gen = subparsers.add_parser(
        "generate-json", help="仅生成文档结构JSON文件，不执行校验/格式化"
    )
    parser_gen.add_argument(
        "--config",
        "-c",
        required=True,
        type=lambda x: validate_file(x, "配置文件"),
        help="格式配置YAML路径（必填），例如：example/undergrad_thesis.yaml",
    )

    # 3.2 模式2：仅执行格式校验
    create_common_parser(
        subparsers,
        name="check-format",
        description="仅执行格式校验（需先生成JSON文件）",
    )

    # 3.3 模式3：仅执行格式应用
    create_common_parser(
        subparsers,
        name="apply-format",
        description="仅执行格式应用/格式化（需先生成JSON文件）",
    )

    # 4. 解析参数
    args = parser.parse_args()
    docx_abs_path = os.path.abspath(args.docx)
    json_abs_path = os.path.abspath(args.json)  # 全局JSON完整路径

    # 自动创建输出目录（若当前模式有output参数）
    if hasattr(args, "output"):
        Path(args.output).mkdir(parents=True, exist_ok=True)

    # 5. 模式执行逻辑【改造：移除JSON路径推导，直接使用全局--json传入的完整路径】
    if args.mode == "generate-json":
        # 模式1：仅生成JSON（使用--json-dir指定的目录生成，保留校验，仅使用其文件名）
        logger.info("=" * 60)
        logger.info("📌 执行模式：仅生成JSON文件")
        logger.info(f"📄 源Word文档：{docx_abs_path}")  # noqa E501
        # 生成JSON路径（使用--json-dir目录 + docx同名）
        gen_json_path = get_json_path(args.docx, args.json_dir)
        logger.info(f"📋 生成的JSON路径：{gen_json_path}")
        logger.info("=" * 60)

        set_tag_main(
            docx_path=args.docx,
            json_save_path=gen_json_path,
            configpath=args.config,
        )
        logger.info("\n✅ JSON文件已生成完成！")
        logger.info(f"📝 JSON路径：{os.path.abspath(gen_json_path)}")
        logger.info("💡 可使用该JSON文件配合 check-format/apply-format 模式执行操作")

    elif args.mode == "check-format":
        # 模式2：仅校验（直接使用全局--json传入的完整路径）
        logger.info("=" * 60)
        logger.info("📌 执行模式：仅执行格式校验")
        logger.info(f"📄 源Word文档：{docx_abs_path}")
        logger.info(f"📋 JSON文件：{json_abs_path}")
        logger.info(f"⚙️  配置文件：{args.config}")
        logger.info(f"💾 输出目录：{args.output}")
        logger.info("=" * 60)

        auto_format_thesis_document(
            jsonpath=json_abs_path,
            docxpath=args.docx,
            configpath=args.config,
            savepath=args.output,
            check=True,
        )
        logger.info(f"\n✅ 格式校验完成！校验后文档已保存至：{args.output}")

    elif args.mode == "apply-format":
        # 模式3：格式化（直接使用全局--json传入的完整路径）
        logger.info("=" * 60)
        logger.info("📌 执行模式：仅执行格式应用/格式化")
        logger.info(f"📄 源Word文档：{docx_abs_path}")
        logger.info(f"📋 JSON文件：{json_abs_path}")
        logger.info(f"⚙️  配置文件：{args.config}")
        logger.info(f"💾 输出目录：{args.output}")
        logger.info("=" * 60)

        auto_format_thesis_document(
            jsonpath=json_abs_path,
            docxpath=args.docx,
            configpath=args.config,
            savepath=args.output,
            check=False,
        )
        logger.info(f"\n✅ 格式化完成！格式化后文档已保存至：{args.output}")
