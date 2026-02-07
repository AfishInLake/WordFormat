#! /usr/bin/env python
# @Time    : 2026/1/18 13:22
# @Author  : afish
# @File    : cli.py
import argparse
import os
from pathlib import Path

from loguru import logger

# 导入核心函数
from wordformat.set_style import auto_format_thesis_document
from wordformat.set_tag import set_tag_main


def validate_file(path: str, file_type: str = "文件") -> str:
    """校验文件是否存在，且为文件类型（非文件夹）"""
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise argparse.ArgumentTypeError(f"{file_type}不存在: {abs_path}")
    if not os.path.isfile(abs_path):
        raise argparse.ArgumentTypeError(f"{file_type}路径非文件: {abs_path}")
    return abs_path


def create_common_parser(
    subparser, name: str, description: str
) -> argparse.ArgumentParser:
    """抽离公共参数：--config（必填）、--output（可选）"""
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


# 新增：针对不同模式的JSON路径校验器（核心修复）
def validate_json_path(path: str, mode: str) -> str:
    """
    按执行模式校验JSON路径：
    - generate-json：仅校验路径合法性，不要求文件存在（自动创建）
    - check/apply-format：强制要求文件存在（需提前生成）
    """
    abs_path = os.path.abspath(path)
    # 统一校验路径是否为合法文件路径（排除文件夹）
    if os.path.exists(abs_path) and os.path.isdir(abs_path):
        raise argparse.ArgumentTypeError(f"JSON路径不能是文件夹: {abs_path}")
    # 仅非生成模式，要求文件存在
    if mode != "generate-json" and not os.path.exists(abs_path):
        raise argparse.ArgumentTypeError(f"JSON文件不存在: {abs_path}")
    return abs_path


def main():
    # 1. 创建参数解析器
    parser = argparse.ArgumentParser(
        description="学位论文格式自动校验工具（多模式控制）"
    )

    # 2. 全局参数（核心简化：仅保留--docx、--json，移除冗余--json-dir）
    parser.add_argument(
        "--docx",
        "-d",
        required=True,
        type=lambda x: validate_file(x, "Word文档"),
        help="待处理的Word文档路径（必填），例如：tmp/毕业设计说明书.docx",
    )
    parser.add_argument(
        "--json",
        "-jf",
        required=True,
        help="JSON文件完整路径（必填）：generate-json模式下为生成路径，check/apply模式下为读取路径",
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

    # 4. 解析参数 + 按模式校验JSON路径（核心修复步骤）
    args = parser.parse_args()
    docx_abs_path = os.path.abspath(args.docx)
    # 关键：传入当前模式，动态校验JSON路径
    json_abs_path = validate_json_path(args.json, args.mode)
    # 提取JSON路径的目录，自动创建（生成模式必备）
    json_dir = os.path.dirname(json_abs_path)
    Path(json_dir).mkdir(parents=True, exist_ok=True)

    # 自动创建输出目录（若当前模式有output参数）
    if hasattr(args, "output"):
        Path(args.output).mkdir(parents=True, exist_ok=True)

    # 5. 模式执行逻辑（彻底简化：所有模式统一使用json_abs_path，无任何路径推导）
    if args.mode == "generate-json":
        # 模式1：生成JSON → 直接使用用户指定的json_abs_path生成（自动创建目录/文件）
        logger.info("=" * 60)
        logger.info("📌 执行模式：仅生成JSON文件")
        logger.info(f"📄 源Word文档：{docx_abs_path}")
        logger.info(f"📋 生成的JSON路径：{json_abs_path}")  # 直接使用用户指定路径
        logger.info("=" * 60)

        set_tag_main(
            docx_path=args.docx,
            json_save_path=json_abs_path,  # 核心：传入用户指定的路径
            configpath=args.config,
        )
        logger.info("\n✅ JSON文件已生成完成！")
        logger.info(f"📝 JSON路径：{json_abs_path}")
        logger.info("💡 可使用该JSON文件配合 check-format/apply-format 模式执行操作")

    elif args.mode == "check-format":
        # 模式2：校验格式 → 直接读取用户指定的json_abs_path
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
        # 模式3：格式化 → 直接读取用户指定的json_abs_path
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


if __name__ == "__main__":
    main()
