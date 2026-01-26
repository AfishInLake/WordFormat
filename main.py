#! /usr/bin/env python
# @Time    : 2026/1/18 13:22
# @Author  : afish
# @File    : main.py
import argparse
import os
import sys
from pathlib import Path

from loguru import logger

# 导入核心函数
from src.set_style import auto_format_thesis_document
from src.set_tag import main as set_tag_main


def validate_file(path: str, file_type: str = "文件") -> str:
    """简单校验文件是否存在"""
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise argparse.ArgumentTypeError(f"{file_type}不存在: {abs_path}")
    return abs_path


def get_json_path(docx_path: str, json_dir: str = "tmp/") -> str:
    """根据docx路径推导JSON路径"""
    docx_path = Path(docx_path)
    return os.path.join(json_dir, f"{docx_path.stem}.json")


if __name__ == "__main__":
    # 1. 创建参数解析器（支持子命令/模式选择）
    parser = argparse.ArgumentParser(
        description="学位论文格式自动校验工具（多模式控制）"
    )

    # 2. 添加全局参数（所有模式共享）
    parser.add_argument(
        "--docx",
        "-d",
        required=True,
        type=lambda x: validate_file(x, "Word文档"),
        help="待处理的Word文档路径（必填），例如：tmp/毕业设计说明书.docx",
    )
    parser.add_argument(
        "--json-dir",
        "-j",
        default="tmp/",
        help="JSON文件保存/读取目录（可选，默认：tmp/）",
    )

    # 3. 添加模式选择参数（核心：区分不同执行场景）
    subparsers = parser.add_subparsers(
        dest="mode",
        required=True,
        help="执行模式选择：\n  "
        "generate-json: 仅生成JSON文件（不执行校验）\n  "
        "check-format: 仅执行格式校验（需已生成/修改好JSON）\n  "
        "full-pipeline: 生成JSON→手动编辑→执行校验",
    )

    # 3.1 模式1：仅生成JSON
    parser_gen = subparsers.add_parser(
        "generate-json", help="仅生成文档结构JSON文件，不执行格式校验"
    )

    # 3.2 模式2：仅执行格式校验（需指定JSON和配置）
    parser_check = subparsers.add_parser(
        "check-format", help="仅执行格式校验（需先手动准备好JSON文件）"
    )
    parser_check.add_argument(
        "--config",
        "-c",
        required=True,
        type=lambda x: validate_file(x, "配置文件"),
        help="格式配置YAML路径（必填），例如：test/undergrad_thesis.yaml",
    )
    parser_check.add_argument(
        "--json",
        "-jf",
        help="指定已修改好的JSON文件路径（可选，默认使用--json-dir下的同名JSON）",
    )
    parser_check.add_argument(
        "--output",
        "-o",
        default="output/",
        help="校验后文档保存目录（可选，默认：output/）",
    )

    # 3.3 模式3：完整流程（生成JSON→手动编辑→校验）
    parser_full = subparsers.add_parser(
        "full-pipeline", help="生成JSON→手动编辑→执行格式校验（完整流程）"
    )
    parser_full.add_argument(
        "--config",
        "-c",
        default="test/undergrad_thesis.yaml",
        type=lambda x: validate_file(x, "配置文件"),
        help="格式配置YAML路径（可选，默认：test/undergrad_thesis.yaml）",
    )
    parser_full.add_argument(
        "--output",
        "-o",
        default="output/",
        help="校验后文档保存目录（可选，默认：output/）",
    )

    # 4. 解析参数
    args = parser.parse_args()
    docx_abs_path = os.path.abspath(args.docx)
    default_json_path = get_json_path(args.docx, args.json_dir)

    # 5. 根据不同模式执行对应逻辑
    if args.mode == "generate-json":
        # 模式1：仅生成JSON
        logger.info("=" * 60)
        logger.info("📌 执行模式：仅生成JSON文件")
        logger.info(f"📄 源Word文档：{docx_abs_path}")
        logger.info(f"📋 生成的JSON路径：{default_json_path}")
        logger.info("=" * 60)

        set_tag_main(docx_path=args.docx, json_save_path=str(default_json_path))
        logger.info("\n✅ JSON文件已生成完成！")
        logger.info(f"📝 JSON路径：{os.path.abspath(default_json_path)}")
        logger.info("💡 你可手动修改该JSON文件后，使用 check-format 模式执行校验")

    elif args.mode == "check-format":
        # 模式2：仅执行格式校验
        logger.info("=" * 60)
        logger.info("📌 执行模式：仅执行格式校验")
        # 确定JSON文件路径（优先使用指定的--json，否则用默认路径）
        json_path = args.json if args.json else default_json_path
        json_abs_path = os.path.abspath(json_path)
        # 校验JSON文件是否存在
        validate_file(json_abs_path, "JSON文件")

        logger.info(f"📄 源Word文档：{docx_abs_path}")
        logger.info(f"📋 使用的JSON文件：{json_abs_path}")
        logger.info(f"⚙️  配置文件：{args.config}")
        logger.info(f"💾 输出目录：{args.output}")
        logger.info("=" * 60)

        auto_format_thesis_document(
            jsonpath=json_abs_path,
            docxpath=args.docx,
            configpath=args.config,
            savepath=args.output,
        )
        logger.info(f"\n✅ 格式校验完成！校验后文档已保存至：{args.output}")

    elif args.mode == "full-pipeline":
        # 模式3：完整流程（生成→编辑→校验）
        logger.info("=" * 60)
        logger.info("📌 执行模式：完整流程（生成JSON→手动编辑→格式校验）")
        logger.info(f"📄 源Word文档：{docx_abs_path}")
        logger.info(f"📋 生成的JSON路径：{default_json_path}")
        logger.info(f"⚙️  配置文件：{args.config}")
        logger.info(f"💾 输出目录：{args.output}")
        logger.info("=" * 60)

        # 第一步：生成JSON
        set_tag_main(args.docx, str(default_json_path))
        json_abs_path = os.path.abspath(default_json_path)
        logger.info(f"\n✅ 第一步完成：JSON文件已生成 → {json_abs_path}")

        # 第二步：手动编辑确认
        logger.info("\n" + "=" * 60)
        logger.info("⚠️  请手动修改以下JSON文件后继续：")
        logger.info(f"📝 JSON文件路径：{json_abs_path}")
        logger.info("\n修改完成后输入 'y' 继续，输入 'n' 退出")
        logger.info("=" * 60)
        while True:
            user_input = input("\n是否已修改完成并继续？(y/n): ").strip().lower()
            if user_input in ["y", "yes"]:
                logger.info("\n✅ 确认继续，开始执行格式校验...")
                break
            elif user_input in ["n", "no"]:
                logger.info("\n❌ 用户选择退出程序")
                sys.exit(0)
            else:
                logger.info("⚠️  输入无效，请输入 y/yes 或 n/no")

        # 第三步：执行格式校验
        auto_format_thesis_document(
            jsonpath=json_abs_path,
            docxpath=args.docx,
            configpath=args.config,
            savepath=args.output,
        )
        logger.info(f"\n🎉 完整流程执行完成！校验后文档已保存至：{args.output}")
