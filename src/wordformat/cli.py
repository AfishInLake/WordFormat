#! /usr/bin/env python
# @Time    : 2026/4/8
# @Author  : afish
# @File    : cli.py
import argparse
import json
import os
import sys
import time
from pathlib import Path

from loguru import logger
from rich.console import Console

from wordformat.classify.tag import set_tag_main
from wordformat.pipeline.orchestrate import auto_format_thesis_document
from wordformat.settings import VERSION
from wordformat.tree import print_tree

console = Console()


def validate_file(
    path: str, name: str, allowed_extensions: list[str] | None = None
) -> str:
    """文件合法性校验"""
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise argparse.ArgumentTypeError(f"{name}不存在: {abs_path}")
    if not os.path.isfile(abs_path):
        raise argparse.ArgumentTypeError(f"{name}必须是文件: {abs_path}")
    if allowed_extensions:
        ext = os.path.splitext(abs_path)[1].lower()
        if ext not in allowed_extensions:
            raise argparse.ArgumentTypeError(
                f"{name}必须是以下格式之一: {', '.join(allowed_extensions)}，但得到: {ext or '(无扩展名)'}"
            )
    return abs_path


def _show_config():
    """显示配置字段参考（生成自各 FormatNode 子类的 DEFAULTS 和 YAML 样例）。"""

    console.print("config.yaml 主要节段参考（详请查看 presets/ 目录下的样例配置）\n")
    console.print(
        "  global_format, abstract, headings, body_text, figures,\n"
        "  tables, references, acknowledgements, numbering,\n"
        "  style_checks_warning\n"
    )
    console.print()


def main():
    from wordformat.log_config import setup_logger

    setup_logger()

    console.print(f"WordFormat v{VERSION}")

    # 无参数直接展示完整帮助
    if len(sys.argv) == 1:
        console.print("""📝 论文格式自动工具（极简命令）
==================================================
【极简命令】
wordf gj    生成文档JSON结构
wordf cf    检查格式错误
wordf af    自动格式化论文
wordf tree  查看文档结构树
wordf config  查看所有可配置字段
wordf startapi    启动API服务

【一键示例】
wordf gj -d 论文.docx -c config.yaml -o output/
wordf cf -d 论文.docx -c config.yaml -f output/xxx.json -o output/
wordf af -d 论文.docx -c config.yaml -f output/xxx.json -o output/
wordf tree -f output/xxx.json
wordf config
wordf startapi -H 127.0.0.1 -p 8000
==================================================
""")
        return

    parser = argparse.ArgumentParser(prog="wordf", description="论文格式自动工具")
    parser.add_argument(
        "--version",
        "-V",
        action="version",
        version=f"%(prog)s {VERSION}",
        help="显示版本号",
    )
    subparsers = parser.add_subparsers(dest="mode", required=True)

    # ------------------------------
    # 1. gj = 生成 JSON（自动命名）
    # ------------------------------
    p_gj = subparsers.add_parser("gj", help="生成JSON结构（自动输出到-o目录）")
    p_gj.add_argument(
        "-d",
        required=True,
        type=lambda x: validate_file(x, "文档", [".docx"]),
        help="Word文档路径",
    )
    p_gj.add_argument(
        "-c",
        default=None,
        type=lambda x: validate_file(x, "配置", [".yaml", ".yml"]),
        help="YAML配置路径（可选）",
    )
    p_gj.add_argument("-o", default="output/", help="输出目录（默认output/）")

    # ------------------------------
    # 2. cf = 检查格式
    # ------------------------------
    p_cf = subparsers.add_parser("cf", help="检查格式错误")
    p_cf.add_argument(
        "-d",
        required=True,
        type=lambda x: validate_file(x, "文档", [".docx"]),
        help="Word文档路径",
    )
    p_cf.add_argument(
        "-c",
        required=True,
        type=lambda x: validate_file(x, "配置", [".yaml", ".yml"]),
        help="YAML配置路径",
    )
    p_cf.add_argument(
        "-f",
        required=True,
        type=lambda x: validate_file(x, "JSON文件", [".json"]),
        help="JSON文件路径",
    )
    p_cf.add_argument("-o", default="output/", help="输出目录")

    # ------------------------------
    # 3. af = 格式化
    # ------------------------------
    p_af = subparsers.add_parser("af", help="自动格式化论文")
    p_af.add_argument(
        "-d",
        required=True,
        type=lambda x: validate_file(x, "文档", [".docx"]),
        help="Word文档路径",
    )
    p_af.add_argument(
        "-c",
        required=True,
        type=lambda x: validate_file(x, "配置", [".yaml", ".yml"]),
        help="YAML配置路径",
    )
    p_af.add_argument(
        "-f",
        required=True,
        type=lambda x: validate_file(x, "JSON文件", [".json"]),
        help="JSON文件路径",
    )
    p_af.add_argument("-o", default="output/", help="输出目录")

    # ------------------------------
    # 4. tree = 查看文档结构
    # ------------------------------
    p_tree = subparsers.add_parser("tree", help="查看文档结构树")
    p_tree.add_argument(
        "-f",
        required=True,
        type=lambda x: validate_file(x, "JSON文件", [".json"]),
        help="JSON文件路径",
    )
    p_tree.add_argument("--confidence", action="store_true", help="显示置信度")
    p_tree.add_argument("--index", action="store_true", help="显示节点序号")
    p_tree.add_argument(
        "--filter",
        default=None,
        help="仅显示指定类别（逗号分隔，如 heading_level_1,body_text）",
    )

    # ------------------------------
    # 5. config = 查看可配置字段 / 输出配置模板
    # ------------------------------
    p_config = subparsers.add_parser("config", help="查看可配置字段 / 输出配置模板")
    p_config.add_argument(
        "-o", default=None, help="输出配置模板到文件（如 config.yaml）"
    )

    # ------------------------------
    # 6. startapi = 启动API服务
    # ------------------------------
    p_startapi = subparsers.add_parser("startapi", help="启动API服务")
    p_startapi.add_argument(
        "-H", "--host", default="127.0.0.1", help="API服务地址（默认127.0.0.1）"
    )

    def _validate_port(x):
        v = int(x)
        if v < 1 or v > 65535:
            raise argparse.ArgumentTypeError(f"端口号必须在1-65535之间，但得到: {v}")
        return v

    p_startapi.add_argument(
        "-p",
        "--port",
        type=_validate_port,
        default=8000,
        help="API服务端口（默认8000）",
    )

    # 解析参数
    args = parser.parse_args()

    # 只在需要输出目录的命令中创建目录
    if args.mode in ["gj", "cf", "af"]:
        output_dir = Path(args.o)
        output_dir.mkdir(parents=True, exist_ok=True)

    # ==============================
    # 执行逻辑
    # ==============================
    if args.mode == "gj":
        docx = Path(args.d)
        config = args.c
        # 自动生成 JSON 文件名：原文档名 + 10位时间戳
        doc_name = docx.stem  # 不带后缀的文件名
        timestamp = str(int(time.time()))  # 10位时间戳
        json_path = output_dir / f"{doc_name}_{timestamp}.json"

        logger.info("📌 开始生成文档结构JSON...")
        logger.info(f"📄 源文档：{docx.resolve()}")
        logger.info(f"📁 输出目录：{output_dir.resolve()}")

        # 生成并保存
        data = set_tag_main(docx_path=str(docx), configpath=config)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

        logger.success(f"✅ JSON 已生成：{json_path.resolve()}")
        logger.info("💡 可复制此路径用于 cf/af 命令")

    elif args.mode == "cf":
        logger.info("🔍 开始格式检查...")
        auto_format_thesis_document(
            jsonpath=args.f,
            docxpath=args.d,
            configpath=args.c,
            savepath=args.o,
            check=True,
        )
        logger.success(f"✅ 检查完成！报告保存在：{args.o}")

    elif args.mode == "af":
        logger.info("✏️ 开始自动格式化...")
        auto_format_thesis_document(
            jsonpath=args.f,
            docxpath=args.d,
            configpath=args.c,
            savepath=args.o,
            check=False,
        )
        logger.success(f"✅ 格式化完成！新文件保存在：{args.o}")

    elif args.mode == "tree":
        logger.info("🌳 开始展示文档结构树...")
        filter_categories = None
        if args.filter:
            filter_categories = [c.strip() for c in args.filter.split(",")]
        print_tree(
            node_or_jsonpath=args.f,
            show_confidence=args.confidence,
            show_index=args.index,
            filter_categories=filter_categories,
        )

    elif args.mode == "config":
        if args.o:
            import warnings

            import yaml

            from wordformat.config.models import NodeConfigRoot

            cfg = NodeConfigRoot()
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                data = cfg.model_dump()
            yaml_str = yaml.dump(
                data, default_flow_style=False, allow_unicode=True, sort_keys=False
            )
            with open(args.o, "w", encoding="utf-8") as f:
                f.write(yaml_str)
            logger.success(f"配置模板已输出 → {args.o}")
        else:
            _show_config()

    elif args.mode == "startapi":
        logger.info("🚀 启动API服务...")
        logger.info(f"🌐 服务地址：http://{args.host}:{args.port}")
        logger.info(f"📖 API文档：http://{args.host}:{args.port}/docs")
        logger.info("💡 按 Ctrl+C 停止服务")

        # 动态导入并启动API服务
        import uvicorn

        from wordformat.api import app

        uvicorn.run(
            app,
            host=args.host,
            port=args.port,
            log_config=None,
            access_log=True,
            reload=False,
            use_colors=False,
        )


if __name__ == "__main__":
    main()
