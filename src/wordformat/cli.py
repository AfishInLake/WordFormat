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

from wordformat.set_style import auto_format_thesis_document
from wordformat.set_tag import set_tag_main
from wordformat.settings import VERSION
from wordformat.tree import print_tree


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
    """输出所有可配置字段及其说明"""
    from wordformat.config.datamodel import (
        CaptionNumberingConfig,
        GlobalFormatConfig,
        KeywordsConfig,
        NumberingConfig,
        NumberingLevelConfig,
        WarningFieldConfig,
    )

    def _fields(cls):
        return [(n, i) for n, i in cls.model_fields.items()]

    def _desc(info):
        return (info.description or "").replace("\n", " ")

    def _default(info):
        if info.default is not None:
            return repr(info.default)
        if info.default_factory is not None:
            return "(子配置)"
        return "—"

    def _print_fields(fields, indent=2):
        prefix = " " * indent
        for name, info in fields:
            print(f"{prefix}{name:<28s} {_desc(info):<40s} 默认: {_default(info)}")

    gf = _fields(GlobalFormatConfig)
    wf = _fields(WarningFieldConfig)
    kw_extra = [
        (n, i) for n, i in _fields(KeywordsConfig) if n not in {f[0] for f in gf}
    ]
    nc = _fields(NumberingConfig)
    ncp = _fields(CaptionNumberingConfig)
    nl = _fields(NumberingLevelConfig)

    # 段落节点名列表（从 NodeConfigRoot 直接读）
    paragraph_nodes = [
        ("global_format", "全局基础格式"),
        ("abstract.chinese.chinese_title", "中文摘要标题"),
        ("abstract.chinese.chinese_content", "中文摘要正文"),
        ("abstract.english.english_title", "英文摘要标题"),
        ("abstract.english.english_content", "英文摘要正文"),
        ("abstract.keywords.chinese", "中文关键词"),
        ("abstract.keywords.english", "英文关键词"),
        ("headings.level_1", "一级标题"),
        ("headings.level_2", "二级标题"),
        ("headings.level_3", "三级标题"),
        ("body_text", "正文"),
        ("figures", "图注"),
        ("tables", "表注"),
        ("references.title", "参考文献标题"),
        ("references.content", "参考文献内容"),
        ("acknowledgements.title", "致谢标题"),
        ("acknowledgements.content", "致谢内容"),
    ]

    # 每个节点的额外字段
    extras = {
        "figures": [("caption_prefix", "图注编号前缀", "'图'")],
        "tables": [
            ("caption_prefix", "表注编号前缀", "'表'"),
            ("content", "表格内容格式（子配置，字段同 global_format）", "(子配置)"),
        ],
        "abstract.keywords.chinese": [(n, _desc(i), _default(i)) for n, i in kw_extra],
        "abstract.keywords.english": [(n, _desc(i), _default(i)) for n, i in kw_extra],
    }

    print("config.yaml 完整字段参考\n")

    for path, label in paragraph_nodes:
        print(f"[{path}] — {label}")
        _print_fields(gf)
        for name, desc, default in extras.get(path, []):
            print(f"  {name:<28s} {desc:<40s} 默认: {default}")
        print()

    print("[style_checks_warning] — 格式警告开关")
    _print_fields(wf)
    print()

    print("[numbering] — 自动编号总开关（仅 wordf af 模式生效）")
    for n, i in nc:
        if n not in ("level_1", "level_2", "level_3", "references", "captions"):
            print(f"  {n:<28s} {_desc(i):<40s} 默认: {_default(i)}")
    print()

    print("[numbering.captions] — 题注编号")
    _print_fields(ncp)
    print()

    for key in ("level_1", "level_2", "level_3", "references"):
        print(f"[numbering.{key}] — 编号配置")
        _print_fields(nl)
        print()


def main():
    from wordformat.log_config import setup_logger

    setup_logger()

    print(f"WordFormat v{VERSION}")

    # 无参数直接展示完整帮助
    if len(sys.argv) == 1:
        print("""📝 论文格式自动工具（极简命令）
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

            from wordformat.config.datamodel import NodeConfigRoot

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
