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


def validate_file(path: str, name: str) -> str:
    """文件合法性校验"""
    abs_path = os.path.abspath(path)
    if not os.path.exists(abs_path):
        raise argparse.ArgumentTypeError(f"{name}不存在: {abs_path}")
    if not os.path.isfile(abs_path):
        raise argparse.ArgumentTypeError(f"{name}必须是文件: {abs_path}")
    return abs_path


def main():
    # 无参数直接展示完整帮助
    if len(sys.argv) == 1:
        print("""
📝 论文格式自动工具（极简命令）
==================================================
【极简命令】
wf gj    生成文档JSON结构（自动生成，无需指定json路径）
wf cf    检查格式错误
wf af    自动格式化论文
wf startapi    启动API服务

【一键示例】
wf gj -d 论文.docx -c config.yaml -o output/
wf cf -d 论文.docx -c config.yaml -f output/xxx.json -o output/
wf af -d 论文.docx -c config.yaml -f output/xxx.json -o output/
wf startapi -H 127.0.0.1 -p 8000

参数：
-d    Word 文档路径（必填）
-c    YAML 配置文件（必填）
-f    JSON 文件路径（仅 cf/af 需要）
-o    输出目录（默认 output/）
-H    API服务地址（默认 127.0.0.1）
-p    API服务端口（默认 8000）
==================================================
""")
        return

    parser = argparse.ArgumentParser(prog="wf", description="论文格式自动工具")
    subparsers = parser.add_subparsers(dest="mode", required=True)

    # ------------------------------
    # 1. gj = 生成 JSON（自动命名）
    # ------------------------------
    p_gj = subparsers.add_parser("gj", help="生成JSON结构（自动输出到-o目录）")
    p_gj.add_argument("-d", required=True, type=lambda x: validate_file(x, "文档"), help="Word文档路径")
    p_gj.add_argument("-c", required=True, type=lambda x: validate_file(x, "配置"), help="YAML配置路径")
    p_gj.add_argument("-o", default="output/", help="输出目录（默认output/）")

    # ------------------------------
    # 2. cf = 检查格式
    # ------------------------------
    p_cf = subparsers.add_parser("cf", help="检查格式错误")
    p_cf.add_argument("-d", required=True, type=lambda x: validate_file(x, "文档"), help="Word文档路径")
    p_cf.add_argument("-c", required=True, type=lambda x: validate_file(x, "配置"), help="YAML配置路径")
    p_cf.add_argument("-f", required=True, type=lambda x: validate_file(x, "JSON文件"), help="JSON文件路径")
    p_cf.add_argument("-o", default="output/", help="输出目录")

    # ------------------------------
    # 3. af = 格式化
    # ------------------------------
    p_af = subparsers.add_parser("af", help="自动格式化论文")
    p_af.add_argument("-d", required=True, type=lambda x: validate_file(x, "文档"), help="Word文档路径")
    p_af.add_argument("-c", required=True, type=lambda x: validate_file(x, "配置"), help="YAML配置路径")
    p_af.add_argument("-f", required=True, type=lambda x: validate_file(x, "JSON文件"), help="JSON文件路径")
    p_af.add_argument("-o", default="output/", help="输出目录")

    # ------------------------------
    # 4. startapi = 启动API服务
    # ------------------------------
    p_startapi = subparsers.add_parser("startapi", help="启动API服务")
    p_startapi.add_argument("-H", "--host", default="127.0.0.1", help="API服务地址（默认127.0.0.1）")
    p_startapi.add_argument("-p", "--port", type=int, default=8000, help="API服务端口（默认8000）")

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
        logger.info(f"💡 可复制此路径用于 cf/af 命令")

    elif args.mode == "cf":
        logger.info("🔍 开始格式检查...")
        auto_format_thesis_document(
            jsonpath=args.f,
            docxpath=args.d,
            configpath=args.c,
            savepath=args.o,
            check=True
        )
        logger.success(f"✅ 检查完成！报告保存在：{args.o}")

    elif args.mode == "af":
        logger.info("✏️ 开始自动格式化...")
        auto_format_thesis_document(
            jsonpath=args.f,
            docxpath=args.d,
            configpath=args.c,
            savepath=args.o,
            check=False
        )
        logger.success(f"✅ 格式化完成！新文件保存在：{args.o}")

    elif args.mode == "startapi":
        logger.info(f"🚀 启动API服务...")
        logger.info(f"🌐 服务地址：http://{args.host}:{args.port}")
        logger.info(f"📖 API文档：http://{args.host}:{args.port}/docs")
        logger.info(f"💡 按 Ctrl+C 停止服务")
        
        # 动态导入并启动API服务
        from wordformat.api import app
        import uvicorn
        
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