#! /usr/bin/env python
# @File    : dump_training_data.py

"""批量生成训练标注数据。

对指定目录下所有 docx 论文跑现有识别流程（BERT + 后处理），
为每篇论文输出两个文件到 training_data/：

1. {论文名}.json —— 机器预测结果，含 paragraph/category/score/comment 字段，
   可直接导入 DocTagChecker UI 人工修正标签后导出（即半自动标注）；
2. {论文名}.txt —— 紧凑文本视图（索引 | 类别 | 置信度 | 内容），
   便于快速浏览错误。

用法：
    python scripts/dump_training_data.py --dir /Users/afish/Desktop/论文 --out training_data
"""

import argparse
import json
import sys
from pathlib import Path

from loguru import logger

from wordformat.base import DocxBase

DEFAULT_CONFIG = "example/undergrad_thesis.yaml"


def dump_one(docx_path: Path, out_dir: Path, config_path: str) -> dict:
    """处理单篇论文，输出预测 JSON + 文本视图。"""
    out_dir.mkdir(parents=True, exist_ok=True)

    base = DocxBase(str(docx_path), config_path)
    result = base.parse()

    stem = docx_path.stem
    json_path = out_dir / f"{stem}.json"
    txt_path = out_dir / f"{stem}.txt"

    json_path.write_text(
        json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    lines = []
    for i, (para, r) in enumerate(zip(base.document.paragraphs, result)):
        text = para.text.replace("\n", "\\n")
        if len(text) > 62:
            text = text[:62] + "…"
        flag = " *" if r.get("needs_review") else ""
        lines.append(
            f"{i:>4} | {r['category']:<30} | {r.get('score', 0):.2f}{flag} | {text}"
        )
    txt_path.write_text("\n".join(lines), encoding="utf-8")

    return {
        "name": docx_path.name,
        "paras": len(result),
        "needs_review": sum(1 for r in result if r.get("needs_review")),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--dir", required=True, help="docx 论文目录")
    parser.add_argument(
        "--out", default="training_data", help="输出目录（默认 training_data）"
    )
    parser.add_argument(
        "--config", default=DEFAULT_CONFIG, help="YAML 配置文件路径"
    )
    parser.add_argument(
        "--skip-existing",
        action="store_true",
        help="跳过已生成的 JSON（续跑剩余论文）",
    )
    args = parser.parse_args()

    src_dir = Path(args.dir)
    if not src_dir.is_dir():
        logger.error(f"目录不存在：{src_dir}")
        sys.exit(1)

    out_dir = Path(args.out)
    docx_files = sorted(src_dir.glob("*.docx"))
    if not docx_files:
        logger.error(f"目录下没有 docx 文件：{src_dir}")
        sys.exit(1)

    skip = 0
    if args.skip_existing:
        todo = []
        for path in docx_files:
            if (out_dir / f"{path.stem}.json").exists():
                skip += 1
            else:
                todo.append(path)
        docx_files = todo
        if skip:
            logger.info(f"已跳过 {skip} 篇已生成，剩余 {len(docx_files)} 篇待处理")
        if not docx_files:
            logger.info("全部已生成，无需处理")
            sys.exit(0)

    logger.info(f"共处理 {len(docx_files)} 篇论文，输出到 {out_dir.resolve()}")

    ok, fail = 0, 0
    n_review = 0
    for fi, path in enumerate(docx_files, 1):
        try:
            info = dump_one(path, out_dir, args.config)
            n_review += info["needs_review"]
            ok += 1
            print(
                f"[{fi:>2}/{len(docx_files)}] OK 段={info['paras']:<5} "
                f"需复核={info['needs_review']:<4} {info['name'][:36]}",
                flush=True,
            )
        except Exception as e:  # noqa: BLE001
            fail += 1
            print(
                f"[{fi:>2}/{len(docx_files)}] FAIL {path.name[:36]}: "
                f"{type(e).__name__}: {e}",
                flush=True,
            )

    print("\n===== 汇总 =====")
    print(f"成功：{ok}  失败：{fail}  总段落数（含需复核 {n_review} 段）")

    if ok > 0:
        print("\n标注工作流：")
        print("  1. 打开 WordFormatUI 的『文档标签核对工具』")
        print("  2. 导入 training_data/*.json（每个文件对应一篇论文）")
        print("  3. 逐段核对并修改错误标签（低分段落已用 * 在 txt 视图中标出）")
        print("  4. 导出 JSON 放回 training_data/（同名覆盖）")
        print("  5. 全部修正后运行 scripts/train_small_bert.py 开始训练")


if __name__ == "__main__":
    main()