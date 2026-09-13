#! /usr/bin/env python
# @File    : fix_training_labels.py

"""批量修正预标注训练数据的系统性标签错误。

发现的三类错误（由 BERT 预测 + 后处理产生，纯规则可判）：
1. 目录条目（标题编号 + 连续点导线 + 页码）被标成 heading_level_*，
   应标 heading_mulu（不参与格式化）；
2. 封面论文题目被标 other，应标 document_title；
3. 附录/正文中的代码行被低分误标（acknowledgements_title 等），
   应标 body_text 并取消 needs_review。

用法：
    # 预览（不写盘）
    python scripts/fix_training_labels.py --dir training_data --dry-run
    # 实际修正（同步刷新同名的 .txt 复核视图）
    python scripts/fix_training_labels.py --dir training_data
    # 回滚此前规则修正（按 comment 中的“原：xxx”恢复类别）
    python scripts/fix_training_labels.py --dir training_data --undo
"""

import argparse
import json
import re
import sys
from pathlib import Path

# ===== 模式 =====
# 目录条目：编号标题 + ≥3 个连续点 + 可选页码（正文标题不会长这样）
TOC_ITEM = re.compile(r"^\s*\d+(\.\d+){0,2}\s+.{2,60}[.．。·]{3,}")

# 目录内附录条目（“附录A.........”，非数字开头）
TOC_APPENDIX = re.compile(r"^(附\s*录|Abstract|ABSTRACT|摘\s*要)\S*[.．。·]{4,}")

# 目录标题（带任意空白）
TOC_TITLE = re.compile(r"^目\s*录\s*$")

# 目录条目尾随页码（1~3 位数字）或罗马页码（“摘要 I”/“Abstract II”）
TOC_PAGE = re.compile(r"\d{1,3}\s*$")
TOC_ROMAN = re.compile(r"^(摘\s*要|Abstract|ABSTRACT)\s*\.*\s*[IVX]+\s*$")

# 代码行特征（硬编码代码段）
CODE_LINE = re.compile(
    r"(^\s*[{}$;]\s*$"          # 孤立大括号/分号
    r"|^\s*}\s*;\s*$"              # };
    r"|^\s*}\s*else\s*\{?\s*$"     # }else { / }else{
    r"|^\s*#\s*(include|define)\b"   # 预处理指令
    r"|^\s*([vV]oid|int|char|float|double|static|unsigned|signed|long|short|bool|"
    r"uint\d+_t|u8|u16|u32)\s+\w+\s*\([^)]*\)\s*[;{]*\s*$"  # 函数定义/声明头
    r"|^\s*\w+\s*\([^)]*\)\s*;\s*$"  # 函数调用
    r"|^\s*[A-Za-z_]\w*\s+\w+\s*;\s*$"  # 类型声明（pthread_t thread_id;）
    r"|^\s*\w+\s*=\s*[^;]+;\s*$"    # 赋值语句
    r"|^\s*\w+\s*--\s*;\s*$"         # 自减语句（Get_Time1--;）
    r"|^\s*(function|var)\s+\w*\s*\("  # JS function/var
    r"|^\s*(if|elif|else|for|while|switch|case|do|return|break|continue|def|class|"
    r"try|except|finally|with)\b.*[:{;]\s*$"   # 控制流/函数定义（行尾 : { ;）
    r"|^\s*import\s+\w+"
    r"|^\s*from\s+\w+\s+import\b"
    r"|;\s*$)")                          # 分号结尾的短行兜底


# 封面装饰文本（不是论文题目，按包含匹配，避免漏掉“（论文）”变体）
COVER_DECOR = ("毕业设计说明书", "毕业论文")

# 封面字段（姓名/学院等，不是题目）
COVER_FIELD = re.compile(r"(学\s*生|学\s*院|专\s*业|指导教师|职\s*称|学\s*号|姓\s*名)")

# 封面日期行（XXXX 年 X 月 X 日 / 2025 年 6 月 6 日 / 年 月 日，日月间可能夹示例字符）
COVER_DATE = re.compile(r"年.{0,6}月.{0,6}日")

# 模板占位符题名（如 xxxxxxxxxxxxxxxxxxxxx）
X_PLACEHOLDER = re.compile(r"^x+\s*$", re.IGNORECASE)

# 全英文短句（用于摘要区英文题名“Design of ...”归位）
EN_TITLE = re.compile(r"^[A-Za-z][A-Za-z0-9 ,;:-]{7,}$")

# 含中文字符（用于中文关键词归位与正文确认）
CN_KW = re.compile(r"[\u4e00-\u9fff]")

# 文档标题：封面区长 8~45 的短句（<8 排除“设计与实现”等题名碎片）
TITLE_MIN_LEN = 8
TITLE_MAX_LEN = 45


def _is_code_line(text: str) -> bool:
    s = text.strip()
    if not s:
        return False
    return bool(CODE_LINE.match(s))


def fix_one(path: Path, dry_run: bool) -> int:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    n_fix = 0

    def apply(i, new_cat, reason):
        nonlocal n_fix
        old = data[i]["category"]
        if old == new_cat:
            return
        if not dry_run:
            data[i]["category"] = new_cat
            data[i]["score"] = 1.0
            data[i]["needs_review"] = False
            data[i]["comment"] = f"人工规则修正：{reason}（原：{old}）"
        n_fix += 1
        print(f"  [{i:>4}] {old:<28} -> {new_cat:<18} | {(data[i].get('paragraph') or '').strip()[:36]}")

    # 1) 目录区条目 → heading_mulu（“目录”标题之后、正文第一章之前）
    in_toc = False
    for i, d in enumerate(data):
        cat = d["category"]
        text = (d.get("paragraph") or "").strip()
        # 进入目录区：以“目录”二字独立成段（无论模型给了什么类别）或已标 heading_mulu
        if not in_toc and (TOC_TITLE.match(text) or cat == "heading_mulu"):
            in_toc = True
            if cat != "heading_mulu" and TOC_TITLE.match(text):
                apply(i, "heading_mulu", "目录标题")
        if not in_toc or not text:
            continue
        # 目录区结束：无页码的一级标题（正文开始）或长正文段
        if cat == "heading_level_1" and not TOC_PAGE.search(text) and not TOC_ITEM.match(text):
            in_toc = False
            continue
        if cat == "body_text" and len(text) > 120:
            in_toc = False
            continue
        # 目录区内：已标 body/图片/其它杂项的不动，其余条目统一转 heading_mulu
        if cat in ("body_text", "figure_image", "other", "heading_mulu"):
            continue
        if TOC_ITEM.match(text) or TOC_PAGE.search(text) or TOC_ROMAN.match(text):
            apply(i, "heading_mulu", "目录条目")

    # 2) 封面论文题目 → document_title
    abstract_idx = None
    for i, d in enumerate(data):
        text = (d.get("paragraph") or "").strip()
        if re.match(r"^(摘\s*要|Abstract|ABSTRACT)", text):
            abstract_idx = i
            break
    if abstract_idx:
        for i in range(abstract_idx):
            d = data[i]
            text = (d.get("paragraph") or "").strip()
            # 候选类别：other（旧模型兜底）或 abstract_chinese_title（题名被误标）
            is_title_mislabeled = d["category"] == "abstract_chinese_title" and "摘" not in text
            if not (d["category"] == "other" or is_title_mislabeled) or not text:
                continue
            if any(dec in text for dec in COVER_DECOR) or COVER_FIELD.search(text):
                continue  # 装饰文本/字段，不是题目
            if COVER_DATE.search(text) or X_PLACEHOLDER.match(text):
                continue  # 日期行/占位符，保持 other
            if text.endswith(("：", ":", "。", "；")):
                continue
            if TITLE_MIN_LEN <= len(text) <= TITLE_MAX_LEN:
                apply(i, "document_title", "封面论文题目")
                break

    # 3) 代码行 → body_text（仅处理 needs_review 或非 body 标签的段）
    for i, d in enumerate(data):
        text = (d.get("paragraph") or "").strip()
        if not text or d["category"] in ("body_text", "figure_image", "heading_mulu"):
            continue
        if _is_code_line(text):
            apply(i, "body_text", "代码行")
            continue

    # 4) 目录点导线兜底（正文不会用点导线填充，可全文判）
    for i, d in enumerate(data):
        if d["category"] not in (
            "heading_level_1", "heading_level_2", "heading_level_3",
            "heading_fulu", "body_text", "references_title",
            "acknowledgements_title", "abstract_chinese_title",
            "abstract_english_title", "caption_figure", "caption_table",
        ):
            continue
        text = (d.get("paragraph") or "").strip()
        if TOC_ITEM.match(text) or TOC_APPENDIX.match(text):
            apply(i, "heading_mulu", "目录点导线兜底")

    # 5) 中文关键词归位（模型偶把中文关键词标成 keywords_english）
    for i, d in enumerate(data):
        if d["category"] == "keywords_english":
            t = (d.get("paragraph") or "").strip()
            if CN_KW.search(t) and "Keywords" not in t[:12]:
                apply(i, "keywords_chinese", "中文关键词归位")

    # 6) 摘要区英文题名归位（heading_fulu 且全英文无页码短句）
    for i, d in enumerate(data):
        if d["category"] == "heading_fulu":
            t = (d.get("paragraph") or "").strip()
            if EN_TITLE.match(t) and not TOC_PAGE.search(t):
                apply(i, "abstract_english_title", "英文题名归位")

    # 7) 参数字段行（如“功率因数：”）被标标题 → body_text
    for i, d in enumerate(data):
        if d["category"] in ("heading_level_2", "heading_level_3"):
            t = (d.get("paragraph") or "").strip()
            if 2 <= len(t) <= 12 and t.endswith("：") and not re.search(r"\d", t):
                apply(i, "body_text", "参数字段行")

    # 8) 低置信 needs_review 段复核：判别性确认后清理标记（类别不变）
    n_ok = 0
    for i, d in enumerate(data):
        if not d.get("needs_review"):
            continue
        t = (d.get("paragraph") or "").strip()
        cat = d["category"]
        ok = (
            (cat == "body_text" and _is_code_line(t))
            or (cat == "body_text" and len(t) >= 10 and CN_KW.search(t))
            or (cat == "body_text" and 1 <= len(t) <= 12 and CN_KW.search(t) and t.endswith("："))
            or (cat == "caption_table" and len(t) <= 12)
            or (cat in ("heading_level_1", "heading_level_2") and 2 <= len(t) <= 40)
        )
        if ok:
            if not dry_run:
                d["needs_review"] = False
                d["comment"] = "复核确认无误"
            n_ok += 1
    if n_ok:
        print(f"  复核清理 {n_ok} 处（needs_review=False，类别不变）{'（未写盘）' if dry_run else ''}")

    if not dry_run:
        Path(path).write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        write_txt_view(path, data)
    return n_fix


def write_txt_view(path: Path, data: list[dict]) -> None:
    """按修正后的标签刷新同名 .txt 复核视图（与 dump 脚本格式一致）。"""
    lines = []
    for i, it in enumerate(data):
        text = (it.get("paragraph") or "").replace("\n", "\\n")
        if len(text) > 62:
            text = text[:62] + "…"
        flag = " *" if it.get("needs_review") else ""
        lines.append(f"{i:>4} | {it['category']:<30} | {it.get('score', 0):.2f}{flag} | {text}")
    path.with_suffix(".txt").write_text("\n".join(lines), encoding="utf-8")


def undo_one(path: Path, dry_run: bool) -> int:
    """回滚本脚本此前写入的修正：按 comment 中“（原：xxx）”恢复类别。"""
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    n = 0
    for it in data:
        c = it.get("comment") or ""
        if c.startswith("人工规则修正"):
            m = re.search(r"（原：([^）]+)）", c)
            if m:
                it["category"] = m.group(1)
                it["comment"] = ""
                n += 1
    if n and not dry_run:
        Path(path).write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        write_txt_view(path, data)
    return n


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dir", default="training_data", help="训练数据目录")
    parser.add_argument("--dry-run", action="store_true", help="预览不写盘")
    parser.add_argument("--undo", action="store_true", help="回滚此前规则修正")
    args = parser.parse_args()

    files = sorted(Path(args.dir).glob("*.json"))
    if not files:
        print(f"目录下没有 JSON：{args.dir}", file=sys.stderr)
        sys.exit(1)

    total = 0
    for path in files:
        n = undo_one(path, args.dry_run) if args.undo else fix_one(path, args.dry_run)
        if n:
            total += n
            action = "回滚" if args.undo else "修正"
            print(f"{'[dry-run] ' if args.dry_run else ''}{path.name}：{action} {n} 处")
    action = "回滚" if args.undo else "修正"
    print(f"\n共{action} {total} 处{'（未写盘）' if args.dry_run else ''}")


if __name__ == "__main__":
    main()