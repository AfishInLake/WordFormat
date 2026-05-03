#!/usr/bin/env python3
"""
WordFormat JSON 标签文件校验脚本

校验 wf gj 生成的 JSON 文件中的 category 字段是否合法，
并输出每个段落的分类结果供人工检查。

用法：
    # 校验 JSON 文件
    python validate_json.py --json output/论文_1234567890.json

    # 输出分类统计
    python validate_json.py --json output/论文_1234567890.json --stats

    # 仅列出可疑分类（置信度低于阈值）
    python validate_json.py --json output/论文_1234567890.json --threshold 0.8
"""

import argparse
import json
import sys
from collections import Counter

# 模型实际输出的所有合法 category（来源于 id2label.json）
# 以及人工修正时使用的辅助 category
VALID_CATEGORIES = {
    "abstract_chinese_title",
    "abstract_chinese_title_content",
    "abstract_english_title",
    "abstract_english_title_content",
    "keywords_chinese",
    "keywords_english",
    "heading_level_1",
    "heading_level_2",
    "heading_level_3",
    "heading_mulu",
    "heading_fulu",
    "references_title",
    "acknowledgements_title",
    "caption_figure",
    "caption_table",
    "body_text",
    "other",
}

# category 中文说明
CATEGORY_LABELS = {
    "abstract_chinese_title": "中文摘要标题",
    "abstract_chinese_title_content": "中文摘要标题+内容",
    "abstract_english_title": "英文摘要标题",
    "abstract_english_title_content": "英文摘要标题+内容",
    "keywords_chinese": "中文关键词",
    "keywords_english": "英文关键词",
    "heading_level_1": "一级标题（章）",
    "heading_level_2": "二级标题（节）",
    "heading_level_3": "三级标题（小节）",
    "heading_mulu": "目录标题",
    "heading_fulu": "附录标题",
    "references_title": "参考文献标题",
    "acknowledgements_title": "致谢标题",
    "caption_figure": "图题注",
    "caption_table": "表题注",
    "body_text": "正文",
    "other": "其他（跳过格式化）",
}


def validate_json(json_path: str, threshold: float = 0.0) -> list[dict]:
    """校验 JSON 文件，返回问题列表"""
    try:
        with open(json_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        return [{"index": -1, "error": f"JSON 语法错误: {e}"}]
    except FileNotFoundError:
        return [{"index": -1, "error": f"文件不存在: {json_path}"}]

    if not isinstance(data, list):
        return [{"index": -1, "error": "JSON 根节点必须是数组"}]

    issues = []
    for i, item in enumerate(data):
        if not isinstance(item, dict):
            issues.append({"index": i, "error": f"第 {i} 项不是字典: {type(item).__name__}"})
            continue

        # 检查 category 字段
        category = item.get("category", "")
        if not category:
            issues.append({"index": i, "error": "缺少 category 字段", "paragraph": item.get("paragraph", "")[:50]})
            continue

        if category not in VALID_CATEGORIES:
            issues.append({
                "index": i,
                "error": f"非法 category: '{category}'",
                "paragraph": item.get("paragraph", "")[:50],
                "hint": f"合法值: {sorted(VALID_CATEGORIES)}"
            })

        # 检查置信度
        score = item.get("score", 0)
        if isinstance(score, (int, float)) and score < threshold:
            issues.append({
                "index": i,
                "warning": f"低置信度: {score}",
                "category": category,
                "paragraph": item.get("paragraph", "")[:80]
            })

    return issues


def print_results(data: list, threshold: float = 0.0):
    """打印分类结果供人工检查"""
    print(f"\n{'='*70}")
    print(f"  JSON 标签检查结果（共 {len(data)} 个段落）")
    print(f"{'='*70}\n")

    for i, item in enumerate(data):
        category = item.get("category", "???")
        score = item.get("score", 0)
        paragraph = item.get("paragraph", "")[:60]
        label = CATEGORY_LABELS.get(category, "未知类型")

        # 低置信度标记
        flag = ""
        if isinstance(score, (int, float)) and score < threshold:
            flag = " ⚠️ 低置信度"
        if category not in VALID_CATEGORIES:
            flag = " ❌ 非法类型"

        print(f"  [{i:3d}] {label:20s}  置信度:{score:.4f}{flag}")
        print(f"        {paragraph}...")
        print()


def print_stats(data: list):
    """打印分类统计"""
    categories = [item.get("category", "unknown") for item in data if isinstance(item, dict)]
    counter = Counter(categories)

    print(f"\n{'='*50}")
    print(f"  分类统计（共 {len(data)} 个段落）")
    print(f"{'='*50}\n")
    print(f"  {'类型':<35s} {'数量':>5s}")
    print(f"  {'-'*42}")
    for cat, count in counter.most_common():
        label = CATEGORY_LABELS.get(cat, cat)
        print(f"  {label:<35s} {count:>5d}")
    print(f"  {'-'*42}")
    print(f"  {'合计':<35s} {len(data):>5d}")


def main():
    parser = argparse.ArgumentParser(description="WordFormat JSON 标签校验工具")
    parser.add_argument("--json", "-j", required=True, help="JSON 文件路径")
    parser.add_argument("--stats", "-s", action="store_true", help="输出分类统计")
    parser.add_argument("--threshold", "-t", type=float, default=0.0,
                        help="低置信度阈值（默认 0.0，即不标记）")
    parser.add_argument("--show-all", action="store_true", help="显示所有段落的分类结果")
    args = parser.parse_args()

    # 加载 JSON
    try:
        with open(args.json, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (json.JSONDecodeError, FileNotFoundError) as e:
        print(f"错误: {e}")
        sys.exit(1)

    # 校验
    issues = validate_json(args.json, args.threshold)

    errors = [i for i in issues if "error" in i]
    warnings = [i for i in issues if "warning" in i]

    if errors:
        print(f"❌ 发现 {len(errors)} 个错误:")
        for item in errors:
            idx = item["index"]
            print(f"  [{idx}] {item['error']}")
            if "paragraph" in item:
                print(f"       内容: {item['paragraph']}")
            if "hint" in item:
                print(f"       {item['hint']}")
        print()

    if warnings:
        print(f"⚠️  发现 {len(warnings)} 个低置信度分类:")
        for item in warnings:
            idx = item["index"]
            print(f"  [{idx}] {item['warning']} | {item.get('category', '')} | {item.get('paragraph', '')}")
        print()

    if not errors and not warnings:
        print(f"✅ JSON 文件校验通过: {args.json}")

    # 统计
    if args.stats:
        print_stats(data)

    # 显示所有结果
    if args.show_all:
        print_results(data, args.threshold)

    if errors:
        sys.exit(1)


if __name__ == "__main__":
    main()
