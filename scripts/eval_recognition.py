#! /usr/bin/env python
# @File    : eval_recognition.py
"""用人工标注的 JSON 评估段落识别准确率（零第三方依赖）。

pred / truth 均为 ``DocxBase.parse()`` 输出的段落列表（list[dict]，至少含
``category`` 字段）。默认按段落顺序（索引）对齐——因为两者都是同一篇文档、
同一套解析流程产出的有序列表。若两边段落数不一致，取较短长度对齐并给出警告。

用法：
    python scripts/eval_recognition.py --pred pred.json --truth truth.json

生成 pred.json（示例）：
    from wordformat.classify.tag import set_tag_main
    import json
    data = set_tag_main(docx_path="paper.docx")
    json.dump(data, open("pred.json", "w", encoding="utf-8"), ensure_ascii=False)
"""

import argparse
import json
from collections import Counter


def load(path: str) -> list:
    """读取 JSON 文件（应为 list[dict]）。"""
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    if not isinstance(data, list):
        raise ValueError(f"{path} 应为段落列表 list[dict]，实际为 {type(data).__name__}")
    return data


def align(pred: list, truth: list) -> tuple[list, list, int]:
    """按索引对齐 pred / truth，返回 (y_true, y_pred, n_aligned)。"""
    n = min(len(pred), len(truth))
    if len(pred) != len(truth):
        print(f"[警告] 预测 {len(pred)} 段 vs 标注 {len(truth)} 段，按前 {n} 段对齐")
    y_pred = [str(pred[i].get("category", "")) for i in range(n)]
    y_true = [str(truth[i].get("category", "")) for i in range(n)]
    return y_true, y_pred, n


def prf(y_true: list, y_pred: list) -> tuple[dict, float]:
    """计算每个类别的 precision/recall/f1/support 及整体 accuracy。"""
    tp, fp, fn = Counter(), Counter(), Counter()
    for t, p in zip(y_true, y_pred, strict=False):
        if t == p:
            tp[t] += 1
        else:
            fp[p] += 1
            fn[t] += 1

    per_label = {}
    for lab in sorted(set(y_true) | set(y_pred)):
        p = tp[lab]
        precision = p / (p + fp[lab]) if (p + fp[lab]) else 0.0
        recall = p / (p + fn[lab]) if (p + fn[lab]) else 0.0
        f1 = (
            2 * precision * recall / (precision + recall)
            if (precision + recall)
            else 0.0
        )
        per_label[lab] = {
            "precision": precision,
            "recall": recall,
            "f1": f1,
            "support": p + fn[lab],
        }
    accuracy = sum(tp.values()) / len(y_true) if y_true else 0.0
    return per_label, accuracy


def format_report(per_label: dict, accuracy: float, total: int) -> str:
    """把 prf 结果排版成类似 sklearn classification_report 的文本。"""
    header = (
        f"{'':>34}{'precision':>11}{'recall':>11}{'f1-score':>11}{'support':>11}"
    )
    lines = [header, "-" * len(header)]
    macro = {"precision": 0.0, "recall": 0.0, "f1": 0.0}
    support_sum = 0
    for lab, m in per_label.items():
        lines.append(
            f"{lab:>34}{m['precision']:>11.3f}{m['recall']:>11.3f}"
            f"{m['f1']:>11.3f}{m['support']:>11}"
        )
        for k in macro:
            macro[k] += m[k]
        support_sum += m["support"]
    n_lab = len(per_label) or 1
    lines.append("-" * len(header))
    lines.append(f"{'accuracy':>34}{'':>22}{accuracy:>11.3f}{total:>11}")
    lines.append(
        f"{'macro avg':>34}{macro['precision'] / n_lab:>11.3f}"
        f"{macro['recall'] / n_lab:>11.3f}{macro['f1'] / n_lab:>11.3f}"
        f"{support_sum:>11}"
    )
    return "\n".join(lines)


def confusion_matrix(y_true: list, y_pred: list) -> tuple[list, list]:
    """返回 (labels, matrix)，matrix[i][j] = 真实 i 被预测为 j 的次数。"""
    labels = sorted(set(y_true) | set(y_pred))
    idx = {lab: i for i, lab in enumerate(labels)}
    matrix = [[0] * len(labels) for _ in labels]
    for t, p in zip(y_true, y_pred, strict=False):
        matrix[idx[t]][idx[p]] += 1
    return labels, matrix


def format_confusion(labels: list, matrix: list) -> str:
    """用列索引 + 图例的方式排版混淆矩阵，避免长标签错位。"""
    col_w = max(5, len(str(len(labels))) + 2)
    lines = ["列索引对应标签："]
    for i, lab in enumerate(labels):
        lines.append(f"  [{i}] {lab}")
    lines.append("")
    row_head = "true\\pred"
    lines.append(
        f"{row_head:>10}" + "".join(f"{i:>{col_w}}" for i in range(len(labels)))
    )
    for i in range(len(labels)):
        lines.append(
            f"{('[' + str(i) + ']'):>10}"
            + "".join(f"{v:>{col_w}}" for v in matrix[i])
        )
    return "\n".join(lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="评估段落识别准确率")
    parser.add_argument("--pred", required=True, help="解析/后处理输出 JSON")
    parser.add_argument("--truth", required=True, help="人工标注 JSON")
    args = parser.parse_args()

    pred = load(args.pred)
    truth = load(args.truth)

    y_true, y_pred, n = align(pred, truth)
    print(f"对齐段落数：{n}")
    print(f"预测总段落：{len(pred)}，标注总段落：{len(truth)}")

    per_label, accuracy = prf(y_true, y_pred)

    print("\n===== Classification Report =====")
    print(format_report(per_label, accuracy, n))

    print("\n===== Confusion Matrix =====")
    labels, matrix = confusion_matrix(y_true, y_pred)
    print(format_confusion(labels, matrix))

    print("\n===== Top 错误转移 =====")
    errors = Counter()
    for t, p in zip(y_true, y_pred, strict=False):
        if t != p:
            errors[(t, p)] += 1
    if not errors:
        print("（无错例）")
    else:
        for (t, p), c in errors.most_common(10):
            print(f"{t:>34} → {p:<34} {c}")

    print(f"\n整体准确率：{accuracy:.4f}")


if __name__ == "__main__":
    main()
