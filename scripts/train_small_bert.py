#! /usr/bin/env python
# @File    : train_small_bert.py

"""小 BERT 微调：段落分类训练 + ONNX 导出 + int8 动态量化。

用 uer/chinese_roberta_L-4_H-256（4 层 / 256 维 / ~12M 参数）替代
BART-base 12 层模型作为段落分类器，大幅降低 CPU 占用与内存开销，
适配 2 核 2G 小服务器。

训练数据格式（由 scripts/dump_training_data.py 生成并人工修正后）：
    training_data/*.json —— 每篇论文一个 JSON，元素含 paragraph 与 category。

上下文建模：
    训练与推理统一使用前缀格式 [PREV={上一段标签}] {当前段文本}，
    模型直接学习章节状态转移（如 references_title → references_content）。

用法：
    1. 先安装训练依赖（项目常规 venv 中默认没有）：
       uv pip install torch transformers -p .venv
    2. 训练 + 导出 + 量化一步到位：
       python scripts/train_small_bert.py \
           --data training_data \
           --extra "temp/preds.json:temp/truth.json" \
           --epochs 3
    3. 产物写入 src/wordformat/data/model/：
       - thesis_paragraph_classifier.onnx（fp32）
       - thesis_paragraph_classifier_int8.onnx（int8 量化，推荐部署）
       - model_info.json（推理侧描述文件：标签表/是否使用上下文/量化标记）
       - id2label.json、label2id.json（新标签表）
       - tokenizer.json（新词表）
       - model_small_bert/（transformers 格式 checkpoint，可继续训练）

兼容性：
    - 原 BERT 模型及标签表/词表会备份到 data/model/backup_legacy_bert/，
      如需回滚：把备份文件拷回 data/model 并删除 model_info.json。
    - 推理侧只加载新模型（单模型，无兼容分支），见 onnx_infer.py。
    - 下载国内 HuggingFace 镜像：export HF_ENDPOINT=https://hf-mirror.com
"""

import argparse
import json
import random
import re
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

# ===== 超参数 =====
MODEL_NAME = "uer/chinese_roberta_L-4_H-256"
MAX_LENGTH = 128
EPOCHS = 3
BATCH_SIZE = 32
LEARNING_RATE = 2e-5
WARMUP_RATIO = 0.1
SEED = 42

# 小样本类上采样目标：每类每 epoch 至少出现该次数，
# 避免长尾类（如 abstract_english_content 仅 9 条）被训练成噪声
MIN_MINORITY = 120

# 模型输出目录
MODEL_DIR = Path("src/wordformat/data/model")

# 由 docx 预处理确定性处理的标签（空段/图片段），不进模型训练
PREDEFINED_DIRECT = {"figure_image"}

# 前缀 token：使模型感知上一段标签（None 表示文档首个文本段）
PREV_PREFIX = "[PREV={prev}] "

# 弱类先验纠正：在逆频率权重基础上额外加权（见 train() 中 weights 公式）
WEIGHT_BOOST = {"heading_level_1": 2.0}


def _seed_all() -> None:
    import numpy as np
    import torch

    random.seed(SEED)
    np.random.seed(SEED)
    torch.manual_seed(SEED)


def _pick_device():
    import torch

    if torch.cuda.is_available():
        return "cuda"
    if getattr(torch.backends, "mps", None) and torch.backends.mps.is_available():
        return "mps"
    return "cpu"


# ==================== 数据加载 ====================

def load_annotations(data_dir: str, extra_pairs: list[str]) -> tuple[list[str], list[str]]:
    """加载训练样本，返回 (texts, labels)。

    :param data_dir: 标注 JSON 目录（每篇论文一个文件）
    :param extra_pairs: 附加的数据对 "pred.json:truth.json"（按索引对齐，
                        pred 提供 paragraph，truth 提供人工修正的 category）
    """
    texts: list[str] = []
    labels: list[str] = []
    n_docs = 0

    # 1. 标注目录（JSON 元素含 paragraph 与 category）
    for path in sorted(Path(data_dir).glob("*.json")):
        items = json.loads(Path(path).read_text(encoding="utf-8"))
        n_docs += 1
        prev = None
        for it in items:
            cat = it.get("category", "")
            text = (it.get("paragraph") or "").strip()
            # 空段/图片段由 docx 预处理确定性标记，不进模型
            if not text or cat in PREDEFINED_DIRECT:
                continue
            texts.append(PREV_PREFIX.format(prev=prev) + text)
            labels.append(cat)
            prev = cat

    # 2. 附加对齐数据对（旧格式：pred-text + truth-label）
    for pair in extra_pairs:
        pred_path, truth_path = pair.split(":")
        preds = json.loads(Path(pred_path).read_text(encoding="utf-8"))
        truths = json.loads(Path(truth_path).read_text(encoding="utf-8"))
        assert len(preds) == len(truths), f"{pair} 长度不一致"
        n_docs += 1
        prev = None
        for pred, truth in zip(preds, truths):
            cat = truth.get("category", "")
            text = (pred.get("paragraph") or "").strip()
            if not text or cat in PREDEFINED_DIRECT:
                continue
            texts.append(PREV_PREFIX.format(prev=prev) + text)
            labels.append(cat)
            prev = cat

    return texts, labels


def build_label_map(labels: list[str]) -> tuple[dict[str, int], dict[int, str]]:
    """按出现频率排序构建 label<->id 映射（保存后供推理侧使用）。"""
    order = [c for c, _ in Counter(labels).most_common()]
    if "body_text" in order:
        order.remove("body_text")
        order.insert(0, "body_text")  # 兜底类排最前，与旧模型习惯一致
    label2id = {c: i for i, c in enumerate(order)}
    id2label = {i: c for c, i in label2id.items()}
    return label2id, id2label


# ==================== 编号变体增强 ====================

HEAD_L1 = re.compile(r"^(\d{1,2})\s+(.+)$")
HEAD_L2 = re.compile(r"^(\d{1,2})[.．](\d{1,2})\s+(.+)$")
HEAD_L3 = re.compile(r"^(\d{1,2})[.．](\d{1,2})[.．](\d{1,2})\s+(.+)$")


def _cn_num(n: str) -> str:
    """阿拉伯数字转中文数字（1~99，用于“第一章/一、”变体）。"""
    v = int(n)
    if v <= 0 or v > 99:
        return n
    if v <= 10:
        return "零一二三四五六七八九十"[v]
    tens, ones = divmod(v, 10)
    prefix = "十" if tens == 1 else f"{'零一二三四五六七八九'[tens]}十"
    return prefix + ("零一二三四五六七八九"[ones] if ones else "")


def augment_headings(
    texts: list[str], labels: list[str]
) -> tuple[list[str], list[str]]:
    """标题类编号变体增强。

    标题层级与编号样式无关（"1 绪论"="第一章 绪论"="一、绪论"），
    但小模型见到训练集外的编号格式会被先验带着误判。对标题类样本
    生成常见编号变体，让模型把编号当作噪声而非类别证据。
    """
    aug_texts: list[str] = list(texts)
    aug_labels: list[str] = list(labels)
    for text, cat in zip(texts, labels):
        prev = ""
        body = text
        if text.startswith("[PREV="):
            prev, body = text.split("] ", 1)
            prev += "] "
        variants: list[str] = []
        if cat == "heading_level_1" and (m := HEAD_L1.match(body)):
            n, title = m.group(1), m.group(2)
            variants = [
                f"{n}、{title}",                 # 1、绪论
                f"第{n}章 {title}",               # 第1章 绪论
                f"第{_cn_num(n)}章 {title}",       # 第一章 绪论
                f"{_cn_num(n)}、{title}",         # 一、绪论
                f"{n}. {title}",                 # 1. 绪论（点+空格，与目录点导线不同）
            ]
        elif cat == "heading_level_2" and (m := HEAD_L2.match(body)):
            a, b, title = m.groups()
            variants = [f"{a}．{b} {title}"]       # 全角点
        elif cat == "heading_level_3" and (m := HEAD_L3.match(body)):
            a, b, c, title = m.groups()
            variants = [f"{a}．{b}．{c} {title}"]   # 全角点
        for v in variants:
            if v != body:
                aug_texts.append(prev + v)
                aug_labels.append(cat)
    return aug_texts, aug_labels


def upsample_minority(
    texts: list[str], labels: list[str], floor: int = MIN_MINORITY
) -> tuple[list[str], list[str]]:
    """小样本类复制上采样：每类至少 floor 条（重复出现，suffle 后摊到不同 batch）。"""
    groups: dict[str, list[str]] = {}
    for text, cat in zip(texts, labels):
        groups.setdefault(cat, []).append(text)
    out_texts: list[str] = []
    out_labels: list[str] = []
    for cat, ts in groups.items():
        if len(ts) < floor:
            ts = (ts * ((floor + len(ts) - 1) // len(ts)))[:floor]
        out_texts.extend(ts)
        out_labels.extend([cat] * len(ts))
    return out_texts, out_labels


# ==================== 训练 ====================

def train(texts: list[str], labels: list[str]) -> None:
    import torch
    from torch.utils.data import DataLoader, Dataset
    from transformers import AutoModelForSequenceClassification, AutoTokenizer

    print(f"设备：{_pick_device()}  样本数：{len(texts)}  类别数：{len(set(labels))}")

    label2id, id2label = build_label_map(labels)
    ids = [label2id[c] for c in labels]

    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
    model = AutoModelForSequenceClassification.from_pretrained(
        MODEL_NAME, num_labels=len(label2id)
    )

    # 类别均衡权重（长尾类少样本 → 高权重），弱类再乘 WEIGHT_BOOST
    counts = Counter(ids)
    base_w = {c: max(1.0, len(ids) / (len(counts) * n)) for c, n in counts.items()}
    weights = {c: base_w[c] * WEIGHT_BOOST.get(id2label[c], 1.0) for c in counts}

    class DS(Dataset):
        def __getitem__(self, i):
            tok = tokenizer(
                texts[i],
                truncation=True,
                max_length=MAX_LENGTH,
                padding="max_length",
                return_tensors="pt",
            )
            return (
                tok["input_ids"][0],
                tok["attention_mask"][0],
                torch.tensor(ids[i], dtype=torch.long),
                torch.tensor(weights[ids[i]], dtype=torch.float),
            )

        def __len__(self):
            return len(texts)

    loader = DataLoader(DS(), batch_size=BATCH_SIZE, shuffle=True)

    device = _pick_device()
    model.to(device)
    optimizer = torch.optim.AdamW(model.parameters(), lr=LEARNING_RATE)
    total_steps = EPOCHS * len(loader)
    warmup = int(total_steps * WARMUP_RATIO)
    scheduler = torch.optim.lr_scheduler.LambdaLR(
        optimizer, lambda s: s / max(1, warmup) if s < warmup else 1.0
    )
    loss_fn = torch.nn.CrossEntropyLoss(reduction="none")

    model.train()
    for epoch in range(EPOCHS):
        running = 0.0
        seen = 0
        for batch, (iids, amask, y, w) in enumerate(loader):
            iids, amask, y, w = (
                iids.to(device), amask.to(device), y.to(device), w.to(device)
            )
            optimizer.zero_grad()
            logits = model(iids, attention_mask=amask).logits
            loss = (loss_fn(logits, y) * w).mean()
            loss.backward()
            optimizer.step()
            scheduler.step()
            running += loss.item() * len(y)
            seen += len(y)
            if (batch + 1) % 50 == 0:
                print(
                    f"  epoch {epoch + 1}/{EPOCHS} batch {batch + 1}/{len(loader)} "
                    f"loss={running / max(1, seen):.4f}",
                    flush=True,
                )
        print(f"epoch {epoch + 1}/{EPOCHS} 平均 loss={running / max(1, seen):.4f}")

    # checkpoint 留存（支持续训）
    ckpt_dir = MODEL_DIR / "model_small_bert"
    if ckpt_dir.exists():
        shutil.rmtree(ckpt_dir)
    ckpt_dir.mkdir(parents=True)
    model.save_pretrained(ckpt_dir)
    tokenizer.save_pretrained(ckpt_dir)
    print(f"checkpoint 已保存：{ckpt_dir}")

    # 标签表落地（供推理侧 onnx_infer.py 读取）
    (MODEL_DIR / "id2label.json").write_text(
        json.dumps({str(k): v for k, v in id2label.items()}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (MODEL_DIR / "label2id.json").write_text(
        json.dumps(label2id, ensure_ascii=False, indent=2), encoding="utf-8"
    )


# ==================== ONNX 导出 + int8 量化 ====================

def export_onnx() -> None:
    import torch
    from transformers import AutoModelForSequenceClassification

    ckpt_dir = MODEL_DIR / "model_small_bert"
    model = AutoModelForSequenceClassification.from_pretrained(str(ckpt_dir))
    model.eval()

    # BERT 系图需 token_type_ids 输入；RoBERTa 系（真正无该参数）不需要。
    # 用 architectures 判断，避免依赖 forward 签名（新版 transformers 中
    # Roberta.forward 也接受 token_type_ids 但忽略它，签名判断会误判）。
    arch = (model.config.architectures or [""])[0]
    has_tt = "BertFor" in arch

    class Wrapper(torch.nn.Module):
        def __init__(self, m, has_tt):
            super().__init__()
            self.m = m
            self.has_tt = has_tt

        def forward(self, input_ids, attention_mask):
            kwargs = {"input_ids": input_ids, "attention_mask": attention_mask}
            if self.has_tt:
                kwargs["token_type_ids"] = torch.zeros_like(input_ids)
            return self.m(**kwargs).logits

    wrapper = Wrapper(model, has_tt)
    # torch 2.14 dynamo 导出图带静态形状（量化推断会差），回退 torchscript 导出器
    input_names = ["input_ids", "attention_mask"]
    dynamic_axes = {n: {0: "batch", 1: "seq"} for n in input_names}

    dummy_ids = torch.zeros((1, 16), dtype=torch.long)
    dummy_mask = torch.ones((1, 16), dtype=torch.long)
    fp32_path = MODEL_DIR / "thesis_paragraph_classifier.onnx"

    torch.onnx.export(
        wrapper,
        (dummy_ids, dummy_mask),
        str(fp32_path),
        input_names=input_names,
        output_names=["logits"],
        dynamic_axes=dynamic_axes,
        opset_version=14,
        dynamo=False,
    )
    print(f"ONNX 导出完成：{fp32_path}")

    # dynamo 导出图部分节点带静态形状，量化器内置的严格 shape 推断会报
    # "Inferred shape and existing shape differ"。先以宽松模式推断一次写回，
    # 消除不一致后再量化。
    import onnx

    m = onnx.load(str(fp32_path), load_external_data=False)
    m = onnx.shape_inference.infer_shapes(m, strict_mode=False)
    onnx.save(m, str(fp32_path))

    # int8 动态量化（权重 int8、激活保持 fp32，CPU 提速且体积降到约 1/4）
    from onnxruntime.quantization import QuantType, quantize_dynamic

    int8_path = MODEL_DIR / "thesis_paragraph_classifier_int8.onnx"
    quantize_dynamic(str(fp32_path), str(int8_path), weight_type=QuantType.QInt8)
    print(f"int8 量化完成：{int8_path}")

    size_fp32 = fp32_path.stat().st_size / 1024 / 1024
    size_int8 = int8_path.stat().st_size / 1024 / 1024
    print(f"体积对比：fp32 {size_fp32:.1f}MB → int8 {size_int8:.1f}MB")


# ==================== 产物整理 ====================

def finalize(use_prev_context: bool, quantized: bool) -> None:
    """备份旧模型、写入推理侧描述文件 model_info.json。"""
    old_model = MODEL_DIR / "bert_paragraph_classifier.onnx"
    if old_model.exists():
        bak_dir = MODEL_DIR / "backup_legacy_bert"
        bak_dir.mkdir(exist_ok=True)
        # 旧模型 + 旧标签表 + 旧词表一并备份，保证可完整回滚
        old_model.rename(bak_dir / "bert_paragraph_classifier.onnx")
        for side in ("id2label.json", "label2id.json", "tokenizer.json"):
            src = MODEL_DIR / side
            if src.exists():
                shutil.copy(src, bak_dir / side)
        print(f"旧 BERT 及其标签表/词表已备份：{bak_dir}")
        print("回滚方法：把 backup_legacy_bert/ 内文件拷回 data/model 并删除 model_info.json")

    # 新 tokenizer 落地（覆盖旧词表，与 checkpoint 中保存的一致）
    ckpt_tokenizer = MODEL_DIR / "model_small_bert" / "tokenizer.json"
    if ckpt_tokenizer.exists():
        shutil.copy(ckpt_tokenizer, MODEL_DIR / "tokenizer.json")

    info = {
        "name": "thesis_paragraph_classifier",
        "base_model": MODEL_NAME,
        "use_prev_context": use_prev_context,
        "quantized": quantized,
        "max_length": MAX_LENGTH,
        "onnx": "thesis_paragraph_classifier_int8.onnx"
        if quantized
        else "thesis_paragraph_classifier.onnx",
        "fp32_onnx": "thesis_paragraph_classifier.onnx",
    }
    (MODEL_DIR / "model_info.json").write_text(
        json.dumps(info, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    print(f"model_info.json 已写入：{json.dumps(info, ensure_ascii=False)}")


# ==================== 主流程 ====================

def main() -> None:
    global EPOCHS
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--data", required=True, help="标注 JSON 目录")
    parser.add_argument(
        "--extra",
        nargs="*",
        default=[],
        help="附加数据对 pred.json:truth.json（可多次）",
    )
    parser.add_argument("--epochs", type=int, default=EPOCHS)
    parser.add_argument("--no-context", action="store_true", help="禁用 PREV 上下文前缀")
    parser.add_argument("--no-augment", action="store_true", help="禁用标题编号变体增强")
    parser.add_argument("--skip-quantize", action="store_true", help="跳过 int8 量化")
    parser.add_argument(
        "--skip-train",
        action="store_true",
        help="跳过训练，直接用已有 checkpoint 导出/量化（续跑）",
    )
    args = parser.parse_args()

    if not Path(args.data).is_dir():
        print(f"数据目录不存在：{args.data}", file=sys.stderr)
        sys.exit(1)

    EPOCHS = args.epochs
    use_prev = not args.no_context

    _seed_all()

    if args.skip_train:
        print("跳过训练，使用已有 checkpoint 导出")
    else:
        print("===== 1/3 加载数据 =====")
        texts, labels = load_annotations(args.data, args.extra)
        if len(texts) < 50:
            print(f"样本过少（{len(texts)} < 50），请先完成标注", file=sys.stderr)
            sys.exit(1)
        n_raw = len(texts)
        if not args.no_augment:
            texts, labels = augment_headings(texts, labels)
            print(f"编号变体增强：{n_raw} -> {len(texts)} 条")
        texts, labels = upsample_minority(texts, labels)
        print(f"小类上采样：-> {len(texts)} 条（每类至少 {MIN_MINORITY}）")
        dist = Counter(labels)
        print("标签分布（Top 15）：")
        for cat, cnt in dist.most_common(15):
            print(f"  {cat:<34} {cnt}")

        print("\n===== 2/3 训练 =====")
        train(texts if use_prev else [t.split("] ", 1)[-1] for t in texts], labels)

    print("\n===== 3/3 导出 + 量化 =====")
    export_onnx()
    finalize(use_prev_context=use_prev, quantized=not args.skip_quantize)

    print("\n完成。小模型已就绪，可通过 api/CLI 直接使用 int8 版本。")


if __name__ == "__main__":
    main()