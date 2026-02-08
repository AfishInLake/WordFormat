import json
from typing import Dict, Optional

import numpy as np
from loguru import logger

# ===== 全局变量（初始为 None）=====
_tokenizer: Optional["Tokenizer"] = None  # noqa F821
_ort_sess: Optional["ort.InferenceSession"] = None  # noqa F821
_id2label: Optional[Dict[int, str]] = None
MAX_LENGTH = 128


# ===== 路径配置（仅定义，不加载）=====
def _get_model_paths():
    from importlib.resources import files

    return {
        "onnx": str(
            files("wordformat.data.model.20260204").joinpath(
                "bert_paragraph_classifier.onnx"
            )
        ),
        "tokenizer": str(files("wordformat.data.model").joinpath("tokenizer.json")),
        "id2label": str(
            files("wordformat.data.model.20260204").joinpath("id2label.json")
        ),
    }


# ===== 模型加载函数 =====
def _load_model():
    global _tokenizer, _ort_sess, _id2label
    if _tokenizer is not None:
        return

    import onnxruntime as ort
    from tokenizers import Tokenizer

    paths = _get_model_paths()
    logger.info(f"首次调用，正在加载模型：{paths['onnx']}")

    _tokenizer = Tokenizer.from_file(paths["tokenizer"])
    _ort_sess = ort.InferenceSession(paths["onnx"], providers=["CPUExecutionProvider"])

    with open(paths["id2label"], encoding="utf-8") as f:
        _id2label = {int(k): v for k, v in json.load(f).items()}


def onnx_single_infer(text: str) -> dict:
    if _tokenizer is None:
        _load_model()

    # 确保类型安全（静态检查友好）
    assert _tokenizer is not None
    assert _ort_sess is not None
    assert _id2label is not None

    # 使用官方 tokenizer 编码
    encoded = _tokenizer.encode(text, add_special_tokens=True)
    input_ids = encoded.ids
    attention_mask = encoded.attention_mask
    token_type_ids = encoded.type_ids

    # 截断 & 补全
    if len(input_ids) > MAX_LENGTH:
        input_ids = input_ids[:MAX_LENGTH]
        attention_mask = attention_mask[:MAX_LENGTH]
        token_type_ids = token_type_ids[:MAX_LENGTH]
    else:
        pad_len = MAX_LENGTH - len(input_ids)
        input_ids += [0] * pad_len
        attention_mask += [0] * pad_len
        token_type_ids += [0] * pad_len

    onnx_input = {
        "input_ids": np.array([input_ids], dtype=np.int64),
        "attention_mask": np.array([attention_mask], dtype=np.int64),
        "token_type_ids": np.array([token_type_ids], dtype=np.int64),
    }

    logits = _ort_sess.run(["logits"], onnx_input)[0]
    logits = logits[0]
    logits -= np.max(logits)
    probs = np.exp(logits) / np.sum(np.exp(logits))
    pred_id = int(np.argmax(probs))
    pred_prob = round(float(probs[pred_id]), 4)

    return {"label": _id2label[pred_id], "score": pred_prob}


def onnx_batch_infer(texts: list[str]) -> list[dict]:
    """
    ONNX模型批量文本推理
    :param texts: 待分类的论文段落文本列表
    :return: 每条文本的预测结果列表，每个元素为字典（同单条推理格式）
    """
    global _tokenizer, _ort_sess, _id2label

    if not texts:
        return []

    if _tokenizer is None:
        _load_model()

    # 确保类型安全（静态检查友好）
    assert _tokenizer is not None
    assert _ort_sess is not None
    assert _id2label is not None

    # 步骤1：批量分词编码（padding到相同长度）
    batch_input_ids = []
    batch_attention_mask = []
    batch_token_type_ids = []

    # 手动批量编码（因为 tokenizers.Tokenizer 不支持直接传 list + padding）
    for text in texts:
        encoded = _tokenizer.encode(text, add_special_tokens=True)
        input_ids = encoded.ids
        attention_mask = encoded.attention_mask
        token_type_ids = encoded.type_ids

        # 截断
        if len(input_ids) > MAX_LENGTH:
            input_ids = input_ids[:MAX_LENGTH]
            attention_mask = attention_mask[:MAX_LENGTH]
            token_type_ids = token_type_ids[:MAX_LENGTH]
        else:
            # 补零到 MAX_LENGTH
            pad_len = MAX_LENGTH - len(input_ids)
            input_ids += [0] * pad_len
            attention_mask += [0] * pad_len
            token_type_ids += [0] * pad_len

        batch_input_ids.append(input_ids)
        batch_attention_mask.append(attention_mask)
        batch_token_type_ids.append(token_type_ids)

    # 步骤2：构造ONNX输入
    onnx_input = {
        "input_ids": np.array(batch_input_ids, dtype=np.int64),
        "attention_mask": np.array(batch_attention_mask, dtype=np.int64),
        "token_type_ids": np.array(batch_token_type_ids, dtype=np.int64),
    }

    # 步骤3：ONNX批量推理
    logits = _ort_sess.run(output_names=["logits"], input_feed=onnx_input)[
        0
    ]  # shape: [batch, 15]
    logits_stable = logits - np.max(logits, axis=-1, keepdims=True)
    probs = np.exp(logits_stable) / np.sum(
        np.exp(logits_stable), axis=-1, keepdims=True
    )

    # 步骤4：逐样本解析结果
    results = []
    for i, text in enumerate(texts):
        pred_id = int(np.argmax(logits[i]))
        pred_label = _id2label[pred_id]
        prob_val = round(float(probs[i, pred_id]), 4)

        results.append(
            {
                "原始文本": text,
                "预测标签": pred_label,
                "预测ID": pred_id,
                "预测概率": prob_val,
            }
        )

    return results


def safe_batch_infer(texts: list[str], max_batch_size: int = 128) -> list[dict]:
    results = []
    for i in range(0, len(texts), max_batch_size):
        batch = texts[i : i + max_batch_size]
        results.extend(onnx_batch_infer(batch))
    return results
