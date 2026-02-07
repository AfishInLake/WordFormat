from typing import Dict, Optional

import numpy as np
from loguru import logger

# ===== 全局变量（初始为 None）=====
_tokenizer: Optional["Tokenizer"] = None  # noqa F821
_ort_sess: Optional["ort.InferenceSession"] = None  # noqa F821
_id2label: Optional[Dict[int, str]] = None


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
    import json

    import onnxruntime as ort
    from tokenizers import Tokenizer

    paths = _get_model_paths()
    logger.info(f"首次调用，正在加载模型：{paths['tokenizer']}")

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
    MAX_LENGTH = 128

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


if __name__ == "__main__":
    text = "摘要：你好啊。"
    a = onnx_single_infer(text)
