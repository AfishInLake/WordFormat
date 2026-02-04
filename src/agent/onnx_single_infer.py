import json

import numpy as np
import onnxruntime as ort
from loguru import logger
from tokenizers import Tokenizer

from src.settings import BERT_MODEL

# ===== 配置 =====
ONNX_MODEL_PATH = BERT_MODEL + "bert_paragraph_classifier.onnx"
TOKENIZER_JSON = "model/" + "tokenizer.json"
ID2LABEL_PATH = BERT_MODEL + "id2label.json"
MAX_LENGTH = 128

# ===== 加载 =====
logger.info(f"加载模型中：{TOKENIZER_JSON}")
tokenizer = Tokenizer.from_file(TOKENIZER_JSON)
ort_sess = ort.InferenceSession(ONNX_MODEL_PATH, providers=["CPUExecutionProvider"])
with open(ID2LABEL_PATH, encoding="utf-8") as f:
    id2label = {int(k): v for k, v in json.load(f).items()}


def onnx_single_infer(text: str) -> dict:
    # 使用官方 tokenizer 编码
    encoded = tokenizer.encode(text, add_special_tokens=True)
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

    logits = ort_sess.run(["logits"], onnx_input)[0]
    logits = logits[0]
    logits -= np.max(logits)
    probs = np.exp(logits) / np.sum(np.exp(logits))
    pred_id = int(np.argmax(probs))
    pred_prob = round(float(probs[pred_id]), 4)

    return {"label": id2label[pred_id], "score": pred_prob}


if __name__ == "__main__":
    text = "摘要：你好啊。"
    a = onnx_single_infer(text)
