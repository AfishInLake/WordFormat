"""ONNX 段落分类推理（单模型：thesis_paragraph_classifier）。

小模型为 uer/chinese_roberta_L-4_H-256 微调产物，推理侧只加载这一个模型
（模型名/标签表/词表由 model_info.json 描述，无旧模型兼容逻辑）。

上下文建模：
    训练时每条样本带 "[PREV={上一段标签}] {文本}" 前缀，推理侧同样按
    "上一段预测标签" 拼接（首段为 "[PREV=None] ..."），跨批次由调用方或
    safe_batch_infer 维护状态链。
"""

import json
import time
from typing import Dict, List, Optional, Tuple

import numpy as np
from loguru import logger

from wordformat.settings import (
    BATCH_SIZE,
    ONNX_INTER_OP_THREADS,
    ONNX_INTRA_OP_THREADS,
)

# ===== 全局缓存（懒加载，初始为 None）=====
_tokenizer: Optional["Tokenizer"] = None  # noqa: F821
_ort_sess: Optional["ort.InferenceSession"] = None  # noqa: F821
_id2label: Optional[Dict[int, str]] = None
_model_info: Optional[dict] = None
MAX_LENGTH = 128


def _get_model_dir():
    from importlib.resources import files

    return files("wordformat.data.model")


def _get_best_onnx_providers() -> List[str]:
    """自动选择最优推理硬件：CUDA > DirectML > CPU"""
    try:
        import onnxruntime as ort

        available = ort.get_available_providers()
        if "CUDAExecutionProvider" in available:
            logger.info("检测到CUDA，优先使用GPU推理")
            return ["CUDAExecutionProvider"]
        elif "DmlExecutionProvider" in available:
            logger.info("检测到核显，使用DirectML推理")
            return ["DmlExecutionProvider"]
        else:
            logger.info("仅检测到CPU，使用CPU多核推理")
            return ["CPUExecutionProvider"]
    except Exception as e:
        logger.warning(f"硬件检测失败，降级为CPU：{e}")
        return ["CPUExecutionProvider"]


def _load_model() -> None:
    """首次调用时加载模型/词表/标签表（后续调用直接复用全局缓存）。"""
    global _tokenizer, _ort_sess, _id2label, _model_info, MAX_LENGTH
    if _ort_sess is not None:
        return

    import onnxruntime as ort
    from tokenizers import Tokenizer

    model_dir = _get_model_dir()
    with open(model_dir.joinpath("model_info.json"), encoding="utf-8") as f:
        _model_info = json.load(f)

    onnx_name = _model_info.get("onnx") or _model_info["fp32_onnx"]
    onnx_path = str(model_dir.joinpath(onnx_name))
    logger.info(
        f"首次调用，正在加载模型：{onnx_path} "
        f"(base={_model_info.get('base_model')}, quantized={_model_info.get('quantized')})"
    )

    _tokenizer = Tokenizer.from_file(str(model_dir.joinpath("tokenizer.json")))
    MAX_LENGTH = int(_model_info.get("max_length", 128))

    ort_options = ort.SessionOptions()
    ort_options.graph_optimization_level = ort.GraphOptimizationLevel.ORT_ENABLE_ALL
    # 线程数走 settings 配置，适配 2 核小服务器
    ort_options.intra_op_num_threads = ONNX_INTRA_OP_THREADS
    ort_options.inter_op_num_threads = ONNX_INTER_OP_THREADS
    ort_options.log_severity_level = 3
    ort_options.enable_cpu_mem_arena = True
    ort_options.enable_mem_pattern = True
    ort_options.enable_mem_reuse = True
    ort_options.execution_mode = ort.ExecutionMode.ORT_SEQUENTIAL

    providers = _get_best_onnx_providers()
    try:
        _ort_sess = ort.InferenceSession(
            onnx_path, sess_options=ort_options, providers=providers
        )
    except Exception as e:
        logger.warning(f"最优硬件加载失败，降级为CPU：{e}")
        _ort_sess = ort.InferenceSession(
            onnx_path, sess_options=ort_options, providers=["CPUExecutionProvider"]
        )

    with open(model_dir.joinpath("id2label.json"), encoding="utf-8") as f:
        _id2label = {int(k): v for k, v in json.load(f).items()}


def _prompt(prev_label: Optional[str], text: str) -> str:
    """拼接 PREV 上下文前缀，与训练侧 PREV_PREFIX 格式严格一致。"""
    return f"[PREV={prev_label}] {text}"


def _encode_batch(texts: List[str]) -> Tuple[np.ndarray, np.ndarray]:
    """批量编码并截断/补齐到 MAX_LENGTH，返回 (input_ids, attention_mask)。"""
    batch_size = len(texts)
    batch_input_ids = np.zeros((batch_size, MAX_LENGTH), dtype=np.int64)
    batch_attention_mask = np.zeros((batch_size, MAX_LENGTH), dtype=np.int64)
    for idx, text in enumerate(texts):
        encoded = _tokenizer.encode(text, add_special_tokens=True)  # type: ignore[misc]
        seq_len = min(len(encoded.ids), MAX_LENGTH)
        batch_input_ids[idx, :seq_len] = encoded.ids[:seq_len]
        batch_attention_mask[idx, :seq_len] = encoded.attention_mask[:seq_len]
    return batch_input_ids, batch_attention_mask


def onnx_single_infer(text: str, prev_label: Optional[str] = None) -> dict:
    """单条推理。prev_label 为上一段预测标签（跨调用维护可串成状态链）。"""
    if _ort_sess is None:
        _load_model()
    assert _ort_sess is not None
    assert _id2label is not None

    input_ids, attention_mask = _encode_batch([_prompt(prev_label, text)])
    onnx_input = {"input_ids": input_ids, "attention_mask": attention_mask}

    try:
        start = time.time()
        logits = _ort_sess.run(["logits"], onnx_input)[0]
        logger.debug(f"单条推理耗时：{time.time() - start:.4f}s")
    except Exception as e:
        logger.error(f"单条推理失败：{e}")
        return {"label": "", "score": 0.0}

    logits = logits[0]
    logits -= np.max(logits)
    probs = np.exp(logits) / np.sum(np.exp(logits))
    pred_id = int(np.argmax(probs))

    return {"label": _id2label[pred_id], "score": round(float(probs[pred_id]), 4)}


def onnx_batch_infer(
    texts: List[str], prev_label: Optional[str] = None
) -> Tuple[List[dict], Optional[str]]:
    """批量推理。整个 batch 共享同一 PREV 上下文（跨批链由调用方推进），
    返回 (结果列表, 本批末段的预测标签)。失败时降级为带状态链的逐条推理。
    """
    global _ort_sess, _id2label
    if not texts:
        return [], prev_label
    if _ort_sess is None:
        _load_model()
    assert _ort_sess is not None
    assert _id2label is not None

    batch_size = len(texts)
    prompts = [_prompt(prev_label, t) for t in texts]
    input_ids, attention_mask = _encode_batch(prompts)
    onnx_input = {"input_ids": input_ids, "attention_mask": attention_mask}

    try:
        start = time.time()
        logits = _ort_sess.run(["logits"], onnx_input)[0]  # shape: [batch, num_classes]
        infer_time = time.time() - start
        logger.info(
            f"批量推理完成 | 批次大小：{batch_size} | 耗时：{infer_time:.4f}s | "
            f"单条耗时：{infer_time / batch_size:.4f}s"
        )
    except Exception as e:
        logger.error(f"批量推理失败，逐条降级：{e}")
        results: List[dict] = []
        last = prev_label
        for text in texts:
            r = onnx_single_infer(text, last)
            results.append(
                {"text": text, "label": r["label"], "pred_id": -1, "score": r["score"]}
            )
            if r["label"]:
                last = r["label"]
        return results, last

    # 向量化 softmax（数值稳定），向量化取 top1
    logits_stable = logits - np.max(logits, axis=-1, keepdims=True)
    probs = np.exp(logits_stable) / np.sum(
        np.exp(logits_stable), axis=-1, keepdims=True
    )
    pred_ids = np.argmax(probs, axis=-1).astype(int)
    pred_probs = np.round(np.max(probs, axis=-1).astype(float), 4)

    results = []
    last = prev_label
    for idx, text in enumerate(texts):
        pred_id = int(pred_ids[idx])
        label = _id2label.get(pred_id, "")
        if label:
            last = label
        results.append(
            {
                "text": text,
                "label": label,
                "pred_id": pred_id,
                "score": float(pred_probs[idx]),
            }
        )

    return results, last


def safe_batch_infer(
    texts: List[str],
    max_batch_size: Optional[int] = None,
    initial_prev: Optional[str] = None,
) -> List[dict]:
    """安全批量推理：自动分片，跨分片维护 PREV 状态链。

    :param texts: 文本列表
    :param max_batch_size: 单批次最大数量（默认取 settings.BATCH_SIZE）
    :param initial_prev: 首段之前的标签（None 表示文档开头）
    :return: 完整结果列表
    """
    if not texts:
        return []

    if max_batch_size is None:
        max_batch_size = BATCH_SIZE

    start = time.time()
    results: List[dict] = []
    total = len(texts)
    prev = initial_prev
    for i in range(0, total, max_batch_size):
        batch = texts[i : i + max_batch_size]
        batch_results, prev = onnx_batch_infer(batch, prev)
        results.extend(batch_results)
        logger.info(f"已处理 {min(i + max_batch_size, total)}/{total} 条文本")

    total_time = time.time() - start
    logger.info(
        f"安全批量推理完成 | 总条数：{total} | 总耗时：{total_time:.4f}s | "
        f"平均单条：{total_time / total:.4f}s"
    )
    return results
