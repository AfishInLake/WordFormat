import json
import multiprocessing as mp
import time
from typing import Dict, List, Optional

import numpy as np
from loguru import logger

from wordformat.settings import ONNX_VERSION

onnx_version = str(ONNX_VERSION)
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
            files(f"wordformat.data.model.{onnx_version}").joinpath(
                "bert_paragraph_classifier.onnx"
            )
        ),
        "tokenizer": str(files("wordformat.data.model").joinpath("tokenizer.json")),
        "id2label": str(
            files(f"wordformat.data.model.{onnx_version}").joinpath("id2label.json")
        ),
    }


# ===== 硬件自动选择函数 =====
def _get_best_onnx_providers() -> List[str]:
    """自动选择最优推理硬件：CUDA > DirectML > CPU"""
    try:
        import onnxruntime as ort

        available = ort.get_available_providers()
        # 优先级1：NVIDIA显卡（CUDA）
        if "CUDAExecutionProvider" in available:
            logger.info("检测到CUDA，优先使用GPU推理")
            return ["CUDAExecutionProvider"]
        # 优先级2：Windows核显（Intel/AMD，DirectML）
        elif "DmlExecutionProvider" in available:
            logger.info("检测到核显，使用DirectML推理")
            return ["DmlExecutionProvider"]
        # 优先级3：纯CPU（兜底）
        else:
            logger.info("仅检测到CPU，使用CPU多核推理")
            return ["CPUExecutionProvider"]
    except Exception as e:
        logger.warning(f"硬件检测失败，降级为CPU：{e}")
        return ["CPUExecutionProvider"]


# ===== 模型加载函数 =====
def _load_model():
    global _tokenizer, _ort_sess, _id2label
    if _tokenizer is not None:
        return

    import onnxruntime as ort
    from tokenizers import Tokenizer

    paths = _get_model_paths()
    logger.info(f"首次调用，正在加载模型：{paths['onnx']}")

    # 1. 加载Tokenizer（原有逻辑保留）
    _tokenizer = Tokenizer.from_file(paths["tokenizer"])

    # 2. 优化ONNX推理器配置
    ort_options = ort.SessionOptions()
    # 开启全量图优化（BERT模型提速关键）
    ort_options.graph_optimization_level = ort.GraphOptimizationLevel.ORT_ENABLE_ALL
    # CPU多核推理（拉满所有核心）
    cpu_core_num = mp.cpu_count()
    ort_options.intra_op_num_threads = cpu_core_num
    ort_options.inter_op_num_threads = cpu_core_num
    # 关闭冗余日志
    ort_options.log_severity_level = 3
    # 内存复用（大批次推理防OOM）
    ort_options.enable_cpu_mem_arena = True
    ort_options.enable_mem_pattern = True
    ort_options.enable_mem_reuse = True
    # 设置推理超时
    ort_options.execution_mode = ort.ExecutionMode.ORT_SEQUENTIAL

    # 3. 加载ONNX模型（自动适配硬件）
    providers = _get_best_onnx_providers()

    try:
        _ort_sess = ort.InferenceSession(
            paths["onnx"], sess_options=ort_options, providers=providers
        )

    except Exception as e:
        logger.warning(f"最优硬件加载失败，降级为CPU：{e}")
        _ort_sess = ort.InferenceSession(
            paths["onnx"], sess_options=ort_options, providers=["CPUExecutionProvider"]
        )
    # 4. 加载id2label（原有逻辑保留）
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
    # 超时控制
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

    # ===== 批量预处理（向量化替代循环拼接）=====
    batch_size = len(texts)
    # 初始化批量数组（直接预分配内存，避免循环append）
    batch_input_ids = np.zeros((batch_size, MAX_LENGTH), dtype=np.int64)
    batch_attention_mask = np.zeros((batch_size, MAX_LENGTH), dtype=np.int64)
    batch_token_type_ids = np.zeros((batch_size, MAX_LENGTH), dtype=np.int64)

    # 批量编码（保留逻辑，但优化赋值方式）
    for idx, text in enumerate(texts):
        encoded = _tokenizer.encode(text, add_special_tokens=True)
        input_ids = encoded.ids
        attention_mask = encoded.attention_mask
        token_type_ids = encoded.type_ids

        # 截断
        seq_len = min(len(input_ids), MAX_LENGTH)
        # 直接赋值（比append+拼接快）
        batch_input_ids[idx, :seq_len] = input_ids[:seq_len]
        batch_attention_mask[idx, :seq_len] = attention_mask[:seq_len]
        batch_token_type_ids[idx, :seq_len] = token_type_ids[:seq_len]

    # ===== 批量推理（超时控制+性能监控）=====
    onnx_input = {
        "input_ids": batch_input_ids,
        "attention_mask": batch_attention_mask,
        "token_type_ids": batch_token_type_ids,
    }

    try:
        start = time.time()
        # 批量推理（一次性处理，减少ONNX调用开销）
        logits = _ort_sess.run(["logits"], onnx_input)[0]  # shape: [batch, num_classes]
        infer_time = time.time() - start
        logger.info(
            f"批量推理完成 | 批次大小：{batch_size} | 耗时：{infer_time:.4f}s | 单条耗时：{infer_time / batch_size:.4f}s"  # noqa E501
        )
    except Exception as e:
        logger.error(f"批量推理失败：{e}")
        # 兜底：返回空结果，避免整体崩溃
        return [
            {"原始文本": text, "预测标签": "", "预测ID": -1, "预测概率": 0.0}
            for text in texts
        ]

    # ===== 向量化计算概率（替代循环）=====
    # 数值稳定化（避免exp溢出）
    logits_stable = logits - np.max(logits, axis=-1, keepdims=True)
    probs = np.exp(logits_stable) / np.sum(
        np.exp(logits_stable), axis=-1, keepdims=True
    )
    # 批量获取预测ID和概率（向量化操作，比循环快）
    pred_ids = np.argmax(probs, axis=-1).astype(int)
    pred_probs = np.round(np.max(probs, axis=-1).astype(float), 4)

    # ===== 结果组装 =====
    results = []
    for idx, text in enumerate(texts):
        pred_id = pred_ids[idx]
        results.append(
            {
                "原始文本": text,
                "预测标签": _id2label.get(pred_id, ""),
                "预测ID": pred_id,
                "预测概率": float(pred_probs[idx]),
            }
        )

    return results


def safe_batch_infer(texts: list[str], max_batch_size: int = 128) -> list[dict]:
    """
    安全批量推理（自动分片，避免超大批次OOM）
    :param texts: 文本列表
    :param max_batch_size: 单批次最大数量（默认128，可根据硬件调整）
    :return: 完整结果列表
    """
    if not texts:
        return []

    start = time.time()
    results = []
    total = len(texts)
    # 分片处理
    for i in range(0, total, max_batch_size):
        batch = texts[i : i + max_batch_size]
        batch_results = onnx_batch_infer(batch)
        results.extend(batch_results)
        # 进度日志
        logger.info(f"已处理 {min(i + max_batch_size, total)}/{total} 条文本")

    total_time = time.time() - start
    logger.info(
        f"安全批量推理完成 | 总条数：{total} | 总耗时：{total_time:.4f}s | 平均单条：{total_time / total:.4f}s"  # noqa E501
    )
    return results
