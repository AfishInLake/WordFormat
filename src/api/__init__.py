#! /usr/bin/env python
# @Time    : 2026/2/5 21:41
# @Author  : afish
# @File    : __init__.py
import os
from pathlib import Path
from typing import Optional
from urllib.parse import quote

from fastapi import Body, FastAPI, File, HTTPException, UploadFile
from loguru import logger
from pydantic import BaseModel
from starlette.responses import FileResponse

# 复用原有项目的核心函数和校验工具
from src.set_style import auto_format_thesis_document
from src.set_tag import set_tag_main
from src.settings import SERVER_HOST, WORK_DIR

# ---------------------- 初始化FastAPI应用 ----------------------
app = FastAPI(
    title="学位论文格式自动校验工具-WebAPI",
    description="基于FastAPI实现，支持generate-json/check-format/apply-format三种模式，兼容原有YAML配置",
    version="1.0.0",
    docs_url="/docs",  # Swagger UI接口文档地址（推荐）
    redoc_url="/redoc",  # ReDoc接口文档地址（备选）
)

# ---------------------- 全局配置 ----------------------
# 临时文件目录（存储上传的docx/配置文件、生成的json），自动创建
TEMP_DIR = WORK_DIR / "temp"
TEMP_DIR.mkdir(parents=True, exist_ok=True)
logger.info(f"临时文件目录：{TEMP_DIR}")
# 输出文件目录（存储校验/格式化后的docx）
OUTPUT_DIR = WORK_DIR / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
logger.info(f"输出文件目录：{OUTPUT_DIR}")  # 修复原日志笔误


# ---------------------- 数据模型（接口参数校验） ----------------------
class OperationResult(BaseModel):
    """统一接口返回格式"""

    code: int  # 200=成功，400=参数错误，500=执行失败
    msg: str  # 结果描述
    data: Optional[dict] = None  # 附加数据（如文件路径、日志信息）


# ---------------------- 核心工具函数 ----------------------
def save_upload_file(upload_file: UploadFile, save_dir: Path) -> str:
    """
    保存上传的文件到指定目录，仅返回【文件绝对路径】（重名自动加_1/_2后缀，避免覆盖）
    :param upload_file: FastAPI上传的文件对象
    :param save_dir: 保存目录
    :return: 文件绝对路径
    """
    try:
        original_filename = upload_file.filename
        file_suffix = os.path.splitext(original_filename)[-1]
        original_name_no_suffix = os.path.splitext(original_filename)[0]

        # 重名处理：自动追加数字后缀，防止文件覆盖
        save_path = os.path.join(save_dir, original_filename)
        counter = 1
        while os.path.exists(save_path):
            save_path = os.path.join(
                save_dir, f"{original_name_no_suffix}_{counter}{file_suffix}"
            )
            counter += 1

        # 保存文件流
        with open(save_path, "wb") as f:
            f.write(upload_file.file.read())
        abs_path = os.path.abspath(save_path)
        logger.info(f"上传文件已保存：{abs_path}")
        return abs_path
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"文件保存失败：{str(e)}") from e


# ---------------------- 核心API接口 ----------------------
@app.post(
    "/generate-json", response_model=OperationResult, summary="生成文档结构JSON文件"
)
async def api_generate_json(
    docx_file: UploadFile = File(..., description="待处理的Word文档（.docx格式）"),  # noqa B008
    config_file: UploadFile = File(..., description="格式配置YAML文件（必填）"),  # noqa B008
):
    """
    对应原命令行generate-json模式：仅生成JSON，不执行校验/格式化
    - 上传docx和yaml配置文件，服务端自动生成JSON并返回数据
    """
    try:
        # 保存上传文件（仅返回路径，无需提取原名称）
        docx_path = save_upload_file(docx_file, TEMP_DIR)
        config_path = save_upload_file(config_file, TEMP_DIR)

        # 基于上传文件路径生成同名JSON（保持原有逻辑）
        json_filename = f"{os.path.splitext(os.path.basename(docx_path))[0]}.json"
        json_path = os.path.join(TEMP_DIR, json_filename)

        # 执行核心逻辑生成JSON
        json_data = set_tag_main(
            docx_path=docx_path, json_save_path=json_path, configpath=config_path
        )

        return OperationResult(
            code=200,
            msg="JSON文件生成成功",
            data={
                "json_data": json_data,
                "json_filename": json_filename,
                "tips": "可使用该JSON数据调用check-format/apply-format接口",
            },
        )
    except Exception as e:
        logger.error(f"生成JSON失败：{str(e)}")
        raise HTTPException(status_code=500, detail=f"生成JSON失败：{str(e)}") from e


@app.post(
    "/check-format",
    response_model=OperationResult,
    summary="仅执行格式校验（需先生成JSON）",
)
async def api_check_format(
    docx_file: UploadFile = File(..., description="待校验的Word文档（.docx格式）"),  # noqa B008
    config_file: UploadFile = File(..., description="格式配置YAML文件（必填）"),  # noqa B008
    json_data: str = Body(..., description="从/generate-json获取的文档结构JSON数据"),
):
    """
    对应原命令行check-format模式：仅执行格式校验，生成【原文件名+--标注版.docx】
    - 基于函数返回的真实路径拼接下载链接，保证路径100%匹配
    """
    try:
        # 1. 保存上传文件（仅返回实际路径）
        docx_path = save_upload_file(docx_file, TEMP_DIR)
        config_path = save_upload_file(config_file, TEMP_DIR)

        # 2. 执行校验逻辑，获取【函数返回的实际保存文件路径】（核心！）
        actual_save_path = auto_format_thesis_document(
            jsonpath=json_data,
            docxpath=docx_path,
            configpath=config_path,
            savepath=OUTPUT_DIR,
            check=True,  # 仅校验模式
        )

        # 3. 从真实路径中提取最终文件名（如：1 (1)_2--标注版.docx）
        final_filename = os.path.basename(actual_save_path)
        # 4. 拼接正确的下载链接（仅编码文件名，无多余路径）
        encoded_filename = quote(final_filename)
        download_url = f"{SERVER_HOST}/download/{encoded_filename}"

        # 5. 记录真实的保存路径日志（与函数内部日志一致）
        logger.info(f"格式校验完成，结果保存：{actual_save_path}")

        # 6. 返回结果（含真实文件名和下载链接）
        return OperationResult(
            code=200,
            msg="格式校验执行成功",
            data={
                "original_docx": docx_file.filename,  # 用户上传的原文件名
                "final_filename": final_filename,  # 实际保存的最终文件名
                "download_url": download_url,  # 可直接点击的下载链接
                "tips": "文件已保存至服务端output目录，点击链接直接下载",
            },
        )
    except Exception as e:
        logger.error(f"格式校验失败：{str(e)}")
        raise HTTPException(status_code=500, detail=f"格式校验失败：{str(e)}") from e


@app.post(
    "/apply-format", response_model=OperationResult, summary="执行格式应用/自动格式化"
)
async def api_apply_format(
    docx_file: UploadFile = File(..., description="待格式化的Word文档（.docx格式）"),  # noqa B008
    config_file: UploadFile = File(..., description="格式配置YAML文件（必填）"),  # noqa B008
    json_data: str = Body(..., description="从/generate-json获取的文档结构JSON数据"),
):
    """
    对应原命令行apply-format模式：自动应用格式，生成【原文件名+--修改版.docx】
    - 基于函数返回的真实路径拼接下载链接，彻底解决404问题
    """
    try:
        # 1. 保存上传文件（仅返回实际路径）
        docx_path = save_upload_file(docx_file, TEMP_DIR)
        config_path = save_upload_file(config_file, TEMP_DIR)

        # 2. 执行格式化逻辑，获取【函数返回的实际保存文件路径】（核心！）
        actual_save_path = auto_format_thesis_document(
            jsonpath=json_data,
            docxpath=docx_path,
            configpath=config_path,
            savepath=OUTPUT_DIR,
            check=False,  # 格式化模式
        )

        # 3. 从真实路径提取最终文件名（如：1 (1)_2--修改版.docx）
        final_filename = os.path.basename(actual_save_path)
        # 4. 拼接正确下载链接（仅编码文件名，避免路径错误）
        encoded_filename = quote(final_filename)
        download_url = f"{SERVER_HOST}/download/{encoded_filename}"

        # 5. 记录真实保存路径日志（与函数内部日志完全一致）
        logger.info(f"文档格式化完成，结果保存：{actual_save_path}")

        # 6. 返回结果（含原文件名、最终文件名、下载链接）
        return OperationResult(
            code=200,
            msg="文档格式化执行成功",
            data={
                "original_docx": docx_file.filename,  # 用户上传的原文件名
                "final_filename": final_filename,  # 服务端实际保存的文件名
                "download_url": download_url,  # 前端可直接使用的下载链接
                "tips": "文件已保存至服务端output目录，点击链接直接下载",
            },
        )
    except Exception as e:
        logger.error(f"文档格式化失败：{str(e)}")
        raise HTTPException(status_code=500, detail=f"文档格式化失败：{str(e)}") from e


@app.get("/download/{filename}", summary="下载格式化/校验后的Word文档")
def download_file(filename: str):
    """
    下载接口：基于真实文件名匹配，增加多层校验，保证下载稳定
    """
    try:
        # 拼接服务端实际文件路径
        file_path = os.path.join(OUTPUT_DIR, filename)
        # 校验1：文件是否存在
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="文件不存在或已被删除")
        # 校验2：是否为有效文件（非文件夹/链接）
        if not os.path.isfile(file_path):
            raise HTTPException(status_code=400, detail="请求路径不是有效文件")
        # 以附件形式返回，浏览器自动触发下载，指定docx专属MIME类型
        return FileResponse(
            file_path,
            filename=filename,  # 强制指定下载显示的文件名（与实际保存一致）
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except HTTPException as e:
        logger.error(str(e))
    except Exception as e:
        logger.error(f"文件下载失败：{filename} -> {str(e)}")
        raise HTTPException(status_code=500, detail=f"文件下载失败：{str(e)}") from e
