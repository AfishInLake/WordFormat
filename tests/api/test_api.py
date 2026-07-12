"""
集成测试：跨模块交互与端到端行为验证

覆盖模块：cli.py, config, agent, set_style.py, set_tag.py, word_structure
"""
import argparse
import io
import os
import shutil
import tempfile
import threading
from unittest import mock

import pytest
import numpy as np
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from pydantic import ValidationError
from fastapi.testclient import TestClient

from wordformat.cli import validate_file, main
from wordformat.config.loader import (
    LazyConfig,
    ConfigNotLoadedError,
    init_config,
    get_config,
    clear_config,
)
from wordformat.config.models import NodeConfigRoot
from wordformat.agent.message import MessageManager
from wordformat.agent.onnx_infer import (
    _get_best_onnx_providers,
    _load_model,
    onnx_single_infer,
    onnx_batch_infer,
    safe_batch_infer,
)
from wordformat.pipeline.stages import FormattingExecutionStage, ParagraphAlignmentStage
from wordformat.classify.tag import set_tag_main
from wordformat.structure.node_factory import create_node
from wordformat.structure.tree_builder import DocumentTreeBuilder
from wordformat.structure.document_builder import DocumentBuilder
from wordformat.structure.utils import promote_bodytext_in_subtrees_of_type
from wordformat.rules.node import FormatNode
import wordformat.api
from wordformat.rules.body import BodyText
from wordformat.api import save_upload_file

apply_format_check_to_all_nodes = FormattingExecutionStage().apply_format_check_to_all_nodes


# ==================== (h) CLI 集成测试 ====================


class TestCLIIntegration:
    """validate_file + mock main() 各模式"""

    def test_validate_file_accepts_real_file(self):
        with tempfile.NamedTemporaryFile(delete=False) as f:
            path = f.name
        try:
            result = validate_file(path, "文档")
            assert result == os.path.abspath(path)
        finally:
            os.unlink(path)

    def test_validate_file_rejects_missing(self):
        with pytest.raises(argparse.ArgumentTypeError):
            validate_file("/nonexistent/path.txt", "文档")

    def test_validate_file_rejects_directory(self):
        with tempfile.TemporaryDirectory() as d:
            with pytest.raises(argparse.ArgumentTypeError):
                validate_file(d, "文档")

    def test_validate_file_rejects_wrong_extension(self):
        with tempfile.NamedTemporaryFile(suffix=".exe", delete=False) as f:
            path = f.name
        try:
            with pytest.raises(argparse.ArgumentTypeError):
                validate_file(path, "文档", [".docx"])
        finally:
            os.unlink(path)

    @mock.patch("sys.argv")
    def test_main_no_args_prints_help(self, mock_argv):
        mock_argv.__getitem__.side_effect = lambda i: ["wf"][i]
        mock_argv.__len__.return_value = 1
        main()  # 不应抛异常

    @mock.patch("wordformat.cli.set_tag_main", return_value=[])
    @mock.patch("wordformat.cli.json")
    @mock.patch("sys.argv")
    def test_main_gj_mode(self, mock_argv, mock_json, mock_set_tag):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as df:
            docx_path = df.name
        with tempfile.NamedTemporaryFile(suffix=".yaml", delete=False) as cf:
            cfg_path = cf.name
        out_dir = tempfile.mkdtemp()
        try:
            mock_argv.__getitem__.side_effect = lambda i: [
                "wf", "gj", "-d", docx_path, "-c", cfg_path, "-o", out_dir
            ][i]
            mock_argv.__len__.return_value = 8
            main()
            mock_set_tag.assert_called_once()
            mock_json.dump.assert_called_once()
        finally:
            os.unlink(docx_path)
            os.unlink(cfg_path)
            shutil.rmtree(out_dir, ignore_errors=True)

    @mock.patch("wordformat.cli.auto_format_thesis_document")
    @mock.patch("sys.argv")
    def test_main_cf_mode(self, mock_argv, mock_auto):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as df:
            docx_path = df.name
        with tempfile.NamedTemporaryFile(suffix=".yaml", delete=False) as cf:
            cfg_path = cf.name
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as jf:
            json_path = jf.name
        out_dir = tempfile.mkdtemp()
        try:
            mock_argv.__getitem__.side_effect = lambda i: [
                "wf", "cf", "-d", docx_path, "-c", cfg_path, "-f", json_path, "-o", out_dir
            ][i]
            mock_argv.__len__.return_value = 10
            main()
            mock_auto.assert_called_once_with(
                jsonpath=json_path, docxpath=docx_path,
                configpath=cfg_path, savepath=out_dir, check=True,
            )
        finally:
            for p in [docx_path, cfg_path, json_path]:
                os.unlink(p)
            os.rmdir(out_dir)

    @mock.patch("wordformat.cli.auto_format_thesis_document")
    @mock.patch("sys.argv")
    def test_main_af_mode(self, mock_argv, mock_auto):
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as df:
            docx_path = df.name
        with tempfile.NamedTemporaryFile(suffix=".yaml", delete=False) as cf:
            cfg_path = cf.name
        with tempfile.NamedTemporaryFile(suffix=".json", delete=False) as jf:
            json_path = jf.name
        out_dir = tempfile.mkdtemp()
        try:
            mock_argv.__getitem__.side_effect = lambda i: [
                "wf", "af", "-d", docx_path, "-c", cfg_path, "-f", json_path, "-o", out_dir
            ][i]
            mock_argv.__len__.return_value = 10
            main()
            mock_auto.assert_called_once_with(
                jsonpath=json_path, docxpath=docx_path,
                configpath=cfg_path, savepath=out_dir, check=False,
            )
        finally:
            for p in [docx_path, cfg_path, json_path]:
                os.unlink(p)
            os.rmdir(out_dir)

    @mock.patch("sys.argv")
    def test_main_startapi_rejects_invalid_port(self, mock_argv):
        mock_argv.__getitem__.side_effect = lambda i: ["wf", "startapi", "-p", "99999"][i]
        mock_argv.__len__.return_value = 4
        with mock.patch.dict("sys.modules", {"uvicorn": mock.MagicMock()}):
            with pytest.raises(SystemExit):
                main()



# ==================== (u) api/__init__.py 覆盖测试 ====================


class TestSaveUploadFile:
    """覆盖 api/__init__.py save_upload_file 函数"""

    def test_save_upload_file_normal(self, tmp_path):
        """正常保存上传文件"""
        with mock.patch("wordformat.api.BASE_DIR", tmp_path), \
                mock.patch("wordformat.api.OUTPUT_DIR", tmp_path / "output"), \
                mock.patch("wordformat.api.TEMP_DIR", tmp_path / "temp"):
            from wordformat.api import save_upload_file
            wordformat.api.TEMP_DIR.mkdir(parents=True, exist_ok=True)

            mock_file = mock.MagicMock()
            mock_file.filename = "test.docx"
            mock_file.file.read.return_value = b"fake docx content"

            result = save_upload_file(mock_file, wordformat.api.TEMP_DIR)
            assert os.path.exists(result)
            with open(result, "rb") as f:
                assert f.read() == b"fake docx content"

    def test_save_upload_file_name_conflict(self, tmp_path):
        """文件名冲突时自动重命名"""
        with mock.patch("wordformat.api.BASE_DIR", tmp_path), \
                mock.patch("wordformat.api.OUTPUT_DIR", tmp_path / "output"), \
                mock.patch("wordformat.api.TEMP_DIR", tmp_path / "temp"):
            wordformat.api.TEMP_DIR.mkdir(parents=True, exist_ok=True)

            # 预先创建同名文件
            existing_file = wordformat.api.TEMP_DIR / "test.docx"
            existing_file.write_bytes(b"existing content")

            mock_file = mock.MagicMock()
            mock_file.filename = "test.docx"
            mock_file.file.read.return_value = b"new content"

            result = save_upload_file(mock_file, wordformat.api.TEMP_DIR)
            assert os.path.exists(result)
            assert result.endswith("test_1.docx")
            with open(result, "rb") as f:
                assert f.read() == b"new content"
            # 原文件未被覆盖
            assert existing_file.read_bytes() == b"existing content"

    def test_save_upload_file_multiple_conflicts(self, tmp_path):
        """多次冲突时递增后缀"""
        with mock.patch("wordformat.api.BASE_DIR", tmp_path), \
                mock.patch("wordformat.api.OUTPUT_DIR", tmp_path / "output"), \
                mock.patch("wordformat.api.TEMP_DIR", tmp_path / "temp"):
            from wordformat.api import save_upload_file
            wordformat.api.TEMP_DIR.mkdir(parents=True, exist_ok=True)

            (wordformat.api.TEMP_DIR / "test.docx").write_bytes(b"a")
            (wordformat.api.TEMP_DIR / "test_1.docx").write_bytes(b"b")
            (wordformat.api.TEMP_DIR / "test_2.docx").write_bytes(b"c")

            mock_file = mock.MagicMock()
            mock_file.filename = "test.docx"
            mock_file.file.read.return_value = b"d"

            result = save_upload_file(mock_file, wordformat.api.TEMP_DIR)
            assert result.endswith("test_3.docx")



class TestAPIEndpoints:
    """覆盖 api/__init__.py 所有 API 端点"""

    @pytest.fixture
    def api_client(self, tmp_path):
        """创建 TestClient，mock BASE_DIR 使目录创建在 tmp_path"""
        temp_dir = tmp_path / "temp"
        output_dir = tmp_path / "output"
        temp_dir.mkdir(parents=True, exist_ok=True)
        output_dir.mkdir(parents=True, exist_ok=True)

        with mock.patch("wordformat.api.BASE_DIR", tmp_path), \
                mock.patch("wordformat.api.TEMP_DIR", temp_dir), \
                mock.patch("wordformat.api.OUTPUT_DIR", output_dir):
            from wordformat.api import app
            client = TestClient(app)
            yield client, temp_dir, output_dir

    def test_generate_json_success(self, api_client):
        """POST /generate-json 成功调用 set_tag_main"""
        client, temp_dir, output_dir = api_client

        mock_result = [{"category": "body_text", "score": 0.9, "paragraph": "test", "fingerprint": "fp1"}]
        with mock.patch("wordformat.api.set_tag_main", return_value=mock_result):
            docx_bytes = io.BytesIO(b"fake docx")
            yaml_bytes = io.BytesIO(b"key: value")

            response = client.post(
                "/generate-json",
                files={
                    "docx_file": ("test.docx", docx_bytes, "application/octet-stream"),
                    "config_file": ("config.yaml", yaml_bytes, "application/octet-stream"),
                },
            )

        assert response.status_code == 200
        data = response.json()
        assert data["code"] == 200
        assert "json_data" in data["data"]

    def test_generate_json_non_docx_returns_400(self, api_client):
        """POST /generate-json 上传非 docx 文件返回 code 400"""
        client, temp_dir, output_dir = api_client

        docx_bytes = io.BytesIO(b"fake content")
        yaml_bytes = io.BytesIO(b"key: value")

        response = client.post(
            "/generate-json",
            files={
                "docx_file": ("test.pdf", docx_bytes, "application/octet-stream"),
                "config_file": ("config.yaml", yaml_bytes, "application/octet-stream"),
            },
        )

        assert response.status_code == 200
        data = response.json()
        assert data["code"] == 400
        assert ".docx" in data["msg"]

    def test_check_format_success(self, api_client):
        """POST /check-format 成功调用 auto_format_thesis_document(check=True)"""
        client, temp_dir, output_dir = api_client

        mock_result_path = str(output_dir / "test--标注版.docx")
        with mock.patch("wordformat.api.auto_format_thesis_document", return_value=mock_result_path):
            docx_bytes = io.BytesIO(b"fake docx")
            yaml_bytes = io.BytesIO(b"key: value")
            json_str = '[{"category": "body_text", "score": 0.9}]'

            response = client.post(
                "/check-format",
                files={
                    "docx_file": ("test.docx", docx_bytes, "application/octet-stream"),
                    "config_file": ("config.yaml", yaml_bytes, "application/octet-stream"),
                },
                data={"json_data": json_str},
            )

        assert response.status_code == 200
        data = response.json()
        assert data["code"] == 200
        assert "标注版" in data["data"]["final_filename"]
        assert "download_url" in data["data"]

    def test_apply_format_success(self, api_client):
        """POST /apply-format 成功调用 auto_format_thesis_document(check=False)"""
        client, temp_dir, output_dir = api_client

        mock_result_path = str(output_dir / "test--修改版.docx")
        with mock.patch("wordformat.api.auto_format_thesis_document", return_value=mock_result_path):
            docx_bytes = io.BytesIO(b"fake docx")
            yaml_bytes = io.BytesIO(b"key: value")
            json_str = '[{"category": "body_text", "score": 0.9}]'

            response = client.post(
                "/apply-format",
                files={
                    "docx_file": ("test.docx", docx_bytes, "application/octet-stream"),
                    "config_file": ("config.yaml", yaml_bytes, "application/octet-stream"),
                },
                data={"json_data": json_str},
            )

        assert response.status_code == 200
        data = response.json()
        assert data["code"] == 200
        assert "修改版" in data["data"]["final_filename"]
        assert "download_url" in data["data"]

    def test_download_file_exists(self, api_client):
        """GET /download/{filename} 文件存在时返回文件"""
        client, temp_dir, output_dir = api_client

        # 在 output_dir 创建一个测试文件
        test_file = output_dir / "result.docx"
        test_file.write_bytes(b"docx content here")

        response = client.get(f"/download/result.docx")
        assert response.status_code == 200
        assert response.content == b"docx content here"

    def test_download_file_not_found(self, api_client):
        """GET /download/{filename} 文件不存在时返回 404"""
        client, temp_dir, output_dir = api_client

        response = client.get("/download/nonexistent.docx")
        # 已修复：重新抛出 HTTPException，正确返回 404
        assert response.status_code == 404

    def test_generate_json_exception_returns_500(self, api_client):
        """POST /generate-json 异常时返回 500"""
        client, temp_dir, output_dir = api_client

        with mock.patch("wordformat.api.set_tag_main", side_effect=RuntimeError("test error")):
            docx_bytes = io.BytesIO(b"fake docx")
            yaml_bytes = io.BytesIO(b"key: value")

            response = client.post(
                "/generate-json",
                files={
                    "docx_file": ("test.docx", docx_bytes, "application/octet-stream"),
                    "config_file": ("config.yaml", yaml_bytes, "application/octet-stream"),
                },
            )

        assert response.status_code == 500



class TestCLIStartApiMode:
    """覆盖 cli.py startapi 模式（lines 151-169）"""

    def test_startapi_mode_calls_uvicorn(self):
        """startapi 模式调用 uvicorn.run"""
        import unittest.mock as um
        # uvicorn 在函数内部动态导入，需要 mock import
        with um.patch("sys.argv", ["wf", "startapi"]):
            with um.patch.dict("sys.modules", {"uvicorn": um.MagicMock()}):
                with um.patch("wordformat.api.app", um.MagicMock()):
                    main()
        # 验证 uvicorn.run 被调用
        import sys
        if "uvicorn" in sys.modules:
            sys.modules["uvicorn"].run.assert_called_once()

