"""API 新增入口：md / jd / af --sync / --preset / tree。"""

import io
import json
from unittest import mock

import pytest
from fastapi.testclient import TestClient

from wordformat.tree import render_tree


@pytest.fixture
def api_client(tmp_path):
    temp_dir = tmp_path / "temp"
    output_dir = tmp_path / "output"
    temp_dir.mkdir()
    output_dir.mkdir()
    with (
        mock.patch("wordformat.api.BASE_DIR", tmp_path),
        mock.patch("wordformat.api.TEMP_DIR", temp_dir),
        mock.patch("wordformat.api.OUTPUT_DIR", output_dir),
    ):
        from wordformat.api import app

        yield TestClient(app), temp_dir, output_dir


class TestMarkdownToDocx:
    def test_success(self, api_client):
        client, _temp, output_dir = api_client
        saved = str(output_dir / "note--生成版.docx")
        with mock.patch("wordformat.api.md_to_docx", return_value=saved) as md_mock:
            response = client.post(
                "/markdown-to-docx",
                files={"md_file": ("note.md", io.BytesIO(b"# t"), "text/markdown")},
                data={"presets": "sci, tight"},
            )

        assert response.status_code == 200
        data = response.json()["data"]
        assert data["final_filename"] == "note--生成版.docx"
        assert data["download_url"].startswith("/download/")
        kwargs = md_mock.call_args.kwargs
        assert kwargs["md_path"].endswith("note.md")
        assert kwargs["config_path"] is None
        assert kwargs["save_dir"] == str(output_dir)
        assert kwargs["presets"] == ["sci", "tight"]

    def test_config_uploaded(self, api_client):
        client, temp_dir, output_dir = api_client
        saved = str(output_dir / "n--生成版.docx")
        with mock.patch("wordformat.api.md_to_docx", return_value=saved) as md_mock:
            response = client.post(
                "/markdown-to-docx",
                files={
                    "md_file": ("n.md", io.BytesIO(b"text"), "text/markdown"),
                    "config_file": ("c.yaml", io.BytesIO(b"a: 1"), "text/yaml"),
                },
            )
        assert response.status_code == 200
        kwargs = md_mock.call_args.kwargs
        assert kwargs["config_path"].endswith("c.yaml")
        assert (temp_dir / "c.yaml").exists()
        assert kwargs["presets"] is None

    def test_wrong_suffix_returns_400(self, api_client):
        client, _temp, _output = api_client
        with mock.patch("wordformat.api.md_to_docx") as md_mock:
            response = client.post(
                "/markdown-to-docx",
                files={"md_file": ("note.txt", io.BytesIO(b"x"), "text/plain")},
            )
        assert response.json()["code"] == 400
        assert "仅支持" in response.json()["msg"]
        md_mock.assert_not_called()

    def test_exception_returns_500(self, api_client):
        client, _temp, _output = api_client
        with mock.patch("wordformat.api.md_to_docx", side_effect=RuntimeError("boom")):
            response = client.post(
                "/markdown-to-docx",
                files={"md_file": ("note.md", io.BytesIO(b"x"), "text/markdown")},
            )
        assert response.status_code == 500
        assert "boom" in response.json()["detail"]


class TestJsonToDocx:
    def test_success(self, api_client):
        client, _temp, output_dir = api_client
        saved = str(output_dir / "model--生成版.docx")
        with mock.patch("wordformat.api.json_to_docx", return_value=saved) as jd_mock:
            response = client.post(
                "/json-to-docx",
                files={
                    "json_file": (
                        "model.json",
                        io.BytesIO(b"[]"),
                        "application/json",
                    )
                },
                data={"presets": "p1"},
            )
        assert response.status_code == 200
        assert response.json()["data"]["final_filename"] == "model--生成版.docx"
        kwargs = jd_mock.call_args.kwargs
        assert kwargs["json_path"].endswith("model.json")
        assert kwargs["presets"] == ["p1"]

    def test_wrong_suffix_returns_400(self, api_client):
        client, _temp, _output = api_client
        response = client.post(
            "/json-to-docx",
            files={"json_file": ("model.yaml", io.BytesIO(b"a: 1"), "text/yaml")},
        )
        assert response.json()["code"] == 400


class TestApplySyncAndPresets:
    def test_sync_passed_through(self, api_client):
        client, _temp, output_dir = api_client
        saved = str(output_dir / "t--修改版.docx")
        with mock.patch(
            "wordformat.api.auto_format_thesis_document", return_value=saved
        ) as af_mock:
            response = client.post(
                "/apply-format",
                files={
                    "docx_file": (
                        "t.docx",
                        io.BytesIO(b"x"),
                        "application/octet-stream",
                    )
                },
                data={"json_data": "[]", "sync": "true", "presets": "a,b"},
            )
        assert response.status_code == 200
        kwargs = af_mock.call_args.kwargs
        assert kwargs["check"] is False
        assert kwargs["sync"] is True
        assert kwargs["presets"] == ["a", "b"]

    def test_defaults_omit_sync_and_presets(self, api_client):
        client, _temp, output_dir = api_client
        saved = str(output_dir / "t--修改版.docx")
        with mock.patch(
            "wordformat.api.auto_format_thesis_document", return_value=saved
        ) as af_mock:
            response = client.post(
                "/apply-format",
                files={
                    "docx_file": (
                        "t.docx",
                        io.BytesIO(b"x"),
                        "application/octet-stream",
                    )
                },
                data={"json_data": "[]"},
            )
        assert response.status_code == 200
        kwargs = af_mock.call_args.kwargs
        assert "sync" not in kwargs
        assert "presets" not in kwargs

    def test_check_format_presets(self, api_client):
        client, _temp, output_dir = api_client
        saved = str(output_dir / "t--标注版.docx")
        with mock.patch(
            "wordformat.api.auto_format_thesis_document", return_value=saved
        ) as cf_mock:
            response = client.post(
                "/check-format",
                files={
                    "docx_file": (
                        "t.docx",
                        io.BytesIO(b"x"),
                        "application/octet-stream",
                    )
                },
                data={"json_data": "[]", "presets": "p1,p2"},
            )
        assert response.status_code == 200
        data = response.json()["data"]
        assert data["final_filename"] == "t--标注版.docx"
        assert data["download_url"].endswith("t--%E6%A0%87%E6%B3%A8%E7%89%88.docx")
        kwargs = cf_mock.call_args.kwargs
        assert kwargs["check"] is True
        assert kwargs["presets"] == ["p1", "p2"]


class TestTreeEndpoint:
    def test_success(self, api_client):
        client, _temp, _output = api_client
        with mock.patch(
            "wordformat.api.render_tree", return_value="root\n  body_text 正文"
        ) as tree_mock:
            response = client.post(
                "/tree",
                files={
                    "json_file": ("doc.json", io.BytesIO(b"[]"), "application/json")
                },
                data={
                    "show_confidence": "true",
                    "show_index": "true",
                    "filter_categories": "body_text, heading_level_1",
                },
            )
        assert response.status_code == 200
        assert response.json()["data"]["tree"] == "root\n  body_text 正文"
        kwargs = tree_mock.call_args.kwargs
        assert kwargs["show_confidence"] is True
        assert kwargs["show_index"] is True
        assert kwargs["filter_categories"] == ["body_text", "heading_level_1"]

    def test_wrong_suffix_returns_400(self, api_client):
        client, _temp, _output = api_client
        response = client.post(
            "/tree",
            files={
                "json_file": ("doc.docx", io.BytesIO(b"x"), "application/octet-stream")
            },
        )
        assert response.json()["code"] == 400


class TestRenderTree:
    def test_plain_text_output(self, tmp_path):
        path = tmp_path / "doc.json"
        path.write_text(
            json.dumps(
                [
                    {
                        "category": "heading_level_1",
                        "paragraph": "第一章 绪论",
                        "score": 0.91,
                        "fingerprint": "a",
                    },
                    {
                        "category": "body_text",
                        "paragraph": "正文内容",
                        "score": 0.88,
                        "fingerprint": "b",
                    },
                ],
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
        text = render_tree(str(path), show_confidence=True)
        assert "第一章 绪论" in text
        assert "正文内容" in text
        assert "91%" in text
        assert "\x1b[" not in text


@pytest.fixture
def configs_client(tmp_path):
    configs_dir = tmp_path / "configs"
    configs_dir.mkdir()
    with mock.patch("wordformat.api.CONFIGS_DIR", configs_dir):
        from wordformat.api import app

        yield TestClient(app), configs_dir


class TestConfigEndpoints:
    def test_list_configs(self, configs_client):
        client, configs_dir = configs_client
        (configs_dir / "b.yml").write_text("b: 1", encoding="utf-8")
        (configs_dir / "a.yaml").write_text("a: 1", encoding="utf-8")
        (configs_dir / "ignore.txt").write_text("x", encoding="utf-8")

        response = client.get("/configs")

        assert response.status_code == 200
        assert response.json()["data"] == ["a.yaml", "b.yml"]

    def test_read_config(self, configs_client):
        client, configs_dir = configs_client
        (configs_dir / "a.yaml").write_text("a: 1\n", encoding="utf-8")

        response = client.get("/configs/a.yaml")

        assert response.status_code == 200
        assert response.json()["data"] == "a: 1\n"

    def test_read_missing_config_returns_404(self, configs_client):
        client, _configs_dir = configs_client
        assert client.get("/configs/none.yaml").status_code == 404

    def test_save_config_appends_yaml_suffix(self, configs_client):
        client, configs_dir = configs_client

        response = client.post(
            "/configs/save", json={"filename": "new", "content": "a: 1"}
        )

        assert response.status_code == 200
        assert "configs/new.yaml" in response.json()["msg"]
        assert (configs_dir / "new.yaml").read_text(encoding="utf-8") == "a: 1"

    def test_save_config_keeps_existing_suffix(self, configs_client):
        client, configs_dir = configs_client

        response = client.post(
            "/configs/save", json={"filename": "x.yml", "content": "b: 2"}
        )

        assert response.status_code == 200
        assert (configs_dir / "x.yml").read_text(encoding="utf-8") == "b: 2"

    def test_save_config_rejects_path_traversal(self, configs_client):
        client, configs_dir = configs_client

        response = client.post(
            "/configs/save",
            json={"filename": "../escape.yaml", "content": "x: 1"},
        )

        assert response.status_code == 400
        assert not (configs_dir.parent / "escape.yaml").exists()


class TestCommandErrors:
    def test_json_to_docx_failure_returns_500(self, api_client):
        client, _temp, _output = api_client
        with mock.patch(
            "wordformat.api.json_to_docx", side_effect=RuntimeError("boom")
        ):
            response = client.post(
                "/json-to-docx",
                files={"json_file": ("a.json", io.BytesIO(b"[]"), "application/json")},
            )
        assert response.status_code == 500

    def test_tree_failure_returns_500(self, api_client):
        client, _temp, _output = api_client
        with mock.patch("wordformat.api.render_tree", side_effect=RuntimeError("boom")):
            response = client.post(
                "/tree",
                files={"json_file": ("a.json", io.BytesIO(b"[]"), "application/json")},
            )
        assert response.status_code == 500

    def test_download_directory_returns_400(self, api_client):
        client, _temp, output_dir = api_client
        (output_dir / "folder.docx").mkdir()

        assert client.get("/download/folder.docx").status_code == 400

    def test_download_rejects_path_traversal(self, api_client):
        # 客户端/路由层会规范化 ../，直接调用接口才能命中文件名守卫
        import wordformat.api as api_module
        from fastapi import HTTPException

        _client, _temp, _output = api_client
        with pytest.raises(HTTPException) as exc_info:
            api_module.download_file("../secret.docx")
        assert exc_info.value.status_code == 400

    def test_check_format_failure_returns_500(self, api_client):
        client, _temp, _output = api_client
        with mock.patch(
            "wordformat.api.auto_format_thesis_document",
            side_effect=RuntimeError("boom"),
        ):
            response = client.post(
                "/check-format",
                files={
                    "docx_file": (
                        "a.docx",
                        io.BytesIO(b"x"),
                        "application/octet-stream",
                    )
                },
                data={"json_data": "[]"},
            )
        assert response.status_code == 500

    def test_apply_format_failure_returns_500(self, api_client):
        client, _temp, _output = api_client
        with mock.patch(
            "wordformat.api.auto_format_thesis_document",
            side_effect=RuntimeError("boom"),
        ):
            response = client.post(
                "/apply-format",
                files={
                    "docx_file": (
                        "a.docx",
                        io.BytesIO(b"x"),
                        "application/octet-stream",
                    )
                },
                data={"json_data": "[]"},
            )
        assert response.status_code == 500
