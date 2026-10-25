"""CLI 其余子命令的参数派发：gj/cf/af 预设、tree/config/md/jd/list-presets。"""

import json
from unittest import mock

import pytest
import yaml

from wordformat.cli import main


def _run(argv):
    with mock.patch("sys.argv", argv):
        main()


@pytest.fixture
def files(tmp_path):
    docx = tmp_path / "a.docx"
    docx.write_bytes(b"docx")
    cfg = tmp_path / "c.yaml"
    cfg.write_text("a: 1", encoding="utf-8")
    js = tmp_path / "a.json"
    js.write_text("{}", encoding="utf-8")
    return docx, cfg, js


class TestGjMode:
    def test_gj_writes_auto_named_json(self, tmp_path, files):
        docx, cfg, _js = files
        out_dir = tmp_path / "gj-out"
        with mock.patch(
            "wordformat.cli.set_tag_main", return_value={"body": []}
        ) as tag:
            _run(
                [
                    "wordf",
                    "gj",
                    "-d",
                    str(docx),
                    "-c",
                    str(cfg),
                    "-o",
                    str(out_dir),
                ]
            )
        tag.assert_called_once_with(docx_path=str(docx), configpath=str(cfg))
        produced = list(out_dir.glob("a_*.json"))
        assert len(produced) == 1
        assert json.loads(produced[0].read_text(encoding="utf-8")) == {"body": []}


class TestCheckAndApplyOptions:
    def test_cf_passes_presets(self, tmp_path, files):
        docx, cfg, js = files
        auto = mock.Mock()
        with mock.patch("wordformat.cli.auto_format_thesis_document", auto):
            _run(
                [
                    "wordf",
                    "cf",
                    "-d",
                    str(docx),
                    "-c",
                    str(cfg),
                    "-f",
                    str(js),
                    "-o",
                    str(tmp_path / "out"),
                    "--preset",
                    "sci",
                    "--preset",
                    "tight",
                ]
            )
        kwargs = auto.call_args.kwargs
        assert kwargs["check"] is True
        assert kwargs["presets"] == ["sci", "tight"]
        assert kwargs["docxpath"] == str(docx)
        assert kwargs["configpath"] == str(cfg)
        assert kwargs["jsonpath"] == str(js)
        assert kwargs["savepath"] == str(tmp_path / "out")

    def test_af_passes_sync_and_presets(self, tmp_path, files):
        docx, cfg, js = files
        auto = mock.Mock()
        with mock.patch("wordformat.cli.auto_format_thesis_document", auto):
            _run(
                [
                    "wordf",
                    "af",
                    "-d",
                    str(docx),
                    "-c",
                    str(cfg),
                    "-f",
                    str(js),
                    "-o",
                    str(tmp_path / "out"),
                    "--preset",
                    "sci",
                    "--sync",
                ]
            )
        kwargs = auto.call_args.kwargs
        assert kwargs["check"] is False
        assert kwargs["sync"] is True
        assert kwargs["presets"] == ["sci"]
        assert kwargs["docxpath"] == str(docx)
        assert kwargs["configpath"] == str(cfg)
        assert kwargs["jsonpath"] == str(js)
        assert kwargs["savepath"] == str(tmp_path / "out")


class TestTreeMode:
    def test_filter_confidence_and_index(self, files):
        _docx, _cfg, js = files
        tree = mock.Mock()
        with mock.patch("wordformat.cli.print_tree", tree):
            _run(
                [
                    "wordf",
                    "tree",
                    "-f",
                    str(js),
                    "--confidence",
                    "--index",
                    "--filter",
                    "heading_level_1, body_text",
                ]
            )
        tree.assert_called_once_with(
            node_or_jsonpath=str(js),
            show_confidence=True,
            show_index=True,
            filter_categories=["heading_level_1", "body_text"],
        )

    def test_no_filter(self, files):
        _docx, _cfg, js = files
        tree = mock.Mock()
        with mock.patch("wordformat.cli.print_tree", tree):
            _run(["wordf", "tree", "-f", str(js)])
        assert tree.call_args.kwargs["filter_categories"] is None


class TestConfigMode:
    def test_export_template(self, tmp_path):
        out = tmp_path / "template.yaml"
        with mock.patch(
            "wordformat.structure.registry.export_defaults",
            return_value={"body": {"text": {"font_size": "小四"}}},
        ):
            _run(["wordf", "config", "-o", str(out)])
        assert yaml.safe_load(out.read_text(encoding="utf-8")) == {
            "body": {"text": {"font_size": "小四"}}
        }

    def test_show_config_without_output(self):
        show = mock.Mock()
        with mock.patch("wordformat.cli._show_config", show):
            _run(["wordf", "config"])
        show.assert_called_once_with()


class TestMarkdownAndJsonModes:
    def test_md_passes_paths_and_presets(self, tmp_path, files):
        _docx, cfg, _js = files
        md_file = tmp_path / "note.md"
        md_file.write_text("# t", encoding="utf-8")
        out_dir = tmp_path / "out"
        md = mock.Mock()
        with mock.patch("wordformat.cli.md_to_docx", md):
            _run(
                [
                    "wordf",
                    "md",
                    "-d",
                    str(md_file),
                    "-c",
                    str(cfg),
                    "-o",
                    str(out_dir),
                    "--preset",
                    "sci",
                ]
            )
        md.assert_called_once_with(
            md_path=str(md_file),
            config_path=str(cfg),
            save_dir=str(out_dir),
            presets=["sci"],
        )

    def test_jd_passes_paths_and_presets(self, tmp_path, files):
        _docx, cfg, js = files
        out_dir = tmp_path / "out"
        jd = mock.Mock(return_value=str(out_dir / "new.docx"))
        with mock.patch("wordformat.cli.json_to_docx", jd):
            _run(
                [
                    "wordf",
                    "jd",
                    "-f",
                    str(js),
                    "-c",
                    str(cfg),
                    "-o",
                    str(out_dir),
                    "--preset",
                    "sci",
                ]
            )
        jd.assert_called_once_with(
            json_path=str(js),
            config_path=str(cfg),
            save_dir=str(out_dir),
            presets=["sci"],
        )


class TestListPresetsMode:
    def test_lists_presets_verbosely(self, tmp_path):
        with (
            mock.patch(
                "wordformat.preset_loader.list_presets",
                return_value=["sci", "tight"],
            ) as listed,
            mock.patch(
                "wordformat.preset_loader.preset_dir", return_value=tmp_path
            ) as directory,
            mock.patch("wordformat.cli.console.print") as printed,
        ):
            _run(["wordf", "list-presets", "-v"])
        listed.assert_called_once_with()
        assert directory.call_count == 2
        printed_lines = [c.args[0] for c in printed.call_args_list]
        assert f"sci <- {tmp_path / 'sci.py'}" in printed_lines
        assert f"tight <- {tmp_path / 'tight.py'}" in printed_lines

    def test_empty_presets(self):
        with (
            mock.patch("wordformat.preset_loader.list_presets", return_value=[]),
            mock.patch("wordformat.preset_loader.preset_dir", return_value=None),
            mock.patch("wordformat.cli.logger.info") as info,
        ):
            _run(["wordf", "list-presets"])
        assert any("暂无可用预设" in str(c.args) for c in info.call_args_list)


def test_show_config_prints_exported_yaml(capsys):
    with mock.patch(
        "wordformat.structure.registry.export_defaults",
        return_value={"body": {"text": {"font_size": "小四"}}},
    ):
        _run(["wordf", "config"])
    assert "font_size" in capsys.readouterr().out


def test_list_presets_non_verbose():
    with (
        mock.patch("wordformat.preset_loader.list_presets", return_value=["sci"]),
        mock.patch("wordformat.preset_loader.preset_dir", return_value=None),
        mock.patch("wordformat.cli.console.print") as printed,
        mock.patch("wordformat.cli.logger.info") as info,
    ):
        _run(["wordf", "list-presets"])
    printed_lines = [c.args[0] for c in printed.call_args_list]
    assert printed_lines[-1] == "sci"
    assert any("加 -v" in str(c.args) for c in info.call_args_list)


def test_startapi_explicit_port(monkeypatch):
    import sys
    import types

    fake_uvicorn = types.ModuleType("uvicorn")
    fake_uvicorn.run = mock.Mock()
    monkeypatch.setitem(sys.modules, "uvicorn", fake_uvicorn)
    monkeypatch.setattr(sys, "argv", ["wordf", "startapi", "-p", "9000"])
    main()
    assert fake_uvicorn.run.call_args.kwargs["port"] == 9000
