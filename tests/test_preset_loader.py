#! /usr/bin/env python
# @Time    : 2026/9/28
# @Author  : afish
# @File    : test_preset_loader.py
"""Preset 预设加载器（preset_loader.py）单元测试。

每个用例使用独立的临时 presets 目录（monkeypatch PRESET_DIR），
避免依赖/污染仓库 presets/ 目录。
"""

import sys

import pytest

from wordformat.hooks import HookRegistry


@pytest.fixture(autouse=True)
def _clean_loaded_preset_modules():
    """清理缓存到 sys.modules 的预设模块，避免跨用例串状态。"""
    for key in list(sys.modules):
        if key.startswith("wordformat_preset_"):
            del sys.modules[key]
    yield
    for key in list(sys.modules):
        if key.startswith("wordformat_preset_"):
            del sys.modules[key]


@pytest.fixture
def presets_dir(tmp_path, monkeypatch):
    """临时 presets 目录。"""
    import wordformat.preset_loader as pl

    monkeypatch.setattr(pl, "PRESET_DIR", tmp_path)
    return tmp_path


def _write_preset(directory, name, manifest=None, register_src=None):
    """写入一个预设文件（默认含 register 函数）。"""
    lines = []
    if manifest is not None:
        lines.append(f"PRESET_MANIFEST = {manifest!r}")
    if register_src is None:
        lines.append(
            "def register(hooks):\n"
            "    def cb(data):\n"
            "        data['text'] = data.get('text', '') + '|preset'\n"
            "    hooks.register('on_comment', cb)"
        )
    else:
        lines.append(register_src)
    (directory / f"{name}.py").write_text("\n".join(lines), encoding="utf-8")


class TestPresetDirListing:
    def test_preset_dir_creates_missing_directory(self, tmp_path, monkeypatch):
        import wordformat.preset_loader as pl

        target = tmp_path / "no_such_dir"
        monkeypatch.setattr(pl, "PRESET_DIR", target)
        assert pl.preset_dir() == target
        assert target.exists()

    def test_list_presets_filters_template_and_non_py(self, presets_dir):
        from wordformat.preset_loader import list_presets

        _write_preset(presets_dir, "tsinghua")
        _write_preset(presets_dir, "_template")
        (presets_dir / "notes.txt").write_text("x", encoding="utf-8")
        (presets_dir / "__init__.py").write_text("", encoding="utf-8")
        assert list_presets() == ["tsinghua"]

    def test_list_presets_empty_when_dir_missing(self, monkeypatch, tmp_path):
        import wordformat.preset_loader as pl

        monkeypatch.setattr(pl, "PRESET_DIR", tmp_path / "absent")
        assert pl.list_presets() == []


class TestFindPreset:
    def test_find_existing(self, presets_dir):
        from wordformat.preset_loader import find_preset

        _write_preset(presets_dir, "tsinghua")
        path = find_preset("tsinghua")
        assert path.name == "tsinghua.py"

    @pytest.mark.parametrize("bad", ["", "  ", "\x00ab", "a/b", "a\\b"])
    def test_find_rejects_illegal_names(self, presets_dir, bad):
        from wordformat.preset_loader import find_preset

        with pytest.raises(FileNotFoundError):
            find_preset(bad)

    def test_find_missing_lists_available(self, presets_dir):
        from wordformat.preset_loader import find_preset

        _write_preset(presets_dir, "tsinghua")
        with pytest.raises(FileNotFoundError, match="可用: tsinghua"):
            find_preset("unknown")


class TestLoadPreset:
    def test_load_registers_callbacks_and_returns_manifest(self, presets_dir):
        from wordformat.preset_loader import load_preset

        _write_preset(presets_dir, "p1", manifest={"template_name": "甲模板"})
        registry = HookRegistry()
        manifest = load_preset("p1", registry=registry)
        assert manifest == {"template_name": "甲模板"}
        assert registry.is_registered("on_comment")
        assert len(registry.registered_callbacks("on_comment")) == 1

    def test_load_without_manifest_returns_empty(self, presets_dir):
        from wordformat.preset_loader import load_preset

        _write_preset(presets_dir, "p2")
        assert load_preset("p2") == {}

    def test_load_missing_register_raises(self, presets_dir):
        from wordformat.preset_loader import load_preset

        _write_preset(presets_dir, "p3", register_src="X = 1")
        with pytest.raises(FileNotFoundError, match="未定义 register"):
            load_preset("p3")

    def test_load_raises_and_cleans_cache_on_exec_error(self, presets_dir):
        from wordformat.preset_loader import load_preset

        _write_preset(presets_dir, "p4", register_src="raise RuntimeError('bad')")
        with pytest.raises(RuntimeError, match="bad"):
            load_preset("p4")
        # 加载失败的模块已被清理，后续重试不再直接命中缓存
        assert not [k for k in sys.modules if k == "wordformat_preset_p4"]

    def test_module_cached_between_loads(self, presets_dir, monkeypatch):
        import wordformat.preset_loader as pl

        _write_preset(presets_dir, "p5", manifest={"a": 1})
        first = pl.load_preset("p5")
        second = pl.load_preset("p5")
        assert first == second == {"a": 1}
        assert "wordformat_preset_p5" in sys.modules

    def test_load_same_preset_twice_registers_once(self, presets_dir):
        from wordformat.preset_loader import load_preset

        _write_preset(presets_dir, "p6")
        registry = HookRegistry()
        load_preset("p6", registry=registry)
        load_preset("p6", registry=registry)
        assert len(registry.registered_callbacks("on_comment")) == 1

    def test_clear_allows_preset_reload(self, presets_dir):
        from wordformat.preset_loader import clear_presets, load_preset

        _write_preset(presets_dir, "p7")
        registry = HookRegistry()
        load_preset("p7", registry=registry)
        clear_presets(registry)
        load_preset("p7", registry=registry)
        assert len(registry.registered_callbacks("on_comment")) == 1

    def test_load_presets_returns_manifests_in_order(self, presets_dir):
        from wordformat.preset_loader import load_presets

        _write_preset(presets_dir, "p1", manifest={"a": 1})
        _write_preset(presets_dir, "p2", manifest={"b": 2})
        manifests = load_presets(["p1", "p2"])
        assert manifests == [{"a": 1}, {"b": 2}]

    def test_load_presets_none_or_empty(self, presets_dir):
        from wordformat.preset_loader import load_presets

        assert load_presets() == []
        assert load_presets([]) == []

    def test_clear_presets(self, presets_dir):
        from wordformat.preset_loader import clear_presets, load_preset

        _write_preset(presets_dir, "p1")
        registry = HookRegistry()
        load_preset("p1", registry=registry)
        assert registry.is_registered("on_comment")
        clear_presets(registry)
        assert registry.registered_events() == []
