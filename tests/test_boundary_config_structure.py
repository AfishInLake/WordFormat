"""边界输入测试：配置合并、注册表、节点工厂与文档树构建。"""

import json

import pytest

from wordformat.config.dotdict import (
    BASE_FORMAT,
    DotDict,
    deep_merge,
    first_option,
    flatten_style_config,
)
from wordformat.config.loader import LazyConfig, clear_config, get_config, load_config
from wordformat.config.models import (
    NodeConfigRoot,
    NumberingConfig,
    NumberingLevelConfig,
)
from wordformat.rules import BodyText
from wordformat.rules.node import FormatNode
from wordformat.structure.document_builder import DocumentBuilder
from wordformat.structure.node_factory import create_node
from wordformat.structure.registry import (
    Registry,
    _deep_set,
    default_registry,
    get_registry,
)
from wordformat.structure.tree_builder import DocumentTreeBuilder


def _write(tmp_path, text, name="config.yaml"):
    path = tmp_path / name
    path.write_text(text, encoding="utf-8")
    return str(path)


# ---------------------------------------------------------------------------
# DotDict / deep_merge / flatten_style_config
# ---------------------------------------------------------------------------


class TestDotDictBoundary:
    def test_missing_key_returns_none_for_attribute(self):
        cfg = DotDict()

        assert cfg.missing is None
        with pytest.raises(KeyError):
            _ = cfg["missing"]

    def test_hasattr_is_true_for_any_missing_key(self):
        # __getattr__ 把缺失键映射为 None，hasattr 因此永远为真
        assert hasattr(DotDict(), "never_defined") is True

    def test_nested_dict_returns_new_dotdict_each_time(self):
        cfg = DotDict({"a": {"b": 1}})

        assert cfg.a.b == 1
        assert cfg.a is not cfg.a

    def test_mutating_nested_view_does_not_change_underlying_dict(self):
        # DotDict(val) 会复制一层，嵌套视图上的写入不会回写原字典
        cfg = DotDict({"a": {}})
        view = cfg.a
        view.b = 2

        assert view.b == 2
        assert cfg["a"] == {}

    def test_set_and_delete_attribute(self):
        cfg = DotDict()
        cfg.answer = 42
        assert cfg["answer"] == 42

        del cfg.answer
        assert "answer" not in cfg

    def test_delete_missing_attribute_raises(self):
        with pytest.raises(AttributeError):
            del DotDict().missing


class TestDeepMergeBoundary:
    def test_nested_dicts_merge(self):
        base = {"a": {"b": 1, "c": 2}}
        override = {"a": {"c": 3, "d": 4}}
        result = deep_merge(base, override)

        assert result == {"a": {"b": 1, "c": 3, "d": 4}}
        assert base == {"a": {"b": 1, "c": 2}}  # 不修改入参

    def test_list_is_replaced_not_concatenated(self):
        assert deep_merge({"a": [1, 2]}, {"a": [3]}) == {"a": [3]}

    def test_scalar_replaces_dict_and_dict_replaces_scalar(self):
        assert deep_merge({"a": {"b": 1}}, {"a": 2}) == {"a": 2}
        assert deep_merge({"a": 1}, {"a": {"b": 2}}) == {"a": {"b": 2}}

    def test_empty_inputs(self):
        assert deep_merge({}, {}) == {}
        assert deep_merge({"a": 1}, {}) == {"a": 1}
        assert deep_merge({}, {"a": 1}) == {"a": 1}


class TestFirstOptionBoundary:
    def test_list_takes_first_item(self):
        assert first_option(["四号", "小四"]) == "四号"

    def test_single_value_passthrough(self):
        assert first_option("小四") == "小四"
        assert first_option(None) is None

    def test_empty_list_becomes_none(self):
        assert first_option([]) is None

    def test_tuple_is_not_treated_as_options(self):
        # tuple 可能是 RGB 等真实值（font_color），不能取首项
        assert first_option((0, 0, 0)) == (0, 0, 0)


class TestFlattenStyleConfigBoundary:
    @pytest.mark.parametrize("value", [None, "文本", 0, [1, 2], ()])
    def test_non_dict_passthrough(self, value):
        assert flatten_style_config(value) == value

    def test_groups_are_merged_and_top_level_wins(self):
        raw = {
            "paragraph": {"alignment": "居中", "font_size": "小四"},
            "font": {"font_size": "五号", "bold": True},
            "font_size": "三号",
            "nested": {"font": {"italic": True}},
        }

        assert flatten_style_config(raw) == {
            "alignment": "居中",
            "font_size": "三号",
            "bold": True,
            "nested": {"italic": True},
        }

    def test_base_format_flattens_to_flat_fields(self):
        flat = flatten_style_config(BASE_FORMAT)

        assert flat["font_size"] == "小四"
        assert flat["line_spacingrule"] == "1.5倍行距"
        assert "paragraph" not in flat
        assert "font" not in flat


# ---------------------------------------------------------------------------
# NodeConfigRoot / collect_style_configs / NumberingConfig
# ---------------------------------------------------------------------------


class TestNodeConfigRootBoundary:
    def test_missing_and_nested_access(self):
        cfg = NodeConfigRoot(a=1, section={"font_size": "小四"})

        assert cfg.a == 1
        assert cfg.missing is None
        assert isinstance(cfg.section, NodeConfigRoot)
        assert cfg.section.font_size == "小四"

    def test_model_dump_returns_plain_dict(self):
        cfg = NodeConfigRoot(a=1)

        dumped = cfg.model_dump()
        assert type(dumped) is dict
        assert dumped == {"a": 1}

    def test_set_attribute_writes_key(self):
        cfg = NodeConfigRoot()
        cfg.name = "模板"

        assert cfg == {"name": "模板"}

    def test_collect_style_configs_empty(self):
        assert NodeConfigRoot().collect_style_configs() == {}

    def test_same_style_keeps_only_identical_fields(self):
        cfg = NodeConfigRoot(
            a={
                "builtin_style_name": "正文",
                "font_size": "小四",
                "bold": False,
            },
            b={
                "builtin_style_name": "正文",
                "font_size": "小四",
                "bold": True,
            },
        )

        assert cfg.collect_style_configs() == {
            "Normal": {"builtin_style_name": "正文", "font_size": "小四"}
        }

    def test_global_format_and_table_object_are_excluded(self):
        cfg = NodeConfigRoot(
            global_format={"builtin_style_name": "正文", "font_size": "小四"},
            tables={
                "object": {"builtin_style_name": "正文", "font_size": "小四"},
            },
        )

        assert cfg.collect_style_configs() == {}

    def test_nested_style_config_is_collected(self):
        cfg = NodeConfigRoot(
            body={"text": {"builtin_style_name": "Normal", "font_size": "小四"}}
        )

        assert cfg.collect_style_configs() == {
            "Normal": {"builtin_style_name": "Normal", "font_size": "小四"}
        }

    def test_unknown_style_name_is_kept_verbatim(self):
        cfg = NodeConfigRoot(
            a={"builtin_style_name": "自定义样式", "font_size": "小四"}
        )

        assert "自定义样式" in cfg.collect_style_configs()

    def test_non_dict_nodes_are_ignored(self):
        cfg = NodeConfigRoot(a="x", b=None, c=[{"builtin_style_name": "正文"}])

        assert cfg.collect_style_configs() == {}


class TestNumberingConfigBoundary:
    def test_level_and_caption_access(self):
        cfg = NumberingConfig(captions={"figure": 1}, level_1={"format": "decimal"})

        assert isinstance(cfg.captions, NumberingConfig)
        assert isinstance(cfg.level_1, NumberingLevelConfig)
        assert cfg.level_1.format == "decimal"
        assert cfg.missing is None

    def test_level_missing_returns_none(self):
        level = NumberingLevelConfig()

        assert level.missing is None


# ---------------------------------------------------------------------------
# 配置加载：YAML 锚点、空文件、嵌套写法
# ---------------------------------------------------------------------------


class TestConfigLoaderBoundary:
    def test_yaml_anchor_merge(self, tmp_path):
        path = _write(
            tmp_path,
            "base: &base\n"
            "  font_size: 小四\n"
            "  bold: false\n"
            "body:\n"
            "  <<: *base\n"
            "  alignment: 两端对齐\n",
        )

        cfg = load_config(path)
        assert cfg.body.font_size == "小四"
        assert cfg.body.bold is False
        assert cfg.body.alignment == "两端对齐"

    def test_get_config_before_load_raises(self):
        clear_config()
        with pytest.raises(RuntimeError):
            get_config()

    def test_nested_paragraph_style_is_rejected(self, tmp_path):
        path = _write(
            tmp_path,
            "body:\n  paragraph:\n    font_size: 小四\n",
        )

        with pytest.raises(ValueError) as exc:
            load_config(path)
        assert "body.paragraph" in str(exc.value)
        assert "平铺" in str(exc.value)

    def test_top_level_nested_style_reports_root(self, tmp_path):
        path = _write(tmp_path, "font:\n  font_size: 小四\n")

        with pytest.raises(ValueError) as exc:
            load_config(path)
        assert "配置根" in str(exc.value)

    def test_yaml_list_root_is_rejected(self, tmp_path):
        path = _write(tmp_path, "- a\n- b\n")

        with pytest.raises(ValueError, match="YAML 映射"):
            load_config(path)

    def test_empty_yaml_raises_readable_error(self, tmp_path):
        path = _write(tmp_path, "")

        with pytest.raises(ValueError, match="YAML 映射"):
            load_config(path)

    def test_lazy_config_reloads_and_clears(self, tmp_path):
        first = _write(tmp_path, "template_name: 甲\n", name="a.yaml")
        second = _write(tmp_path, "template_name: 乙\n", name="b.yaml")
        lazy = LazyConfig(first)

        assert lazy.load().template_name == "甲"
        lazy.init(second)
        assert lazy.load().template_name == "乙"
        assert lazy.get().template_name == "乙"

        lazy.clear()
        with pytest.raises(RuntimeError):
            lazy.get()


# ---------------------------------------------------------------------------
# Registry / _deep_set
# ---------------------------------------------------------------------------


class TestRegistryBoundary:
    def test_register_without_level_keeps_default_999(self):
        reg = Registry()

        class Node(FormatNode):
            pass

        reg.register("no_level")(Node)

        assert reg.get("no_level") is Node
        assert "no_level" in reg
        assert reg.level_of("no_level") == 999

    def test_register_duplicate_overrides_class_and_level(self):
        reg = Registry()

        class First(FormatNode):
            pass

        class Second(FormatNode):
            pass

        reg.register("same", level=1)(First)
        reg.register("same", level=2)(Second)

        assert reg.get("same") is Second
        assert reg.level_of("same") == 2
        assert len(dict(reg.items())) == 1

    def test_setdefault_does_not_override(self):
        reg = Registry()

        class Real(FormatNode):
            pass

        class Fallback(FormatNode):
            pass

        reg.register("cat", level=1)(Real)
        reg.setdefault("cat", Fallback, level=2)
        reg.setdefault("new", Fallback, level=3)

        assert reg.get("cat") is Real
        assert reg.level_of("cat") == 1
        assert reg.get("new") is Fallback
        assert reg.level_of("new") == 3

    def test_reverse_category_returns_name_or_none(self):
        reg = Registry()

        class Node(FormatNode):
            pass

        reg.register("cat")(Node)

        assert reg.reverse_category(Node) == "cat"
        assert reg.reverse_category(FormatNode) is None

    def test_unknown_category_get_and_contains(self):
        reg = Registry()

        assert reg.get("missing") is None
        assert reg.get("missing", "fallback") == "fallback"
        assert "missing" not in reg
        assert reg.level_of("missing") == 999

    def test_export_defaults_skips_incomplete_classes(self):
        reg = Registry()

        class WithDefaults(FormatNode):
            NODE_TYPE = "custom.section"
            DEFAULTS = {
                "paragraph": {"alignment": "居中"},
                "font": {"font_size": "小四"},
                "font_size": "五号",
            }

        class NoDefaults(FormatNode):
            NODE_TYPE = "custom.empty"
            DEFAULTS = {}

        class NoNodeType(FormatNode):
            NODE_TYPE = ""
            DEFAULTS = {"font_size": "小四"}

        reg.register("with_defaults")(WithDefaults)
        reg.register("no_defaults")(NoDefaults)
        reg.register("no_node_type")(NoNodeType)

        exported = reg.export_defaults()

        assert exported["template_name"] == "未知模板"
        assert exported["numbering"] == {"enabled": False}
        assert exported["custom"]["section"] == {
            "alignment": "居中",
            "font_size": "五号",
        }
        assert "custom" in exported and "empty" not in exported["custom"]

    def test_deep_set_creates_and_merges_paths(self):
        result = {}
        _deep_set(result, "a.b.c", {"x": 1})
        _deep_set(result, "a.b.c", {"y": 2})

        assert result == {"a": {"b": {"c": {"x": 1, "y": 2}}}}

    def test_get_registry_thesis_is_default_instance(self):
        assert get_registry("thesis") is default_registry

    def test_get_registry_unknown_domains_are_cached_and_isolated(self):
        first = get_registry("boundary-domain-a")
        second = get_registry("boundary-domain-a")
        other = get_registry("boundary-domain-b")

        assert first is second
        assert first is not other
        assert first is not default_registry

    def test_items_view_reflects_registry(self):
        reg = Registry()

        class Node(FormatNode):
            pass

        reg.register("a")(Node)

        assert dict(reg.items()) == {"a": Node}


# ---------------------------------------------------------------------------
# create_node
# ---------------------------------------------------------------------------


class TestCreateNodeBoundary:
    def test_known_category_returns_instance_with_item_identity(self):
        item = {"category": "body_text", "paragraph": "正文"}

        node = create_node(item, level=7, config={})

        assert isinstance(node, BodyText)
        assert node.value is item
        assert node.level == 7

    @pytest.mark.parametrize("item", [{}, {"category": ""}, {"category": None}])
    def test_missing_category_raises_value_error(self, item):
        with pytest.raises(ValueError):
            create_node(item, level=0, config={})

    def test_unknown_category_returns_none(self):
        assert create_node({"category": "无此类别"}, level=0, config={}) is None

    def test_none_config_is_accepted(self):
        node = create_node({"category": "body_text"}, level=0, config=None)

        assert node is not None
        assert node.pydantic_config.font_size == "小四"

    def test_custom_registry_is_used(self):
        reg = Registry()

        class Custom(FormatNode):
            NODE_TYPE = "custom.node"
            NODE_LABEL = "自定义"

        reg.register("custom")(Custom)

        node = create_node(
            {"category": "custom", "paragraph": "x"},
            level=3,
            config={},
            registry=reg,
        )

        assert isinstance(node, Custom)
        assert node.level == 3

    def test_custom_category_without_registry_is_unknown(self):
        assert create_node({"category": "custom"}, level=0, config={}) is None

    def test_empty_style_config_does_not_crash_on_check(self):
        from docx import Document

        class Node(FormatNode):
            NODE_TYPE = "probe"

        node = Node(value={}, level=0, paragraph=Document().add_paragraph("中文"))
        node.load_config({"probe": {"chinese_font_name": "宋体"}})
        node.check_format(Document())


# ---------------------------------------------------------------------------
# DocumentTreeBuilder / DocumentBuilder
# ---------------------------------------------------------------------------


class TestTreeBuilderBoundary:
    def test_empty_items_returns_bare_root(self):
        root = DocumentTreeBuilder().build_tree([])

        assert root.value["category"] == "top"
        assert root.children == []

    def test_heading_hierarchy_and_body_attachment(self):
        items = [
            {"category": "heading_level_1", "paragraph": "第一章"},
            {"category": "heading_level_2", "paragraph": "1.1"},
            {"category": "body_text", "paragraph": "正文"},
            {"category": "heading_level_1", "paragraph": "第二章"},
        ]

        root = DocumentTreeBuilder().build_tree(items)

        assert [c.value.get("paragraph") for c in root.children] == [
            "第一章",
            "第二章",
        ]
        first_h1 = root.children[0]
        assert [c.value.get("paragraph") for c in first_h1.children] == ["1.1"]
        h2 = first_h1.children[0]
        assert [c.value["paragraph"] for c in h2.children] == ["正文"]

    def test_body_before_heading_attaches_to_root(self):
        root = DocumentTreeBuilder().build_tree(
            [{"category": "body_text", "paragraph": "前言"}]
        )

        assert [c.value["paragraph"] for c in root.children] == ["前言"]

    def test_same_level_headings_are_siblings(self):
        root = DocumentTreeBuilder().build_tree(
            [
                {"category": "heading_level_1", "paragraph": "A"},
                {"category": "heading_level_1", "paragraph": "B"},
            ]
        )

        assert [c.value["paragraph"] for c in root.children] == ["A", "B"]
        assert root.children[0].children == []

    def test_unknown_category_is_skipped_without_shifting_others(self):
        root = DocumentTreeBuilder().build_tree(
            [
                {"category": "unknown_x", "paragraph": "丢弃"},
                {"category": "body_text", "paragraph": "保留"},
            ]
        )

        assert [c.value["paragraph"] for c in root.children] == ["保留"]

    def test_custom_registry_heading_is_created(self):
        reg = Registry()

        class CustomHeading(FormatNode):
            NODE_TYPE = "custom.heading"
            NODE_LABEL = "自定义标题"

        reg.register("custom_heading", level=1)(CustomHeading)

        root = DocumentTreeBuilder(registry=reg).build_tree(
            [{"category": "custom_heading", "paragraph": "自定义标题"}]
        )

        assert len(root.children) == 1
        assert isinstance(root.children[0], CustomHeading)


class TestDocumentBuilderBoundary:
    def test_load_paragraphs_returns_list_as_is(self):
        items = [{"category": "body_text", "paragraph": "x"}]

        assert DocumentBuilder.load_paragraphs(items) is items

    def test_load_paragraphs_parses_json_string(self):
        items = [{"category": "body_text", "paragraph": "x"}]

        assert DocumentBuilder.load_paragraphs(json.dumps(items)) == items

    def test_load_paragraphs_reads_file_when_not_json(self, tmp_path):
        path = _write(
            tmp_path,
            json.dumps([{"category": "body_text", "paragraph": "x"}]),
            name="doc.json",
        )

        assert DocumentBuilder.load_paragraphs(path) == [
            {"category": "body_text", "paragraph": "x"}
        ]

    def test_load_paragraphs_missing_path_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            DocumentBuilder.load_paragraphs(str(tmp_path / "missing.json"))

    def test_build_from_json_list_and_string(self):
        items = [{"category": "heading_level_1", "paragraph": "标题"}]

        from_list = DocumentBuilder.build_from_json(items, config={})
        from_text = DocumentBuilder.build_from_json(json.dumps(items), config={})

        assert [c.value["paragraph"] for c in from_list.children] == ["标题"]
        assert [c.value["paragraph"] for c in from_text.children] == ["标题"]

    def test_build_from_json_empty(self):
        root = DocumentBuilder.build_from_json([], config={})

        assert root.value["category"] == "top"
        assert root.children == []
