"""
rules 模块测试 —— 聚焦真实行为验证，无填充。
"""

from unittest.mock import patch

import pytest
from docx import Document
from docx.shared import Pt

from wordformat.rules import (
    KeywordsCN,
    KeywordsEN,
)

# ---------------------------------------------------------------------------
# 共享 fixtures / helpers
# ---------------------------------------------------------------------------


def _make_node(cls, text="测试文本"):
    """创建一个带 paragraph 的 FormatNode 实例。"""
    doc = Document()
    p = doc.add_paragraph(text)
    return cls(value=p, level=0, paragraph=p)


def _load_root_config(config_path):
    """从 YAML 路径加载配置 dict。"""
    return _load_yaml(config_path)


def _load_yaml(path):
    import yaml

    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


@pytest.fixture
def root_config(sample_yaml_config):
    """从 sample_yaml_config 加载 NodeConfigRoot，与示例文件解耦。"""
    from wordformat.config.loader import init_config

    init_config(sample_yaml_config)
    return _load_root_config(sample_yaml_config)


@pytest.fixture
def doc():
    return Document()


@pytest.fixture
def para(doc):
    return doc.add_paragraph("测试段落")


@pytest.fixture
def run_with_text(para):
    r = para.add_run("关键词")
    r.font.size = Pt(12)
    r.font.bold = True
    return r


# ---------------------------------------------------------------------------
# 1. FormatNode 基类行为
# ---------------------------------------------------------------------------


class TestKeywordsLogic:
    """关键词节点的标签识别、数量校验、标点校验。"""

    def test_cn_label_detection(self):
        """中文关键词标签识别（使用真实 paragraph runs）。"""
        doc = Document()
        p = doc.add_paragraph("关键词")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        assert node._check_keyword_label(p.runs[0]) is True

        p2 = doc.add_paragraph("关 键 词")
        node.paragraph = p2
        assert node._check_keyword_label(p2.runs[0]) is True

        p3 = doc.add_paragraph("机器学习")
        node.paragraph = p3
        assert node._check_keyword_label(p3.runs[0]) is False

    def test_en_label_detection(self):
        """英文关键词标签识别（使用真实 paragraph runs）。"""
        doc = Document()
        p = doc.add_paragraph("Keywords")
        node = KeywordsEN(value=p, level=0, paragraph=p)
        assert node._check_keyword_label(p.runs[0]) is True

        p2 = doc.add_paragraph("Keyword")
        node.paragraph = p2
        assert node._check_keyword_label(p2.runs[0]) is True

        p3 = doc.add_paragraph("KEY WORDS")
        node.paragraph = p3
        assert node._check_keyword_label(p3.runs[0]) is True

        p4 = doc.add_paragraph("machine learning")
        node.paragraph = p4
        assert node._check_keyword_label(p4.runs[0]) is False

    def test_cn_count_validation_too_few(self, root_config):
        """中文关键词数量不足时应触发 add_comment。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("关键词：机器学习")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment") as mock_comment:
            node.check_format(doc)
        # 应至少有一条数量不足的 comment
        texts = [c.kwargs["text"] for c in mock_comment.call_args_list]
        assert any("数量过少" in t for t in texts)

    def test_cn_count_validation_too_many(self, root_config):
        """中文关键词数量超限时应触发 add_comment。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("关键词：A；B；C；D；E；F")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment") as mock_comment:
            node.check_format(doc)
        texts = [c.kwargs["text"] for c in mock_comment.call_args_list]
        assert any("数量过多" in t for t in texts)

    def test_cn_trailing_punct_detection(self, root_config):
        """中文关键词末尾标点校验。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("关键词：机器学习；深度学习；")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment") as mock_comment:
            node.check_format(doc)
        texts = [c.kwargs["text"] for c in mock_comment.call_args_list]
        assert any("标点错误" in t for t in texts)

    def test_en_count_validation_too_few(self, root_config):
        """英文关键词数量不足。"""
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Keywords: AI")
        node = KeywordsEN(value=p, level=0, paragraph=p)
        node.load_config(root_config)
        with patch.object(node, "add_comment") as mock_comment:
            node.check_format(doc)
        texts = [c.kwargs["text"] for c in mock_comment.call_args_list]
        assert any("数量过少" in t for t in texts)


# ---------------------------------------------------------------------------
# 6. _base 实现实际调用 diff/apply 逻辑
# ---------------------------------------------------------------------------


class TestKeywordsENBase:
    """英文关键词的字符样式与数量校验（_handle_character_style / _check_keyword_count）。"""

    def _make_en_node(self, config_dict=None):
        """Helper to create a KeywordsEN node with config loaded"""
        from wordformat.rules.keywords import KeywordsEN

        node = KeywordsEN(
            value={"category": "abstract.keywords.english", "fingerprint": "fp"},
            level=1,
        )
        if config_dict:
            node.load_config(config_dict)
        return node

    @staticmethod
    def _comments(node, doc, **kwargs):
        with patch.object(node, "add_comment") as mock_comment:
            node._handle_character_style(doc, None, **kwargs)
        return [c.kwargs["text"] for c in mock_comment.call_args_list]

    def test_empty_run_skip(self, sample_yaml_config):
        """空白 run 不参与字符样式检查。"""
        from wordformat.config.loader import get_config, init_config

        init_config(sample_yaml_config)
        node = self._make_en_node(get_config())
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("   ")
        node.paragraph = p
        assert self._comments(node, doc, p=True) == []

    def test_label_style_check(self, sample_yaml_config):
        """标签加粗错误应报告到“英文关键词标签”。"""
        from wordformat.config.loader import get_config, init_config

        init_config(sample_yaml_config)
        node = self._make_en_node(get_config())
        doc = Document()
        p = doc.add_paragraph()
        label_run = p.add_run("Keywords: ")
        label_run.font.bold = False
        node.paragraph = p
        texts = self._comments(node, doc, p=True)
        assert any("英文关键词标签" in t and "加粗错误" in t for t in texts)

    def test_content_style_check(self, sample_yaml_config):
        """内容加粗错误应报告到“英文关键词内容”。"""
        from wordformat.config.loader import get_config, init_config

        init_config(sample_yaml_config)
        node = self._make_en_node(get_config())
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Keywords: ").font.bold = True
        p.add_run("AI, ML").font.bold = True
        node.paragraph = p
        texts = self._comments(node, doc, p=True)
        assert any("英文关键词内容" in t and "加粗错误" in t for t in texts)

    def test_keyword_count_validation_min(self, sample_yaml_config):
        """英文关键词少于下限：报告当前数量。"""
        from wordformat.config.loader import get_config, init_config

        init_config(sample_yaml_config)
        node = self._make_en_node(get_config())
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Keywords: ").font.bold = True
        p.add_run("AI")
        node.paragraph = p
        with patch.object(node, "add_comment") as mock_comment:
            node.check_format(doc)
        texts = [c.kwargs["text"] for c in mock_comment.call_args_list]
        assert any("数量过少" in t and "当前1个" in t for t in texts)

    def test_keyword_count_validation_max(self, sample_yaml_config):
        """英文关键词多于上限：报告当前数量。"""
        from wordformat.config.loader import get_config, init_config

        init_config(sample_yaml_config)
        node = self._make_en_node(get_config())
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("Keywords: ").font.bold = True
        p.add_run("AI, ML, NLP, CV, DB, SE")
        node.paragraph = p
        with patch.object(node, "add_comment") as mock_comment:
            node.check_format(doc)
        texts = [c.kwargs["text"] for c in mock_comment.call_args_list]
        assert any("数量过多" in t and "当前6个" in t for t in texts)


class TestKeywordRuleBranches:
    """关键词数量/末尾标点/标签拆分的边界与关闭开关。"""

    @staticmethod
    def _make(cls, text, config=None):
        doc = Document()
        p = doc.add_paragraph(text)
        node = cls(value=p, level=0, paragraph=p)
        node.load_config(config or {})
        return doc, node

    @staticmethod
    def _comments(node, doc, mode="check"):
        with patch.object(node, "add_comment") as mock_comment:
            if mode == "check":
                node.check_format(doc)
            else:
                node.apply_format(doc)
        return [call.kwargs["text"] for call in mock_comment.call_args_list]

    @pytest.mark.parametrize("keywords", ["A；B；C", "A；B；C；D；E"])
    def test_count_within_range(self, keywords):
        doc, node = self._make(KeywordsCN, f"关键词：{keywords}")
        texts = self._comments(node, doc)
        assert not any("数量" in text for text in texts)

    @pytest.mark.parametrize(
        ("keywords", "label"),
        [("A；B", "数量过少"), ("A；B；C；D；E；F", "数量过多")],
    )
    def test_count_out_of_range(self, keywords, label):
        doc, node = self._make(KeywordsCN, f"关键词：{keywords}")
        texts = self._comments(node, doc)
        expected_count = len(keywords.split("；"))
        assert any(
            label in text and f"当前{expected_count}个" in text for text in texts
        )

    def test_custom_count_range(self):
        cfg = {
            "abstract": {
                "chinese": {
                    "keywords": {
                        "rules": {"keyword_count": {"count_min": 2, "count_max": 2}}
                    }
                }
            }
        }
        doc, node = self._make(KeywordsCN, "关键词：A；B；C", cfg)
        assert any("数量过多" in text for text in self._comments(node, doc))

    def test_trailing_punctuation_ok(self):
        doc, node = self._make(KeywordsCN, "关键词：A；B；C")
        assert not any("标点错误" in text for text in self._comments(node, doc))

    def test_trailing_punctuation_reported(self):
        doc, node = self._make(KeywordsCN, "关键词：A；B；C；")
        texts = self._comments(node, doc)
        issue = next(text for text in texts if "标点错误" in text)
        assert "末尾有'；'" in issue

    def test_trailing_punctuation_can_be_disabled(self):
        cfg = {
            "abstract": {
                "chinese": {
                    "keywords": {"rules": {"trailing_punctuation": {"enabled": False}}}
                }
            }
        }
        doc, node = self._make(KeywordsCN, "关键词：A；B；C；", cfg)
        assert not any("标点错误" in text for text in self._comments(node, doc))

    def test_split_mixed_runs_cn(self):
        doc, node = self._make(KeywordsCN, "关键词：A；B")
        node._split_mixed_runs()
        assert [run.text for run in doc.paragraphs[0].runs] == ["关键词：", "A；B"]

    def test_split_mixed_runs_en(self):
        doc, node = self._make(KeywordsEN, "Keywords: AI, ML")
        node._split_mixed_runs()
        assert [run.text for run in doc.paragraphs[0].runs] == [
            "Keywords: ",
            "AI, ML",
        ]

    @pytest.mark.parametrize(
        ("cls", "text", "expected"),
        [
            (KeywordsCN, "关键词", True),
            (KeywordsCN, "关 键 词", True),
            (KeywordsCN, "机器学习", False),
            (KeywordsEN, "Keywords", True),
            (KeywordsEN, "ML", False),
        ],
    )
    def test_label_detection_without_paragraph(self, cls, text, expected):
        doc = Document()
        run = doc.add_paragraph().add_run(text)
        node = cls(value={}, level=0, paragraph=None)
        assert node._check_keyword_label(run) is expected

    def test_blank_run_is_not_label(self):
        doc = Document()
        run = doc.add_paragraph().add_run("   ")
        for cls in (KeywordsCN, KeywordsEN):
            node = cls(value={}, level=0, paragraph=None)
            assert node._check_keyword_label(run) is False

    def test_foreign_run_is_not_label(self):
        doc = Document()
        run = doc.add_paragraph().add_run("关键词")
        p = doc.add_paragraph("关键词：A")
        for cls in (KeywordsCN, KeywordsEN):
            node = cls(value=p, level=0, paragraph=p)
            assert node._check_keyword_label(run) is False

    def test_en_label_must_be_at_start(self):
        doc, node = self._make(KeywordsEN, "前言 Keywords: AI")
        assert node._check_keyword_label(doc.paragraphs[0].runs[0]) is False

    def test_character_style_skips_blank_run(self):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("关键词：")
        p.add_run("   ")
        p.add_run("机器学习")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config({})
        with patch.object(node, "add_comment") as mock_comment:
            node._handle_character_style(doc, None, p=False)
        assert mock_comment.call_args_list
        assert all(
            call.kwargs["runs"].text.strip() for call in mock_comment.call_args_list
        )


class TestKeywordSplitEdges:
    def test_split_skips_blank_run(self):
        doc = Document()
        p = doc.add_paragraph()
        p.add_run("   ")
        p.add_run("关键词：A；B")
        node = KeywordsCN(value=p, level=0, paragraph=p)
        node.load_config({})
        node._split_mixed_runs()
        assert [run.text for run in p.runs] == ["   ", "关键词：", "A；B"]

    def test_en_foreign_run_with_label_paragraph(self):
        doc = Document()
        foreign = doc.add_paragraph().add_run("AI")
        p = doc.add_paragraph("Keywords: AI")
        node = KeywordsEN(value=p, level=0, paragraph=p)
        assert node._check_keyword_label(foreign) is False
