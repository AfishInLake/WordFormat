"""正文段落字数/句数规则（rules/body.py）。"""

from unittest.mock import patch

import pytest
from docx import Document

from wordformat.rules.body import BodyText, _length_limit, _paragraph_size


def _node(text, config=None):
    doc = Document()
    p = doc.add_paragraph(text)
    node = BodyText(value={"category": "body_text"}, level=1, paragraph=p)
    node.load_config(config or {})
    return doc, node


def _issues(node, doc, rule_cfg):
    with patch.object(node, "add_comment") as mock_comment:
        node._check_paragraph_length(doc, rule_cfg)
    return [call.kwargs["text"] for call in mock_comment.call_args_list]


class TestParagraphSize:
    @pytest.mark.parametrize(
        ("text", "expected"),
        [
            ("你好世界。", (5, 1)),
            ("第一句。第二句！Third?", (14, 3)),
            ("你好 世界\n。", (5, 1)),
            ("没有句末标点", (6, 1)),
            ("", (0, 0)),
            ("   ", (0, 0)),
        ],
    )
    def test_counts(self, text, expected):
        assert _paragraph_size(text) == expected

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (None, None),
            (0, None),
            (-3, None),
            (True, None),
            (30, 30),
            ("30", 30),
            ("30.9", 30),
            ("abc", None),
        ],
    )
    def test_limit_normalization(self, value, expected):
        assert _length_limit(value) == expected


class TestParagraphLengthRule:
    def test_default_thresholds_report_nothing(self):
        doc, node = _node("短句。")
        assert _issues(node, doc, node.pydantic_config.rules.paragraph_length) == []

    def test_min_chars(self):
        doc, node = _node("短句。")
        texts = _issues(node, doc, {"min_chars": 10})
        assert texts == ["正文段落-字数过少：当前 3 字，规范：不少于 10 字"]

    def test_max_chars(self):
        doc, node = _node("这是一段五个字的文本。")
        texts = _issues(node, doc, {"max_chars": 5})
        assert texts == ["正文段落-字数过多：当前 11 字，规范：不超过 5 字"]

    def test_min_sentences(self):
        doc, node = _node("只有一句。")
        texts = _issues(node, doc, {"min_sentences": 2})
        assert texts == ["正文段落-句数过少：当前 1 句，规范：不少于 2 句"]

    def test_max_sentences(self):
        doc, node = _node("第一句。第二句。第三句。")
        texts = _issues(node, doc, {"max_sentences": 2})
        assert texts == ["正文段落-句数过多：当前 3 句，规范：不超过 2 句"]

    def test_boundary_values_pass(self):
        doc, node = _node("这是文本。")
        assert _issues(node, doc, {"min_chars": 5, "max_chars": 5}) == []
        assert _issues(node, doc, {"min_sentences": 1, "max_sentences": 1}) == []

    def test_multiple_thresholds_reported_together(self):
        doc, node = _node("短。")
        texts = _issues(node, doc, {"min_chars": 10, "min_sentences": 2})
        assert len(texts) == 2
        assert any("字数过少" in t for t in texts)
        assert any("句数过少" in t for t in texts)

    def test_empty_paragraph_skipped(self):
        doc, node = _node("")
        assert _issues(node, doc, {"min_chars": 10}) == []

    def test_string_thresholds(self):
        doc, node = _node("短句。")
        assert _issues(node, doc, {"min_chars": "10"})[0].endswith("不少于 10 字")


class TestRuleWiring:
    def test_enabled_by_default_with_zero_limits(self):
        doc, node = _node("短句。")
        cfg = node.pydantic_config.rules.paragraph_length
        assert cfg.enabled is True
        assert cfg.max_chars == 0

    def test_configured_through_check_format(self):
        doc, node = _node(
            "短句。",
            {
                "body": {
                    "text": {
                        "rules": {"paragraph_length": {"max_chars": 1}},
                    }
                }
            },
        )
        with patch.object(node, "add_comment") as mock_comment:
            node.check_format(doc)
        texts = [call.kwargs["text"] for call in mock_comment.call_args_list]
        assert any("字数过多" in t for t in texts)

    def test_disabled_rule_skipped(self):
        doc, node = _node(
            "短句。",
            {
                "body": {
                    "text": {
                        "rules": {
                            "paragraph_length": {"enabled": False, "max_chars": 1}
                        },
                    }
                }
            },
        )
        with patch.object(node, "add_comment") as mock_comment:
            node.check_format(doc)
        assert not any(
            "字数过多" in call.kwargs["text"] for call in mock_comment.call_args_list
        )
