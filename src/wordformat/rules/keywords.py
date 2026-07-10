import re
from copy import deepcopy
from typing import Literal

from docx.oxml.ns import qn

from wordformat.config.models import (
    KeywordCountRule,
    KeywordsConfig,
    NodeConfigRoot,
    TrailingPunctRule,
)
from wordformat.rules.node import FormatNode
from wordformat.style.diff import CharacterStyle, ParagraphStyle


# 第一步：提取关键词基类，复用通用逻辑
class BaseKeywordsNode(FormatNode[KeywordsConfig]):
    """关键词节点基类（复用中英文通用逻辑）"""

    # 子类必须定义的属性
    LANG: Literal["cn", "en"] = ""  # 语言类型：cn/en
    NODE_TYPE: str = ""
    CONFIG_MODEL = KeywordsConfig

    def _get_lang_config(self, root_config: NodeConfigRoot) -> KeywordsConfig:
        """根据语言类型获取对应配置"""
        lang_config_map = {
            "cn": root_config.abstract.keywords["chinese"],
            "en": root_config.abstract.keywords["english"],
        }
        return lang_config_map.get(self.LANG, root_config.abstract.keywords["chinese"])

    def load_config(self, root_config: dict | NodeConfigRoot):
        """重载加载配置，自动匹配对应语言的关键词配置"""
        if isinstance(root_config, dict):
            # 从字典中提取对应语言的配置，通过父类方法设置 __config
            lang_config = (
                root_config.get("abstract", {}).get("keywords", {}).get(self.LANG, {})
            )
            # 直接设置父类的 __config（避免名称修饰问题）
            self._TreeNode__config = lang_config
            self._pydantic_config = self.CONFIG_MODEL(**self.config)
        elif isinstance(root_config, NodeConfigRoot):
            # 从Pydantic模型提取对应语言的配置
            self._pydantic_config = self._get_lang_config(root_config)
            self._TreeNode__config = self._pydantic_config.model_dump()
        else:
            raise TypeError(
                f"配置类型不支持：{type(root_config)}，仅支持dict或NodeConfigRoot"
            )

    def _check_paragraph_style(self, cfg: KeywordsConfig, p: bool) -> str:
        """通用段落样式检查（复用）"""
        ps = ParagraphStyle.from_config(cfg)
        if p:
            issues = ps.diff_from_paragraph(self.paragraph)
        else:
            issues = ps.apply_to_paragraph(self.paragraph)
        return ParagraphStyle.to_string(issues, target=self.NODE_LABEL)

    def _split_mixed_runs(self) -> None:
        """检测标签和内容混合在同一个 run 中的情况，拆分为两个独立的 run。

        例如 "关键词：校园二手交易；Django" 在同一个 run 中时，
        拆分为 "关键词：" 和 "校园二手交易；Django" 两个 run，
        以便分别应用不同的字符样式。

        拆分后第二个 run 会继承第一个 run 的所有格式（rPr），
        后续样式检查会覆盖需要修改的属性。
        """
        label_pattern = self._get_label_split_pattern()
        if not label_pattern:
            return

        runs_to_split = []

        for run in self.paragraph.runs:
            text = run.text
            if not text.strip():
                continue
            match = re.search(label_pattern, text)
            if match and match.end() < len(text):
                # 标签后面还有内容，需要拆分
                split_pos = match.end()
                runs_to_split.append((run, split_pos))

        for run, split_pos in runs_to_split:
            r_elem = run._element
            t_elem = r_elem.find(qn("w:t"))
            if t_elem is None:
                continue

            full_text = t_elem.text or ""
            label_text = full_text[:split_pos]
            content_text = full_text[split_pos:]

            # 修改原 run 为标签部分
            t_elem.text = label_text

            # 创建新 run（深拷贝原 run 的 XML，保留所有格式）
            new_r_elem = deepcopy(r_elem)
            new_t_elem = new_r_elem.find(qn("w:t"))
            if new_t_elem is not None:
                new_t_elem.text = content_text

            # 在原 run 后面插入新 run
            r_elem.addnext(new_r_elem)

    def _get_label_split_pattern(self) -> re.Pattern | None:
        """返回用于拆分标签的正则表达式，由子类重写。"""
        return None

    # 第二步：英文关键词节点（专属逻辑）


class KeywordsEN(BaseKeywordsNode):
    """关键词节点-英文"""

    LANG = "en"
    NODE_TYPE = "abstract.keywords.english"
    NODE_LABEL = "英文关键词"
    RULES = {"keyword_count": "_check_keyword_count"}

    _LABEL_RE = re.compile(r"Keywords?:?\s*", re.IGNORECASE)
    _SEPARATOR_RE = re.compile(r"[,;]")

    @staticmethod
    def extract_keywords(text: str) -> list[str]:
        """从文本中提取英文关键词列表（去除标签前缀）。"""
        kw_text = KeywordsEN._LABEL_RE.sub("", text)
        return [k.strip() for k in KeywordsEN._SEPARATOR_RE.split(kw_text) if k.strip()]

    def _check_keyword_label(self, run) -> bool:
        """判断 run 是否属于英文关键词标签部分。"""
        if not run.text.strip():
            return False
        if self.paragraph is None:
            return bool(re.search(r"Keywords?|KEY\s*WORDS", run.text, re.IGNORECASE))
        full = "".join(r.text for r in self.paragraph.runs)
        m = re.match(r"Keywords?\s*[:：]?\s*|KEY\s*WORDS\s*", full, re.IGNORECASE)
        if not m:
            return False
        label_end = m.end()
        # 找出当前 run 在全文中的位置（用 XML 元素 identity 比较）
        pos = 0
        for r in self.paragraph.runs:
            rl = len(r.text)
            if r._element is run._element:
                return pos < label_end
            pos += rl
        return False

    def _get_label_split_pattern(self) -> re.Pattern | None:
        """英文标签拆分模式：匹配 'Keywords:' 或 'Keywords ' 及其变体"""
        return re.compile(r"Keywords?\s*[:：]?\s*", re.IGNORECASE)

    def _check_keyword_count(self, doc, rule_cfg: KeywordCountRule, p: bool = False):
        """校验英文关键词数量"""
        keyword_text = "".join([run.text for run in self.paragraph.runs])
        keyword_list = KeywordsEN.extract_keywords(keyword_text)
        from wordformat.style.comments import format_comment

        target = self.NODE_LABEL
        if len(keyword_list) < rule_cfg.count_min:
            issue = format_comment(
                target,
                "数量过少",
                f"当前{len(keyword_list)}个",
                f"{rule_cfg.count_min}-{rule_cfg.count_max}个",
            )
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)
        if len(keyword_list) > rule_cfg.count_max:
            issue = format_comment(
                target,
                "数量过多",
                f"当前{len(keyword_list)}个",
                f"{rule_cfg.count_min}-{rule_cfg.count_max}个",
            )
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)

    # 覆盖默认 handler：区分标签与内容的字符样式
    def _handle_character_style(self, doc, rule_cfg, p: bool):
        cfg = self.pydantic_config
        label_style = CharacterStyle(
            font_name_cn=cfg.label.chinese_font_name,
            font_name_en=cfg.label.english_font_name,
            font_size=cfg.label.font_size,
            font_color=cfg.label.font_color,
            bold=cfg.label.bold,
            italic=cfg.label.italic,
            underline=cfg.label.underline,
        )
        content_style = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )
        for run in self.paragraph.runs:
            if not run.text.strip():
                continue
            is_label = self._check_keyword_label(run)
            cstyle = label_style if is_label else content_style
            target = f"{self.NODE_LABEL}标签" if is_label else f"{self.NODE_LABEL}内容"
            if p:
                diff = cstyle.diff_from_run(run)
            else:
                diff = cstyle.apply_to_run(run)
            if diff:
                self.add_comment(
                    doc=doc,
                    runs=run,
                    text=CharacterStyle.to_string(diff, target=target),
                )

    def _base(self, doc, p: bool, r: bool):
        """拆分标签和内容混合的 run（仅在应用模式下执行）"""
        if not p:
            self._split_mixed_runs()


# 第三步：中文关键词节点（专属逻辑）
class KeywordsCN(BaseKeywordsNode):
    """关键词节点-中文"""

    LANG = "cn"
    NODE_TYPE = "abstract.keywords.chinese"
    NODE_LABEL = "中文关键词"
    RULES = {
        "keyword_count": "_check_keyword_count",
        "trailing_punctuation": "_check_trailing_punctuation",
    }

    _LABEL_RE = re.compile(r"关键词[:：]?\s*")
    _SEPARATOR_RE = re.compile(r"[；;]")

    @staticmethod
    def extract_keywords(text: str) -> list[str]:
        """从文本中提取中文关键词列表（去除标签前缀）。"""
        kw_text = KeywordsCN._LABEL_RE.sub("", text)
        return [k.strip() for k in KeywordsCN._SEPARATOR_RE.split(kw_text) if k.strip()]

    def _check_keyword_label(self, run) -> bool:
        """判断 run 是否属于中文关键词标签部分（防拆分）。"""
        if not run.text.strip():
            return False
        if self.paragraph is None:
            p = r"关[^a-zA-Z0-9\u4e00-\u9fff]*键[^a-zA-Z0-9\u4e00-\u9fff]*词"
            return bool(re.search(p, run.text))
        full = "".join(r.text for r in self.paragraph.runs)
        m = re.search(
            r"关[^a-zA-Z0-9\u4e00-\u9fff]*键[^a-zA-Z0-9\u4e00-\u9fff]*词\s*[:：]?\s*",
            full,
        )
        if not m:
            return False
        label_end = m.end()
        pos = 0
        for r in self.paragraph.runs:
            rl = len(r.text)
            if r._element is run._element:
                return pos < label_end
            pos += rl
        return False

    def _get_label_split_pattern(self) -> re.Pattern | None:
        """中文标签拆分模式：匹配 '关键词：' 或 '关键词:' 及其变体"""
        return re.compile(
            r"关[^a-zA-Z0-9\u4e00-\u9fff]*键[^a-zA-Z0-9\u4e00-\u9fff]*词\s*[:：]?\s*"
        )

    def _check_keyword_count(self, doc, rule_cfg: KeywordCountRule, p: bool = False):
        """校验中文关键词数量"""
        keyword_text = "".join([run.text for run in self.paragraph.runs])
        keyword_list = KeywordsCN.extract_keywords(keyword_text)
        from wordformat.style.comments import format_comment

        target = self.NODE_LABEL
        if len(keyword_list) < rule_cfg.count_min:
            issue = format_comment(
                target,
                "数量过少",
                f"当前{len(keyword_list)}个",
                f"{rule_cfg.count_min}-{rule_cfg.count_max}个",
            )
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)
        if len(keyword_list) > rule_cfg.count_max:
            issue = format_comment(
                target,
                "数量过多",
                f"当前{len(keyword_list)}个",
                f"{rule_cfg.count_min}-{rule_cfg.count_max}个",
            )
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)

    def _check_trailing_punctuation(
        self, doc, rule_cfg: TrailingPunctRule, p: bool = False
    ):
        """校验中文关键词末尾标点"""
        from wordformat.style.comments import format_comment

        keyword_text = "".join([run.text for run in self.paragraph.runs])
        if (
            keyword_text.strip()
            and keyword_text.strip()[-1] in rule_cfg.forbidden_chars
        ):
            punct = keyword_text.strip()[-1]
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=format_comment(
                    self.NODE_LABEL,
                    "标点错误",
                    f"末尾有'{punct}'",
                    "末尾无标点",
                ),
            )

    # 覆盖默认 handler：区分标签与内容的字符样式
    def _handle_character_style(self, doc, rule_cfg, p: bool):
        cfg = self.pydantic_config
        label_style = CharacterStyle(
            font_name_cn=cfg.label.chinese_font_name,
            font_name_en=cfg.label.english_font_name,
            font_size=cfg.label.font_size,
            font_color=cfg.label.font_color,
            bold=cfg.label.bold,
            italic=cfg.label.italic,
            underline=cfg.label.underline,
        )
        content_style = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )
        for run in self.paragraph.runs:
            if not run.text.strip():
                continue
            is_label = self._check_keyword_label(run)
            cstyle = label_style if is_label else content_style
            target = f"{self.NODE_LABEL}标签" if is_label else f"{self.NODE_LABEL}内容"
            if p:
                diff = cstyle.diff_from_run(run)
            else:
                diff = cstyle.apply_to_run(run)
            if diff:
                self.add_comment(
                    doc=doc,
                    runs=run,
                    text=CharacterStyle.to_string(diff, target=target),
                )

    def _base(self, doc, p: bool, r: bool):
        """拆分标签和内容混合的 run（仅在应用模式下执行）"""
        if not p:
            self._split_mixed_runs()
