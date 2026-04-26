import re
from copy import deepcopy
from typing import Literal

from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from wordformat.config.datamodel import KeywordsConfig, NodeConfigRoot
from wordformat.rules.node import FormatNode
from wordformat.style.check_format import CharacterStyle, ParagraphStyle



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
        return ParagraphStyle.to_string(issues)

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

        p_elem = self.paragraph._element
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

    def _check_keyword_label(self, run) -> bool:
        """检查run是否包含英文关键词标签（Keywords/KEY WORDS）"""
        pattern = r"Keywords?|KEY\s*WORDS"
        return bool(re.search(pattern, run.text, re.IGNORECASE))

    def _get_label_split_pattern(self) -> re.Pattern | None:
        """英文标签拆分模式：匹配 'Keywords:' 或 'Keywords ' 及其变体"""
        return re.compile(r"Keywords?\s*[:：]?\s*", re.IGNORECASE)

    def _base(self, doc, p: bool, r: bool):  # noqa C901
        """
        校验英文关键词格式：
        - 段落整体格式（对齐、行距等）
        - "Keywords:" 部分加粗，其余内容不加粗
        """
        cfg = self.pydantic_config

        # 2. 段落样式检查
        paragraph_issues = self._check_paragraph_style(cfg, p)
        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=paragraph_issues,
            )

        # 3. 拆分标签和内容混合的 run（仅在格式化模式下执行）
        if not p:
            self._split_mixed_runs()

        # 4. 字符样式检查（区分标签/内容）
        # (a) "Keywords:" 标签 —— 加粗
        label_style = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.keywords_bold,  # 修复拼写错误：kewords → keywords
            italic=cfg.italic,
            underline=cfg.underline,
        )
        # (b) 关键词内容 —— 不加粗
        content_style = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        # 4. 遍历run校验
        for run in self.paragraph.runs:
            if not run.text.strip():
                continue

            if self._check_keyword_label(run):
                # 检查标签样式
                if r:
                    diff = label_style.diff_from_run(run)
                else:
                    diff = label_style.apply_to_run(run)
                if diff:
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=CharacterStyle.to_string(diff),
                    )
            else:
                # 检查内容样式
                if r:
                    diff = content_style.diff_from_run(run)
                else:
                    diff = content_style.apply_to_run(run)
                if diff:
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=CharacterStyle.to_string(diff),
                    )

        # 5. 校验关键词数量（KeywordsConfig特有配置）
        keyword_text = "".join([run.text for run in self.paragraph.runs])
        # 提取英文关键词（按逗号/分号分割）
        keyword_list = re.split(
            r"[,;]", re.sub(r"Keywords?:", "", keyword_text, flags=re.IGNORECASE)
        )
        keyword_list = [k.strip() for k in keyword_list if k.strip()]
        if len(keyword_list) < cfg.count_min:
            issue = f"英文关键词数量不足（最少{cfg.count_min}个，当前{len(keyword_list)}个）"
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)
        if len(keyword_list) > cfg.count_max:
            issue = f"英文关键词数量超限（最多{cfg.count_max}个，当前{len(keyword_list)}个）"
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)


# 第三步：中文关键词节点（专属逻辑）
class KeywordsCN(BaseKeywordsNode):
    """关键词节点-中文"""

    LANG = "cn"
    NODE_TYPE = "abstract.keywords.chinese"

    def _check_keyword_label(self, run) -> bool:
        """检查run是否包含中文关键词标签（关键词）"""
        pattern = r"关[^a-zA-Z0-9\u4e00-\u9fff]*键[^a-zA-Z0-9\u4e00-\u9fff]*词"
        return bool(re.search(pattern, run.text))

    def _get_label_split_pattern(self) -> re.Pattern | None:
        """中文标签拆分模式：匹配 '关键词：' 或 '关键词:' 及其变体"""
        return re.compile(r"关[^a-zA-Z0-9\u4e00-\u9fff]*键[^a-zA-Z0-9\u4e00-\u9fff]*词\s*[:：]?\s*")

    def _base(self, doc, p: bool, r: bool):  # noqa C901
        """
        校验中文关键词格式：
        - 段落整体格式（对齐、行距等）
        - "关键词："部分加粗，其余内容不加粗
        - 校验关键词数量、末尾标点
        """
        # 1. 空值校验
        if self.pydantic_config is None:
            self.add_comment(
                doc=doc, runs=self.paragraph.runs, text="中文关键词配置未加载，跳过检查"
            )
            return [{"error": "配置未加载", "lang": "cn", "node_type": self.NODE_TYPE}]

        cfg = self.pydantic_config

        # 2. 段落样式检查（复用基类方法）
        paragraph_issues = self._check_paragraph_style(cfg, p)
        if paragraph_issues:
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=paragraph_issues,
            )

        # 3. 拆分标签和内容混合的 run（仅在格式化模式下执行）
        if not p:
            self._split_mixed_runs()

        # 4. 字符样式检查（区分标签/内容）
        # (a) "关键词：" 标签 —— 加粗（修复拼写错误：kewords_bold → kewords_bold）
        label_style = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.keywords_bold,  # 原代码拼写错误：kewords → keywords
            italic=cfg.italic,
            underline=cfg.underline,
        )
        # (b) 关键词内容 —— 不加粗
        content_style = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        # 4. 遍历run校验
        for run in self.paragraph.runs:
            if not run.text.strip():
                continue

            if self._check_keyword_label(run):
                # 检查标签样式
                if r:
                    diff = label_style.diff_from_run(run)
                else:
                    diff = label_style.apply_to_run(run)
                if diff:
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=CharacterStyle.to_string(diff),
                    )
            else:
                # 检查内容样式
                if r:
                    diff = content_style.diff_from_run(run)
                else:
                    diff = content_style.apply_to_run(run)
                if diff:
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=CharacterStyle.to_string(diff),
                    )

        # 5. 专属校验：关键词数量 + 末尾标点（KeywordsConfig特有配置）
        keyword_text = "".join([run.text for run in self.paragraph.runs])
        # 提取中文关键词（按分号分割）
        keyword_list = re.split(r"；", re.sub(r"关键词[:：]", "", keyword_text))
        keyword_list = [k.strip() for k in keyword_list if k.strip()]

        # 数量校验
        if len(keyword_list) < cfg.count_min:
            issue = f"中文关键词数量不足（最少{cfg.count_min}个，当前{len(keyword_list)}个）"
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)
        if len(keyword_list) > cfg.count_max:
            issue = f"中文关键词数量超限（最多{cfg.count_max}个，当前{len(keyword_list)}个）"
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)

        # 末尾标点校验
        if (
            cfg.trailing_punct_forbidden
            and keyword_text.strip()
            and keyword_text.strip()[-1] in "；，。、"
        ):
            issue = "中文关键词末尾禁止出现标点符号"
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)
