import re
from typing import List, Literal

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
            # 从字典中提取对应语言的配置
            self.__config = (
                root_config.get("abstract", {}).get("keywords", {}).get(self.LANG, {})
            )
            self._pydantic_config = self.CONFIG_MODEL(**self.config)
        elif isinstance(root_config, NodeConfigRoot):
            # 从Pydantic模型提取对应语言的配置
            self._pydantic_config = self._get_lang_config(root_config)
            self.__config = self._pydantic_config.model_dump()
        else:
            raise TypeError(
                f"配置类型不支持：{type(root_config)}，仅支持dict或NodeConfigRoot"
            )

    def _check_paragraph_style(self, cfg: KeywordsConfig, p: bool) -> List[str]:
        """通用段落样式检查（复用）"""
        ps = ParagraphStyle(
            alignment=cfg.alignment,
            space_before=cfg.space_before,
            space_after=cfg.space_after,
            line_spacing=cfg.line_spacing,
            line_spacingrule=cfg.line_spacingrule,
            first_line_indent=cfg.first_line_indent,
            builtin_style_name=cfg.builtin_style_name,
        )
        if p:
            return [o.comment for o in ps.diff_from_paragraph(self.paragraph)]
        else:
            return [o.comment for o in ps.apply_to_paragraph(self.paragraph)]


# 第二步：英文关键词节点（专属逻辑）
class KeywordsEN(BaseKeywordsNode):
    """关键词节点-英文"""

    LANG = "en"
    NODE_TYPE = "abstract.keywords.english"

    def _check_keyword_label(self, run) -> bool:
        """检查run是否包含英文关键词标签（Keywords/KEY WORDS）"""
        pattern = r"Keywords?|KEY\s*WORDS"
        return bool(re.search(pattern, run.text, re.IGNORECASE))

    def _base(self, doc, p: bool, r: bool):  # noqa C901
        """
        校验英文关键词格式：
        - 段落整体格式（对齐、行距等）
        - “Keywords:” 部分加粗，其余内容不加粗
        """
        # 1. 空值校验
        if self.pydantic_config is None:
            self.add_comment(
                doc=doc, runs=self.paragraph.runs, text="英文关键词配置未加载，跳过检查"
            )
            return [{"error": "配置未加载", "lang": "en", "node_type": self.NODE_TYPE}]

        cfg = self.pydantic_config
        all_issues = []

        # 2. 段落样式检查
        paragraph_issues = self._check_paragraph_style(cfg, p)
        if paragraph_issues:
            all_issues.extend(paragraph_issues)
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=f"英文关键词段落格式错误：{''.join(str(issue) for issue in paragraph_issues)}",
            )

        # 3. 字符样式检查（区分标签/内容）
        # (a) "Keywords:" 标签 —— 加粗
        label_style = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.kewords_bold,  # 修复拼写错误：kewords → keywords
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
                    all_issues.append(
                        {"run_text": run.text, "diff": diff, "type": "label"}
                    )
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=f"英文关键词标签格式错误：{''.join(str(d) for d in diff)}",
                    )
            else:
                # 检查内容样式
                if r:
                    diff = content_style.diff_from_run(run)
                else:
                    diff = content_style.apply_to_run(run)
                if diff:
                    all_issues.append(
                        {"run_text": run.text, "diff": diff, "type": "content"}
                    )
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=f"英文关键词内容格式错误：{''.join(str(d) for d in diff)}",
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
            all_issues.append({"type": "count", "message": issue})
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)
        if len(keyword_list) > cfg.count_max:
            issue = f"英文关键词数量超限（最多{cfg.count_max}个，当前{len(keyword_list)}个）"
            all_issues.append({"type": "count", "message": issue})
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

    def _base(self, doc, p: bool, r: bool):  # noqa C901
        """
        校验中文关键词格式：
        - 段落整体格式（对齐、行距等）
        - “关键词：”部分加粗，其余内容不加粗
        - 校验关键词数量、末尾标点
        """
        # 1. 空值校验
        if self.pydantic_config is None:
            self.add_comment(
                doc=doc, runs=self.paragraph.runs, text="中文关键词配置未加载，跳过检查"
            )
            return [{"error": "配置未加载", "lang": "cn", "node_type": self.NODE_TYPE}]

        cfg = self.pydantic_config
        all_issues = []

        # 2. 段落样式检查（复用基类方法）
        paragraph_issues = self._check_paragraph_style(cfg, p)
        if paragraph_issues:
            all_issues.extend(paragraph_issues)
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=f"中文关键词段落格式错误：{''.join(str(issue) for issue in paragraph_issues)}",
            )

        # 3. 字符样式检查（区分标签/内容）
        # (a) "关键词：" 标签 —— 加粗（修复拼写错误：kewords_bold → kewords_bold）
        label_style = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.kewords_bold,  # 原代码拼写错误：kewords → keywords
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
                    all_issues.append(
                        {"run_text": run.text, "diff": diff, "type": "label"}
                    )
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=f"中文关键词标签格式错误：{''.join(str(d) for d in diff)}",
                    )
            else:
                # 检查内容样式
                if r:
                    diff = content_style.diff_from_run(run)
                else:
                    diff = content_style.apply_to_run(run)
                if diff:
                    all_issues.append(
                        {"run_text": run.text, "diff": diff, "type": "content"}
                    )
                    self.add_comment(
                        doc=doc,
                        runs=run,
                        text=f"中文关键词内容格式错误：{''.join(str(d) for d in diff)}",
                    )

        # 5. 专属校验：关键词数量 + 末尾标点（KeywordsConfig特有配置）
        keyword_text = "".join([run.text for run in self.paragraph.runs])
        # 提取中文关键词（按分号分割）
        keyword_list = re.split(r"；", re.sub(r"关键词[:：]", "", keyword_text))
        keyword_list = [k.strip() for k in keyword_list if k.strip()]

        # 数量校验
        if len(keyword_list) < cfg.count_min:
            issue = f"中文关键词数量不足（最少{cfg.count_min}个，当前{len(keyword_list)}个）"
            all_issues.append({"type": "count", "message": issue})
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)
        if len(keyword_list) > cfg.count_max:
            issue = f"中文关键词数量超限（最多{cfg.count_max}个，当前{len(keyword_list)}个）"
            all_issues.append({"type": "count", "message": issue})
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)

        # 末尾标点校验
        if (
            cfg.trailing_punct_forbidden
            and keyword_text.strip()
            and keyword_text.strip()[-1] in "；，。、"
        ):
            issue = "中文关键词末尾禁止出现标点符号"
            all_issues.append({"type": "punct", "message": issue})
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=issue)
