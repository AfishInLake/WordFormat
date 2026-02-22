#! /usr/bin/env python
# @Time    : 2026/1/11 19:38
# @Author  : afish
# @File    : heading.py

from docx.document import Document
from loguru import logger  # 推荐添加日志，便于调试

from wordformat.config.datamodel import HeadingLevelConfig, NodeConfigRoot
from wordformat.rules.node import FormatNode
from wordformat.style.check_format import CharacterStyle, ParagraphStyle


class BaseHeadingNode(FormatNode[HeadingLevelConfig]):
    """标题节点基类（复用1/2/3级标题的通用逻辑）"""

    LEVEL: int = 0  # 标题层级（1/2/3）
    NODE_TYPE: str = ""
    CONFIG_MODEL = HeadingLevelConfig

    def _get_level_config(self, root_config: NodeConfigRoot) -> HeadingLevelConfig:
        """根据层级获取对应配置"""
        level_config_map = {
            1: root_config.headings.level_1,
            2: root_config.headings.level_2,
            3: root_config.headings.level_3,
        }
        target_config = level_config_map.get(self.LEVEL, root_config.headings.level_1)
        return target_config

    def load_config(self, root_config: dict | NodeConfigRoot):
        """重载加载配置方法，自动匹配对应层级的配置"""
        try:
            if isinstance(root_config, dict):
                # 修复：使用单下划线 _config（匹配基类@property的底层属性）
                level_config_dict = root_config.get("headings", {}).get(
                    f"level_{self.LEVEL}", {}
                )
                self._config = level_config_dict  # 正确赋值给单下划线私有属性
                logger.debug(f"{self.LEVEL}级标题字典配置：{self._config}")
                self._pydantic_config = self.CONFIG_MODEL(
                    **self._config
                )  # 读取赋值后的 _config

            elif isinstance(root_config, NodeConfigRoot):
                # 修复：先赋值 _pydantic_config，再同步到 _config
                self._pydantic_config = self._get_level_config(root_config)
                self._config = (
                    self._pydantic_config.model_dump()
                )  # 读取底层 _pydantic_config
            else:
                raise TypeError(
                    f"配置类型不支持：{type(root_config)}，仅支持dict或NodeConfigRoot"
                )

        except Exception as e:
            logger.error(f"{self.LEVEL}级标题配置加载失败：{str(e)}")
            raise  # 抛出异常，避免后续使用错误配置

    def _base(self, doc: Document, p: bool, r: bool):
        """通用标题格式检查逻辑"""
        # 修复：空值校验（直接检查底层私有属性 _pydantic_config）
        if self._pydantic_config is None:
            error_msg = f"{self.LEVEL}级标题配置未加载，跳过检查"
            self.add_comment(doc=doc, runs=self.paragraph.runs, text=error_msg)
            logger.warning(error_msg)
            return [
                {"error": error_msg, "level": self.LEVEL, "node_type": self.NODE_TYPE}
            ]

        # 类型断言（读取@property的 pydantic_config，本质是 _pydantic_config）
        cfg = self.pydantic_config
        # 段落样式检查（完整使用所有配置字段）
        ps = ParagraphStyle.from_config(cfg)
        if p:
            paragraph_issues = ps.diff_from_paragraph(self.paragraph)
        else:
            paragraph_issues = ps.apply_to_paragraph(self.paragraph)

        # 字符样式检查（完整使用所有配置字段）
        cstyle = CharacterStyle(
            font_name_cn=cfg.chinese_font_name,
            font_name_en=cfg.english_font_name,
            font_size=cfg.font_size,
            font_color=cfg.font_color,
            bold=cfg.bold,
            italic=cfg.italic,
            underline=cfg.underline,
        )

        # 逐run检查并添加批注
        run_issues = []
        for idx, run in enumerate(self.paragraph.runs):
            if not run.text.strip():
                continue  # 跳过空白run，减少无效检查
            if r:
                diff_result = cstyle.diff_from_run(run)
            else:
                diff_result = cstyle.apply_to_run(run)
            if diff_result:
                run_issue = {
                    "run_index": idx,
                    "run_text": run.text,
                    "diff": diff_result,
                }
                run_issues.append(run_issue)
                self.add_comment(
                    doc=doc, runs=run, text=CharacterStyle.to_string(diff_result)
                )

        # 段落样式批注
        all_issues = []
        if paragraph_issues:
            para_issue = {
                "type": "paragraph_style",
                "diff": paragraph_issues,
                "paragraph_text": self.paragraph.text[:50] + "..."
                if len(self.paragraph.text) > 50
                else self.paragraph.text,
            }
            all_issues.append(para_issue)
            self.add_comment(
                doc=doc,
                runs=self.paragraph.runs,
                text=ParagraphStyle.to_string(paragraph_issues),
            )

        # 添加字符样式问题
        if run_issues:
            all_issues.extend(run_issues)

        return all_issues


# 各层级标题节点（无需重写check_format，直接复用基类逻辑）
class HeadingLevel1Node(BaseHeadingNode):
    """一级标题节点"""

    LEVEL = 1
    NODE_TYPE = "headings.level_1"


class HeadingLevel2Node(BaseHeadingNode):
    """二级标题节点"""

    LEVEL = 2
    NODE_TYPE = "headings.level_2"


class HeadingLevel3Node(BaseHeadingNode):
    """三级标题节点"""

    LEVEL = 3
    NODE_TYPE = "headings.level_3"
