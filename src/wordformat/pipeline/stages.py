#! /usr/bin/env python
# @Time    : 2026/7/7 12:25
# @Author  : afish
# @File    : loadpipline.py
# 加载配置、文档
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentObject
from docx.shared import Pt, RGBColor

from wordformat.config.loader import load_config
from wordformat.config.models import NodeConfigRoot
from wordformat.log_config import logger
from wordformat.rules.node import FormatNode
from wordformat.settings import VOIDNODELIST
from wordformat.structure.document_builder import DocumentBuilder
from wordformat.style.defs import (
    Alignment,
    FirstLineIndent,
    FontColor,
    FontSize,
    LeftIndent,
    LineSpacing,
    LineSpacingRule,
    RightIndent,
    SpaceAfter,
    SpaceBefore,
    ensure_style_exists,
)
from wordformat.style.writer import (
    SetFirstLineIndent,
    SetIndent,
    SetSpacing,
)
from wordformat.style.xml_ops import (
    ensure_pPr,
)
from wordformat.utils import (
    ensure_directory_exists,
    get_file_name,
)

from .context import FormatContext


class LoadConfigStage:
    """加载配置pipline"""

    def process(self, ctx: FormatContext) -> FormatContext:
        if ctx.config_path:
            try:
                ctx.config_model = load_config(ctx.config_path)
                logger.info("配置文件验证通过")
            except Exception as e:
                logger.error(f"配置加载失败: {str(e)}")
                raise
        else:
            ctx.config_model = NodeConfigRoot()
            logger.info("未提供配置文件，使用默认配置")

        # 预设 manifest 深合并：用户配置最优先，后加载预设覆盖先加载预设
        if ctx.preset_manifests:
            from wordformat.config.dotdict import deep_merge

            merged = dict(ctx.config_model)
            for manifest in ctx.preset_manifests:
                merged = deep_merge(manifest, merged)

            ctx.config_model = NodeConfigRoot(**merged)

        # 配置加载完成 hook：预设可修改/补充配置
        from wordformat.hooks import hooks

        result = hooks.emit("on_config_loaded", config=ctx.config_model)
        ctx.config_model = result.get("config", ctx.config_model)

        # 结构类别策略：YAML `structure` 段覆盖默认值（哪些类别格式化/隔离子树）
        from wordformat.settings import resolve_structure_categories

        resolved = resolve_structure_categories(ctx.config_model)
        ctx.void_categories = resolved["void"]
        ctx.skip_children_categories = resolved["skip_children"]
        return ctx


class LoadDocxStage:
    """加载docx的pipline"""

    def process(self, ctx: FormatContext) -> FormatContext:
        """加载docx"""
        ctx.document = Document(ctx.docx_path)
        return ctx


class TreeBuildingStage:
    """构建tree pipline。
    优先使用 ctx.paragraphs（内存中的段落列表，如 MD 解析结果），
    否则从 ctx.json_path 加载 JSON 文件。
    """

    def __init__(self, registry=None):
        """registry: 领域注册表实例（get_registry(domain)），None = 默认论文域。"""
        self._registry = registry

    def process(self, ctx: FormatContext) -> FormatContext:
        source = ctx.paragraphs if ctx.paragraphs else ctx.json_path
        ctx.root_node = DocumentBuilder.build_from_json(
            source, config=ctx.config_model, registry=self._registry
        )
        # 文档树构建完成 hook：预设可自定义节点分类
        from wordformat.hooks import hooks

        result = hooks.emit(
            "on_tree_built", root_node=ctx.root_node, config=ctx.config_model
        )
        ctx.root_node = result.get("root_node", ctx.root_node)
        return ctx


class ParagraphAlignmentStage:
    """展平树并对齐段落 pipline"""

    @staticmethod
    def _is_detached_table(node) -> bool:
        """原文档已有表格（exists）或 sync 忽略的表格不占用段落槽位。"""
        value = getattr(node, "value", None)
        if not isinstance(value, dict) or value.get("category") != "table_object":
            return False
        return bool(value.get("exists") or value.get("sync_skip"))

    def process(self, ctx: FormatContext) -> FormatContext:
        nodes = self._flatten_tree_nodes(ctx.root_node)
        # 表格节点（exists/sync_skip）不绑定段落，因此不参与 zip 计数
        paragraph_nodes = [n for n in nodes if not self._is_detached_table(n)]
        paragraphs = ctx.document.paragraphs
        if len(paragraph_nodes) != len(paragraphs):
            # 节点数不等于段落数时，zip 会静默错位：后面的节点整体前移，
            # 导致 keywords_chinese/caption_figure/heading_level_* 等规则
            # 被应用到错误段落（如正文被加“图1.1”前缀）。必须显式报错。
            # 常见诱因：生成节点后更换了文档；或前端页面版本过旧
            # （旧版会过滤 figure_image 占位节点）。
            raise ValueError(
                f"当前文档({len(paragraphs)} 段)与节点 JSON({len(paragraph_nodes)} 节点)不一致，"
                "两者必须一一对应。可能是生成节点后又换过文档，"
                "或页面版本过旧（旧版会过滤占位节点）。"
                "请刷新页面后重新上传文档并点击「生成节点JSON」，再执行格式化。"
            )
        for node, para in zip(paragraph_nodes, paragraphs, strict=True):
            node.paragraph = para
        return ctx

    def _flatten_tree_nodes(self, root_node):
        """DFS 前序遍历展平树中所有节点（排除虚拟根节点），顺序与文档段落一致。"""
        result = []

        def dfs(node):
            for child in node.children:
                result.append(child)
                dfs(child)

        dfs(root_node)
        return result


class TreeNormalizationStage:
    """提升子树（摘要、参考文献）—— 机制由论文领域 handler 承担。"""

    def process(self, ctx: FormatContext) -> FormatContext:
        from wordformat.hooks import hooks

        hooks.emit(
            "on_tree_normalize",
            root_node=ctx.root_node,
            config=ctx.config_model,
            check=ctx.check,
        )
        return ctx


class StyleDefinitionFixStage:
    """修正样式定义（仅 apply 模式）"""

    def _fix_style_run_properties(self, style, cfg, style_name: str):
        """修正样式定义中的字符格式属性。

        python-docx 的 Style.font API 覆盖 size/color/bold/italic/underline，
        但 eastAsia 字体名需通过 XML 设置（Font 类不支持该属性）。
        """
        # 字体名：西文用 style.font.name，东亚用 XML（python-docx 不支持 eastAsia）
        cn_name = getattr(cfg, "chinese_font_name", None)
        en_name = getattr(cfg, "english_font_name", None)
        if cn_name or en_name:
            from wordformat.style.xml_ops import ensure_rPr, rPr_set_font

            rPr = ensure_rPr(style.element)
            rPr_set_font(rPr, cn_name=cn_name, en_name=en_name)

        font_size = getattr(cfg, "font_size", None)
        if font_size is not None:
            try:
                style.font.size = Pt(FontSize(font_size).rel_value)
            except Exception as e:
                logger.warning(f"设置样式 '{style_name}' 字号失败: {e}")

        font_color = getattr(cfg, "font_color", None)
        if font_color is not None:
            try:
                rgb = FontColor(font_color).rel_value
                style.font.color.rgb = RGBColor(*rgb)
            except Exception as e:
                logger.warning(f"设置样式 '{style_name}' 颜色失败: {e}")

        bold = getattr(cfg, "bold", None)
        if bold is not None:
            style.font.bold = bold

        italic = getattr(cfg, "italic", None)
        if italic is not None:
            style.font.italic = italic

        underline = getattr(cfg, "underline", None)
        if underline is not None:
            style.font.underline = underline

    def _fix_style_paragraph_properties(self, style, cfg, style_name: str):
        """修正样式定义中的段落格式属性。

        优先使用 python-docx 的 style.paragraph_format API（对齐、行距）。
        行单位（段前/段后间距）和字符单位（缩进）因 python-docx 不支持，
        回退到 XML 操作。
        """
        # --- 对齐方式（python-docx API） ---
        alignment = getattr(cfg, "alignment", None)
        if alignment is not None:
            try:
                style.paragraph_format.alignment = Alignment(alignment).rel_value
            except Exception as e:
                logger.warning(f"设置样式 '{style_name}' 对齐方式失败: {e}")

        # --- 行距（python-docx API） ---
        line_spacingrule = getattr(cfg, "line_spacingrule", None)
        if line_spacingrule is not None:
            try:
                lsr = LineSpacingRule(line_spacingrule)
                style.paragraph_format.line_spacing_rule = lsr.rel_value
                line_spacing = getattr(cfg, "line_spacing", None)
                if line_spacing is not None:
                    ls = LineSpacing(line_spacing)
                    if ls.rel_unit == "pt":
                        style.paragraph_format.line_spacing = Pt(ls.rel_value)
                    else:
                        style.paragraph_format.line_spacing = ls.rel_value
                else:
                    logger.warning(
                        f"样式 '{style_name}' 设置了 line_spacingrule 但未设置 line_spacing，已跳过行距"
                    )
            except Exception as e:
                logger.warning(f"设置样式 '{style_name}' 行距失败: {e}")

        # --- 段前/段后间距（仅支持行单位，需 XML） ---
        pPr = ensure_pPr(style.element)
        for attr_name, cls, spacing_type in [
            ("space_before", SpaceBefore, "before"),
            ("space_after", SpaceAfter, "after"),
        ]:
            val = getattr(cfg, attr_name, None)
            if val is None:
                continue
            try:
                inst = cls(val)
                if inst.rel_unit == "hang":
                    SetSpacing._set_hang_on_pPr(pPr, spacing_type, inst.rel_value)
                else:
                    logger.warning(
                        f"样式 '{style_name}' {attr_name} 使用了 '{inst.rel_unit}' 单位，"
                        f"样式定义仅支持'行'单位，已跳过"
                    )
            except Exception as e:
                logger.warning(f"设置样式 '{style_name}' {attr_name} 失败: {e}")

        # --- 缩进（仅支持字符单位，需 XML） ---
        first_line_indent = getattr(cfg, "first_line_indent", None)
        if first_line_indent is not None:
            try:
                inst = FirstLineIndent(first_line_indent)
                if inst.rel_unit == "char":
                    SetFirstLineIndent._clear_ind_on_pPr(pPr)
                    SetFirstLineIndent._set_char_on_pPr(pPr, inst.rel_value)
                else:
                    logger.warning(
                        f"样式 '{style_name}' first_line_indent 使用了 '{inst.rel_unit}' 单位，"
                        f"样式定义仅支持'字符'单位，已跳过"
                    )
            except Exception as e:
                logger.warning(f"设置样式 '{style_name}' first_line_indent 失败: {e}")

        for attr_name, cls, indent_type in [
            ("left_indent", LeftIndent, "R"),
            ("right_indent", RightIndent, "X"),
        ]:
            val = getattr(cfg, attr_name, None)
            if val is None:
                continue
            try:
                inst = cls(val)
                if inst.rel_unit == "char":
                    SetIndent._set_char_on_pPr(pPr, indent_type, inst.rel_value)
                else:
                    logger.warning(
                        f"样式 '{style_name}' {attr_name} 使用了 '{inst.rel_unit}' 单位，"
                        f"样式定义仅支持'字符'单位，已跳过"
                    )
            except Exception as e:
                logger.warning(f"设置样式 '{style_name}' {attr_name} 失败: {e}")

    def _fix_all_style_definitions(self, document: DocumentObject, config_model):
        """在格式化开始前，统一修正文档中所有使用的样式定义。

        遍历配置中所有段（body_text、headings、abstract、figures、tables、
        references、acknowledgements 等），收集唯一的 builtin_style_name
        及其在所有节点间一致的格式字段，然后修正样式定义。

        修正内容：
        1. 字符格式：中英文字体、字号、颜色（清除主题色）、加粗、斜体、下划线
        2. 段落格式：对齐方式
        3. 确保样式定义存在（不存在则创建）

        注意：
        - 同名样式被多个节点复用时，只应用取值一致的字段，冲突字段由
          各节点的直接格式保证，避免把某一节点的缩进/字体泄漏给其他段落。
        - Normal 是表格单元格、页眉页脚、题注等的公共基样式，正文/参考
          文献/致谢对它的配置通常互相冲突；写入任一份都会污染这些未显式
          覆盖的段落，因此不修改 Normal 定义本身。
        """

        style_configs = config_model.collect_style_configs()

        for eng_style_name, cfg in style_configs.items():
            if eng_style_name == "Normal":
                logger.debug("跳过共享基样式定义修正: Normal")
                continue

            ensure_style_exists(document, eng_style_name)

            try:
                style = document.styles[eng_style_name]
            except KeyError:
                logger.warning(f"样式 '{eng_style_name}' 创建失败，跳过修正")
                continue

            # collect_style_configs 返回普通 dict；转成支持点号访问的容器，
            # 否则 getattr(cfg, ...) 全部取到 None，修正会静默失效。
            cfg_view = NodeConfigRoot()
            cfg_view.update(cfg)
            self._fix_style_run_properties(style, cfg_view, eng_style_name)
            self._fix_style_paragraph_properties(style, cfg_view, eng_style_name)

            logger.debug(f"已修正样式定义: {eng_style_name}")

    def process(self, ctx: FormatContext) -> FormatContext:
        if not ctx.check:
            self._fix_all_style_definitions(ctx.document, ctx.config_model)
        return ctx


class FormattingExecutionStage:
    """执行格式化/检查（核心遍历）"""

    def __init__(self, skip_comments: bool = False):
        self.skip_comments = skip_comments

    def apply_format_check_to_all_nodes(
        self,
        root_node: FormatNode,
        document,
        config,
        check=True,
        void_categories=None,
        skip_children_categories=None,
    ):
        """
        递归遍历文档树中的所有节点，
        对每个具有 check_format 方法的节点执行该方法。

        :param root_node: 树的根节点（FormatNode 或其子类实例）
        :param document: docx文档的实例
        :param config: 配置文件
        :param check: 用来控制是仅检查还是仅修改
        :param void_categories: 不参与格式化的类别集合（None = 内置 VOIDNODELIST；[] = 不排除任何类别）
        :param skip_children_categories: 子树不遍历的类别集合（None = 内置默认；[] = 全部遍历）
        """
        from wordformat.hooks import hooks
        from wordformat.settings import DEFAULT_SKIP_CHILDREN_CATEGORIES

        void_set = set(VOIDNODELIST if void_categories is None else void_categories)
        skip_children_set = set(
            DEFAULT_SKIP_CHILDREN_CATEGORIES
            if skip_children_categories is None
            else skip_children_categories
        )

        def traverse(node, parent_category=""):
            category = (
                node.value.get("category", "") if isinstance(node.value, dict) else ""
            )

            if hasattr(node, "check_format"):
                try:
                    # top 节点直接关联的 body_text 不参与格式化（如封面页、原创性声明等）
                    # 但间接关联的 body_text（作为 heading 子节点）正常格式化
                    is_top_direct_body_text = (
                        parent_category == "top" and category == "body_text"
                    )
                    if category not in void_set and not is_top_direct_body_text:
                        node.load_config(config)

                        if node.paragraph:
                            # 先执行内容替换（check/format 两种模式均执行）
                            node.apply_replace(document)
                            # 节点格式化 hook：领域（题注编号注入）与预设都在此处理
                            hooks.emit(
                                "before_node_format",
                                node=node,
                                paragraph=node.paragraph,
                                config=config,
                                check=check,
                            )
                            if check:
                                node.check_format(document)
                            elif self.skip_comments:
                                node.apply_style(document)
                            else:
                                node.apply_format(document)
                            hooks.emit(
                                "after_node_format",
                                node=node,
                                paragraph=node.paragraph,
                                config=config,
                                check=check,
                            )
                except Exception as e:
                    logger.warning(f"Node {node} not format, because: {str(e)}")
                    raise e

            # 目录、附录、封面/声明等 skip_children 类别的子节点跳过格式化
            if category not in skip_children_set:
                for child in node.children:
                    traverse(child, parent_category=category)

        traverse(root_node)

    def process(self, ctx: FormatContext) -> FormatContext:
        if ctx.check:
            FormatNode.reset_stats()
        # 格式化开始 hook：领域 handler 在此重置自己的遍历状态（如论文题注编号）
        from wordformat.hooks import hooks

        hooks.emit(
            "on_format_begin",
            root_node=ctx.root_node,
            document=ctx.document,
            config=ctx.config_model,
            check=ctx.check,
        )
        self.apply_format_check_to_all_nodes(
            ctx.root_node,
            ctx.document,
            ctx.config_model,
            ctx.check,
            void_categories=ctx.void_categories,
            skip_children_categories=ctx.skip_children_categories,
        )
        return ctx


class SummaryGenerationStage:
    """生成检测报告摘要（仅 check 模式）—— 摘要文本由 on_summary_build 内置 handler 产出。"""

    def _add_summary_comment(self, document, summary: str) -> None:
        """将检测报告摘要作为批注添加到文档第一段。空段临时塞空 run 做锚点。"""
        para = document.paragraphs[0]
        if not para.runs:
            para.add_run("")
        document.add_comment(
            runs=para.runs, text=summary, author="Wordformat", initials="afish"
        )

    def process(self, ctx: FormatContext) -> FormatContext:
        if ctx.check:
            from wordformat.hooks import hooks

            result = hooks.emit(
                "on_summary_build",
                root_node=ctx.root_node,
                document=ctx.document,
                config=ctx.config_model,
                check=ctx.check,
            )
            summary = result.get("summary")
            if summary:
                self._add_summary_comment(ctx.document, summary)
        return ctx


class DocumentSavingStage:
    """保存文档"""

    def process(self, ctx: FormatContext) -> FormatContext:
        # 保存前 hook：机制（标题编号/超链接、页眉页脚修正等）及预设都在此触发
        from wordformat.hooks import hooks

        hooks.emit(
            "before_document_save",
            document=ctx.document,
            config=ctx.config_model,
            ctx=ctx,
            check=ctx.check,
        )

        ensure_directory_exists(ctx.save_dir)
        filename = get_file_name(ctx.docx_path)
        suffix = "--标注版.docx" if ctx.check else "--修改版.docx"
        out_path = Path(ctx.save_dir) / f"{filename}{suffix}"

        ctx.document.save(str(out_path))
        logger.info(f"保存文件到 {out_path}")
        ctx.output_path = str(out_path)
        return ctx
