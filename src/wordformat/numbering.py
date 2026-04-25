#! /usr/bin/env python
"""
标题自动编号模块

功能：
1. 清除标题段落中的手动编号文本（正则匹配）
2. 应用 Word 自动编号（通过 XML 操作）

流程：
  格式化完成后 → 对 heading 节点 → 清除手动编号 → 应用自动编号
"""

import re
from copy import deepcopy
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from loguru import logger


def strip_manual_numbering(paragraph, pattern: str) -> bool:
    """
    清除段落开头的手动编号文本。

    仅修改第一个 run 的文本，保留其余 run 的格式不变。
    如果段落文本不匹配正则，不做任何修改。

    Args:
        paragraph: docx 段落对象
        pattern: 正则表达式，用于匹配段落开头的编号

    Returns:
        bool: 是否成功清除了编号
    """
    if not pattern or not paragraph.runs:
        return False

    full_text = paragraph.text
    match = re.match(pattern, full_text)
    if not match:
        return False

    # 计算需要删除的字符数
    stripped_len = match.end()

    # 从第一个 run 开始逐字符删除
    remaining = stripped_len
    for run in paragraph.runs:
        if remaining <= 0:
            break
        run_text = run.text
        if len(run_text) <= remaining:
            # 整个 run 都要删除
            remaining -= len(run_text)
            run.text = ""
        else:
            # 只删除 run 的前 remaining 个字符
            run.text = run_text[remaining:]
            remaining = 0

    # 清除第一个 run 开头可能残留的空白
    if paragraph.runs and paragraph.runs[0].text:
        paragraph.runs[0].text = paragraph.runs[0].text.lstrip()

    logger.debug(f"已清除手动编号: '{match.group()}' → 段落剩余: '{paragraph.text[:30]}...'")
    return True


def apply_auto_numbering(paragraph, num_id: str, ilvl: str = "0"):
    """
    为段落应用 Word 自动编号。

    通过在段落的 <w:pPr> 中添加 <w:numPr> 元素，
    引用文档中已有的 numbering 定义。

    Args:
        paragraph: docx 段落对象
        num_id: 编号定义 ID（对应 numbering.xml 中的 w:numId）
        ilvl: 编号级别（默认 "0"）
    """
    pPr = paragraph._element.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        paragraph._element.insert(0, pPr)

    # 移除已有的 numPr（避免重复）
    existing_numPr = pPr.find(qn("w:numPr"))
    if existing_numPr is not None:
        pPr.remove(existing_numPr)

    # 创建新的 numPr
    numPr = OxmlElement("w:numPr")
    ilvl_elem = OxmlElement("w:ilvl")
    ilvl_elem.set(qn("w:val"), ilvl)
    numId_elem = OxmlElement("w:numId")
    numId_elem.set(qn("w:val"), num_id)
    numPr.append(ilvl_elem)
    numPr.append(numId_elem)
    pPr.append(numPr)

    logger.debug(f"已应用自动编号: numId={num_id}, ilvl={ilvl}")


def create_numbering_definition(document, config) -> dict[str, str]:
    """
    在文档中创建自动编号定义（如果不存在）。

    根据配置中的 template 生成 abstractNum 和 num 定义，
    返回 {level_key: num_id} 映射。

    Args:
        document: docx Document 对象
        config: NumberingConfig 配置对象

    Returns:
        dict: {"level_1": "100", "level_2": "101", "level_3": "102"}
    """
    if not config.enabled:
        return {}

    # 获取或创建 numbering part
    try:
        numbering_part = document.part.numbering_part
    except (AttributeError, KeyError, NotImplementedError):
        # 文档没有 numbering part，需要手动创建
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        from docx.opc.packuri import PackURI
        from docx.parts.numbering import NumberingPart

        numbering_elm = OxmlElement("w:numbering")

        # 手动创建 NumberingPart 实例并建立关系
        numbering_part = NumberingPart(
            PackURI("/word/numbering.xml"),
            "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            numbering_elm,
            document.part.package,
        )
        document.part.relate_to(numbering_part, RT.NUMBERING)

    numbering_elm = numbering_part._element

    # 查找已有的最大 abstractNumId 和 numId
    max_abstract_num_id = -1
    max_num_id = 0
    for elem in numbering_elm.findall(qn("w:abstractNum")):
        num_id = int(elem.get(qn("w:abstractNumId"), "0"))
        max_abstract_num_id = max(max_abstract_num_id, num_id)
    for elem in numbering_elm.findall(qn("w:num")):
        num_id = int(elem.get(qn("w:numId"), "0"))
        max_num_id = max(max_num_id, num_id)

    result = {}
    level_configs = [
        ("level_1", config.level_1, 0),
        ("level_2", config.level_2, 1),
        ("level_3", config.level_3, 2),
    ]

    for level_key, level_config, ilvl in level_configs:
        if not level_config.enabled or not level_config.template:
            continue

        abstract_num_id = max_abstract_num_id + 1
        max_abstract_num_id += 1
        num_id = max_num_id + 1
        max_num_id += 1

        # 创建 abstractNum
        abstract_num = OxmlElement("w:abstractNum")
        abstract_num.set(qn("w:abstractNumId"), str(abstract_num_id))

        # 创建多级 lvl（需要创建所有上级级别才能正确计数）
        for lvl_i in range(ilvl + 1):
            lvl = OxmlElement("w:lvl")
            lvl.set(qn("w:ilvl"), str(lvl_i))

            start = OxmlElement("w:start")
            start.set(qn("w:val"), "1")
            lvl.append(start)

            numFmt = OxmlElement("w:numFmt")
            # 根据模板判断格式
            template = level_config.template
            if "第" in template and "章" in template:
                numFmt.set(qn("w:val"), "chineseCountingThousand")
            else:
                numFmt.set(qn("w:val"), "decimal")
            lvl.append(numFmt)

            lvlText = OxmlElement("w:lvlText")
            # 根据级别生成 lvlText
            if lvl_i == ilvl:
                lvlText.set(qn("w:val"), level_config.template)
            else:
                # 上级级别使用简单的 %N 格式
                lvlText.set(qn("w:val"), f"%{lvl_i + 1}")
            lvl.append(lvlText)

            lvlJc = OxmlElement("w:lvlJc")
            lvlJc.set(qn("w:val"), "left")
            lvl.append(lvlJc)

            # pPr
            pPr = OxmlElement("w:pPr")
            ind = OxmlElement("w:ind")
            ind.set(qn("w:left"), "0")
            ind.set(qn("w:hanging"), str((lvl_i + 1) * 420))  # 每级缩进约 0.3cm
            pPr.append(ind)
            lvl.append(pPr)

            abstract_num.append(lvl)

        numbering_elm.append(abstract_num)

        # 创建 num 引用
        num = OxmlElement("w:num")
        num.set(qn("w:numId"), str(num_id))
        abstract_num_id_ref = OxmlElement("w:abstractNumId")
        abstract_num_id_ref.set(qn("w:val"), str(abstract_num_id))
        num.append(abstract_num_id_ref)
        numbering_elm.append(num)

        result[level_key] = str(num_id)
        logger.debug(f"创建编号定义: {level_key} → numId={num_id}, template={level_config.template}")

    return result


def process_heading_numbering(root_node, document, config):
    """
    处理所有标题节点的编号：清除手动编号 + 应用自动编号。

    Args:
        root_node: 文档树根节点
        document: docx Document 对象
        config: NumberingConfig 配置对象
    """
    if not config.enabled:
        return

    # 创建编号定义
    num_id_map = create_numbering_definition(document, config)
    if not num_id_map:
        return

    # 标题节点类型到配置级别的映射
    heading_map = {
        "level_1": ("level_1", "0"),
        "level_2": ("level_2", "1"),
        "level_3": ("level_3", "2"),
    }

    def traverse(node):
        """递归遍历文档树"""
        # 获取节点的 category
        category = node.value.get("category", "") if isinstance(node.value, dict) else ""

        for level_key, (config_key, ilvl) in heading_map.items():
            if category == f"heading_{level_key}":
                level_config = getattr(config, config_key, None)
                if not level_config or not level_config.enabled:
                    break

                paragraph = getattr(node, "paragraph", None)
                if not paragraph:
                    break

                # 1. 清除手动编号
                if level_config.strip_pattern:
                    strip_manual_numbering(paragraph, level_config.strip_pattern)

                # 2. 应用自动编号
                num_id = num_id_map.get(config_key)
                if num_id:
                    apply_auto_numbering(paragraph, num_id, ilvl)

                break

        for child in node.children:
            traverse(child)

    traverse(root_node)
    logger.info("标题自动编号处理完成")
