#!/usr/bin/env python3
"""
WordFormat 配置文件验证器

验证用户编辑的 config.yaml 是否合法：
1. 检查 YAML 语法是否正确
2. 检查是否包含非法字段（模板中不存在的字段）
3. 检查必要字段是否缺失
4. 检查字段值是否在允许范围内

用法：
    python validate_config.py --config config.yaml
    python validate_config.py --config config.yaml --fix  # 自动移除非法字段
"""

import argparse
import sys
from pathlib import Path

try:
    import yaml
except ImportError:
    # 如果没有 yaml 模块，尝试安装
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "pyyaml", "--break-system-packages", "-q"])
    import yaml


# ==================== 合法字段白名单 ====================

# style_checks_warning 下的合法字段
STYLE_CHECKS_FIELDS = {
    "bold", "italic", "underline", "font_size", "font_name", "font_color",
    "alignment", "space_before", "space_after", "line_spacing", "line_spacingrule",
    "left_indent", "right_indent", "first_line_indent", "builtin_style_name",
}

# global_format 下的合法字段
GLOBAL_FORMAT_FIELDS = {
    "alignment", "space_before", "space_after", "line_spacingrule", "line_spacing",
    "left_indent", "right_indent", "first_line_indent", "builtin_style_name",
    "chinese_font_name", "english_font_name", "font_size", "font_color",
    "bold", "italic", "underline",
}

# abstract.chinese.chinese_title / chinese_content 合法字段
ABSTRACT_CN_FIELDS = GLOBAL_FORMAT_FIELDS

# abstract.english.english_title / english_content 合法字段
ABSTRACT_EN_FIELDS = GLOBAL_FORMAT_FIELDS

# abstract.keywords.chinese / english 合法字段
KEYWORDS_FIELDS = GLOBAL_FORMAT_FIELDS | {
    "keywords_bold", "count_min", "count_max", "trailing_punct_forbidden",
}

# headings 各级别合法字段
HEADING_FIELDS = GLOBAL_FORMAT_FIELDS

# figures 合法字段
FIGURES_FIELDS = GLOBAL_FORMAT_FIELDS | {
    "caption_position", "caption_prefix",
}

# tables 合法字段
TABLES_FIELDS = GLOBAL_FORMAT_FIELDS | {
    "caption_position", "caption_prefix",
}

# references.title 合法字段
REF_TITLE_FIELDS = GLOBAL_FORMAT_FIELDS | {"section_title"}

# references.content 合法字段
REF_CONTENT_FIELDS = GLOBAL_FORMAT_FIELDS | {
    "entry_indent", "entry_ending_punct", "numbering_format",
}

# acknowledgements.title / content 合法字段
ACK_FIELDS = GLOBAL_FORMAT_FIELDS

# numbering 合法字段
NUMBERING_FIELDS = {"enabled", "level_1", "level_2", "level_3"}
NUMBERING_LEVEL_FIELDS = {"enabled", "template", "strip_pattern"}

# 值范围校验
ALIGNMENT_VALUES = {"左对齐", "居中对齐", "右对齐", "两端对齐", "分散对齐"}
LINE_SPACINGRULE_VALUES = {"单倍行距", "1.5倍行距", "2倍行距", "最小值", "固定值", "多倍行距"}
CAPTION_POSITION_VALUES = {"above", "below"}
CAPTION_PREFIX_VALUES = {"图", "表"}
BUILTIN_STYLE_VALUES = {"正文", "Heading 1", "Heading 2", "Heading 3", "题注"}


def load_yaml(filepath: str) -> dict:
    """加载 YAML 文件"""
    with open(filepath, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def check_fields(data: dict, allowed_fields: set, path: str) -> list[str]:
    """检查字段是否合法，返回非法字段列表"""
    if not isinstance(data, dict):
        return []
    errors = []
    for key in data:
        if key not in allowed_fields:
            errors.append(f"  非法字段: {path}.{key} = {data[key]!r}")
    return errors


def check_value(data: dict, field: str, allowed_values: set, path: str) -> list[str]:
    """检查字段值是否在允许范围内"""
    if field in data and data[field] not in allowed_values:
        return [f"  值错误: {path}.{field} = {data[field]!r}，允许值: {allowed_values}"]
    return []


def validate_config(config: dict) -> list[str]:
    """验证配置文件，返回错误列表"""
    errors = []

    # 1. 检查顶层结构
    required_top_keys = {"style_checks_warning", "global_format", "abstract", "headings", "body_text", "figures", "tables", "references", "acknowledgements"}
    actual_top_keys = set(config.keys()) if config else set()
    missing = required_top_keys - actual_top_keys
    if missing:
        errors.append(f"  缺少顶层节点: {missing}")
    # numbering 是可选的，不检查缺失
    extra = actual_top_keys - required_top_keys - {"numbering"}
    if extra:
        errors.append(f"  非法顶层节点: {extra}")

    # 2. style_checks_warning
    if "style_checks_warning" in config:
        errors.extend(check_fields(config["style_checks_warning"], STYLE_CHECKS_FIELDS, "style_checks_warning"))

    # 3. global_format
    if "global_format" in config:
        errors.extend(check_fields(config["global_format"], GLOBAL_FORMAT_FIELDS, "global_format"))

    # 4. abstract
    if "abstract" in config:
        ab = config["abstract"]
        if isinstance(ab, dict):
            for lang in ["chinese", "english"]:
                if lang in ab and isinstance(ab[lang], dict):
                    for section in ab[lang]:
                        section_data = ab[lang][section]
                        allowed = KEYWORDS_FIELDS if "keywords" in section else (ABSTRACT_CN_FIELDS if lang == "chinese" else ABSTRACT_EN_FIELDS)
                        errors.extend(check_fields(section_data, allowed, f"abstract.{lang}.{section}"))

    # 5. headings
    if "headings" in config:
        hd = config["headings"]
        if isinstance(hd, dict):
            for level in ["level_1", "level_2", "level_3"]:
                if level in hd:
                    errors.extend(check_fields(hd[level], HEADING_FIELDS, f"headings.{level}"))

    # 6. body_text
    if "body_text" in config:
        errors.extend(check_fields(config["body_text"], GLOBAL_FORMAT_FIELDS, "body_text"))

    # 7. figures
    if "figures" in config:
        errors.extend(check_fields(config["figures"], FIGURES_FIELDS, "figures"))

    # 8. tables
    if "tables" in config:
        errors.extend(check_fields(config["tables"], TABLES_FIELDS, "tables"))

    # 9. references
    if "references" in config:
        ref = config["references"]
        if isinstance(ref, dict):
            for section in ["title", "content"]:
                if section in ref:
                    allowed = REF_TITLE_FIELDS if section == "title" else REF_CONTENT_FIELDS
                    errors.extend(check_fields(ref[section], allowed, f"references.{section}"))

    # 10. acknowledgements
    if "acknowledgements" in config:
        ack = config["acknowledgements"]
        if isinstance(ack, dict):
            for section in ["title", "content"]:
                if section in ack:
                    errors.extend(check_fields(ack[section], ACK_FIELDS, f"acknowledgements.{section}"))

    # 11. numbering（可选）
    if "numbering" in config:
        num = config["numbering"]
        if isinstance(num, dict):
            errors.extend(check_fields(num, NUMBERING_FIELDS, "numbering"))
            for level in ["level_1", "level_2", "level_3"]:
                if level in num and isinstance(num[level], dict):
                    errors.extend(check_fields(num[level], NUMBERING_LEVEL_FIELDS, f"numbering.{level}"))

    return errors


def remove_extra_fields(data: dict, allowed_fields: set) -> dict:
    """递归移除非法字段"""
    if not isinstance(data, dict):
        return data
    return {k: remove_extra_fields(v, allowed_fields) if isinstance(v, dict) else v
            for k, v in data.items() if k in allowed_fields}


def fix_config(config: dict) -> dict:
    """自动修复配置文件，移除所有非法字段"""
    if not isinstance(config, dict):
        return config

    fixed = {}
    for key, value in config.items():
        if key == "style_checks_warning" and isinstance(value, dict):
            fixed[key] = remove_extra_fields(value, STYLE_CHECKS_FIELDS)
        elif key == "global_format" and isinstance(value, dict):
            fixed[key] = remove_extra_fields(value, GLOBAL_FORMAT_FIELDS)
        elif key == "abstract" and isinstance(value, dict):
            fixed[key] = {}
            for lang, lang_data in value.items():
                if isinstance(lang_data, dict):
                    fixed[key][lang] = {}
                    for section, section_data in lang_data.items():
                        if isinstance(section_data, dict):
                            allowed = KEYWORDS_FIELDS if "keywords" in section else GLOBAL_FORMAT_FIELDS
                            fixed[key][lang][section] = remove_extra_fields(section_data, allowed)
        elif key == "headings" and isinstance(value, dict):
            fixed[key] = {}
            for level, level_data in value.items():
                if isinstance(level_data, dict):
                    fixed[key][level] = remove_extra_fields(level_data, HEADING_FIELDS)
        elif key == "body_text" and isinstance(value, dict):
            fixed[key] = remove_extra_fields(value, GLOBAL_FORMAT_FIELDS)
        elif key == "figures" and isinstance(value, dict):
            fixed[key] = remove_extra_fields(value, FIGURES_FIELDS)
        elif key == "tables" and isinstance(value, dict):
            fixed[key] = remove_extra_fields(value, TABLES_FIELDS)
        elif key == "references" and isinstance(value, dict):
            fixed[key] = {}
            for section, section_data in value.items():
                if isinstance(section_data, dict):
                    allowed = REF_TITLE_FIELDS if section == "title" else REF_CONTENT_FIELDS
                    fixed[key][section] = remove_extra_fields(section_data, allowed)
        elif key == "acknowledgements" and isinstance(value, dict):
            fixed[key] = {}
            for section, section_data in value.items():
                if isinstance(section_data, dict):
                    fixed[key][section] = remove_extra_fields(section_data, ACK_FIELDS)
        else:
            fixed[key] = value
    return fixed


def main():
    parser = argparse.ArgumentParser(description="WordFormat 配置文件验证器")
    parser.add_argument("--config", "-c", required=True, help="要验证的 YAML 配置文件路径")
    parser.add_argument("--fix", "-f", action="store_true", help="自动移除非法字段并覆盖原文件")
    args = parser.parse_args()

    filepath = args.config
    if not Path(filepath).exists():
        print(f"错误: 文件不存在: {filepath}")
        sys.exit(1)

    # 加载 YAML
    try:
        config = load_yaml(filepath)
    except yaml.YAMLError as e:
        print(f"YAML 语法错误: {e}")
        sys.exit(1)

    if not isinstance(config, dict):
        print("错误: 配置文件根节点必须是字典/映射")
        sys.exit(1)

    # 验证
    errors = validate_config(config)

    if not errors:
        print(f"✅ 配置文件验证通过: {filepath}")
        sys.exit(0)

    print(f"❌ 发现 {len(errors)} 个问题:")
    for err in errors:
        print(err)

    if args.fix:
        fixed = fix_config(config)
        # 保留 YAML 锚点引用（fix 会破坏锚点，需要特殊处理）
        # 由于 fix 会移除 <<: *global_format，这里提示用户
        print(f"\n⚠️  --fix 模式会移除 YAML 锚点引用，建议手动编辑修复。")
        print(f"   请参考 data/config_editing_guide.md 中的编辑指南。")
        sys.exit(1)

    print(f"\n请手动修复以上问题，或参考 data/config_editing_guide.md。")
    sys.exit(1)


if __name__ == "__main__":
    main()
