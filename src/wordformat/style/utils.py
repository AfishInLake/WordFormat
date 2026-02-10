import re
from dataclasses import dataclass
from typing import Dict, Optional


# 用dataclass简化类定义（Python 3.7+支持）
@dataclass
class UnitResult:
    """单位提取结果类"""

    # 原始单位（如"磅"/"CM"/"行"/"字符"）
    original_unit: Optional[str] = None
    # 标准化英文字母单位（pt/cm/inch/mm/emu/行/字符）
    standard_unit: Optional[str] = None
    # 数值（float类型）
    value: Optional[float] = None
    # 是否为合法格式（单位在数值后）
    is_valid: bool = False

    def to_dict(self) -> Dict:
        """转换为字典（兼容原有代码）"""
        return {
            "original_unit": self.original_unit,
            "standard_unit": self.standard_unit,
            "value": self.value,
            "is_valid": self.is_valid,
        }

    def convert_to_emu(self) -> Optional[int]:
        """扩展方法：将数值转换为EMU（Office内部单位）
        注意：「行、字符」为非国际标准单位，不参与转换，返回None
        """
        # 非合法格式 或 非国际标准单位 → 返回None
        if not self.is_valid or self.value is None or self.standard_unit is None:
            return None
        if self.standard_unit in ["行", "字符"]:
            return None

        # EMU换算规则（1英寸=914400 EMU）
        emu_mapping = {
            "pt": self.value * 12700,  # 1pt=12700 EMU
            "cm": self.value * 360000,  # 1cm=360000 EMU
            "inch": self.value * 914400,  # 1英寸=914400 EMU
            "mm": self.value * 36000,  # 1mm=36000 EMU
            "emu": self.value,  # 本身就是EMU
        }
        return round(emu_mapping.get(self.standard_unit, 0))

    @property
    def unit_ch(self):
        """
        返回中文单位
        """
        match self.standard_unit:
            case "pt":
                return "磅"
            case "cm":
                return "厘米"
            case "inch":
                return "英寸"
            case "mm":
                return "毫米"
            case "emu":
                return "emu"
            case "hang":
                return "行"
            case "char":
                return "字符"
            case "bei":
                return "倍"
            case None:
                return ""
            case _:
                raise ValueError(f"Invalid unit: {self.standard_unit}")


def extract_unit_from_string(text: str) -> UnitResult:
    """
    从字符串中提取指定单位（新增「行、字符」单位，返回UnitResult类实例）
    支持单位：pt/磅、厘米/CM、英寸/Inches、毫米/mm、Emu、行、字符
    Args:
        text str
    Returns:
        UnitResult
    """
    # 1. 定义单位映射规则（新增「行、字符」，标准化后仍为中文）
    unit_mapping = {
        # 国际标准单位
        "磅": "pt",
        "pt": "pt",
        "PT": "pt",
        "厘米": "cm",
        "cm": "cm",
        "CM": "cm",
        "英寸": "inch",
        "inch": "inch",
        "inches": "inch",
        "Inch": "inch",
        "INCHES": "inch",
        "毫米": "mm",
        "mm": "mm",
        "MM": "mm",
        "emu": "emu",
        "Emu": "emu",
        "EMU": "emu",
        # 非国际标准单位
        "行": "hang",
        "字符": "char",
        "倍": "bei",
    }

    # 2. 初始化结果类
    result = UnitResult()

    # 3. 正则匹配（数值+可选空格+单位）→ 新增「行、字符」匹配
    unit_patterns = [
        r"pt|磅",
        r"厘米|cm",
        r"英寸|inches|inch",
        r"毫米|mm",
        r"emu",
        r"行",
        r"字符",
        r"倍",
    ]
    combined_pattern = r"(?i)(\d+\.?\d*)\s*(?:{})".format("|".join(unit_patterns))
    match = re.search(combined_pattern, text)

    if match:
        # 提取数值
        result.value = float(match.group(1))
        # 提取原始单位
        full_match = match.group(0)
        result.original_unit = re.sub(r"(?i)\d+\.?\d*\s*", "", full_match).strip()
        # 转换为标准化单位
        result.standard_unit = unit_mapping.get(result.original_unit, None)
        # 标记为合法
        result.is_valid = True

    return result
