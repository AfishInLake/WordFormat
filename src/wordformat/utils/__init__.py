"""通用工具包（子模块拆分，顶层重导出保持向后兼容）。"""

from wordformat.utils._docx import para_contains_image, remove_all_numbering
from wordformat.utils._fs import (
    ensure_directory_exists,
    ensure_is_directory,
    get_file_name,
)
from wordformat.utils._text import (
    _count_numbering_levels,
    _format_number,
    _from_chinese_num,
    _from_roman,
    _get_level_fmt,
    _to_chinese_num,
    _to_roman,
    count_chinese_chars,
    extract_chinese_chars,
    get_paragraph_numbering_text,
    has_chinese,
    is_chinese_char,
    parse_caption_text,
)
from wordformat.utils._yaml import load_yaml_with_merge
