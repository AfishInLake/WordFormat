# =============================================================================
# 主界面导航栏
MENU_ITEMS = [
    ("全局格式", "global_format"),
    ("摘要设置", "abstract"),
    ("标题设置", "headings"),
    ("正文段落", "body_text"),
    ("插图设置", "figures"),
    ("表格设置", "tables"),
    ("参考文献", "references"),
    ("致谢设置", "acknowledgements")
]

OPTIONS_MAP = {
    'alignment': [('左对齐', '左对齐'), ('居中对齐', '居中对齐'), ('右对齐', '右对齐'), ('两端对齐', '两端对齐'),
                  ('分散对齐', '分散对齐')],
    # 对齐方式
    'space_before': [('无', 'NONE'), ('极小', 'TINY'), ('小', 'SMALL'), ('半直线', 'HALF_LINE'), ('正常', 'NORMAL'),
                     ('中等', 'MEDIUM'), ('大', 'LARGE'), ('超大型', 'EXTRA_LARGE')],
    # 段前间距
    'space_after': [('无', 'NONE'), ('极小', 'TINY'), ('小', 'SMALL'), ('半直线', 'HALF_LINE'), ('正常', 'NORMAL'),
                    ('中等', 'MEDIUM'), ('大', 'LARGE'), ('超大型', 'EXTRA_LARGE')],
    # 段后间距
    'line_spacing': [('单倍行距', '单倍行距'), ('1.5倍行距', '1.5倍'), ('双倍行距', '双倍')],
    'first_line_indent': [('无缩进', 'none'), ('1字符', 'one_chars'), ('2字符', 'two_chars'), ('3字符', 'three_chars')],
    'chinese_font_name': [('宋体', '宋体'), ('黑体', 'SIM_HEI'), ('楷体', '楷体'), ('仿宋', '仿宋'),
                          ('微软雅黑', '微软雅黑'), ('汉仪小标宋', '汉仪小标宋')],
    'english_font_name': [('Times New Roman', 'Times New Roman'), ('Arial', 'Arial'), ('Calibri', 'Calibri'),
                          ('Courier New', 'Courier New'), ('Helvetica', 'Helvetica')],
    'font_size': [('一号', '一号'), ('小一', '小一'), ('二号', '二号'), ('小二', '小二'), ('三号', '三号'),
                  ('小三', '小三'), ('四号', '四号'),
                  ('小四', '小四'), ('五号', '五号'), ('小五', '小五'), ('六号', '六号'), ('七号', '七号')],
    'caption_position': [('上方', 'above'), ('下方', 'below')],  # 插图（Figure）格式
    'caption_numbering': [('按文档', 'per_document'), ('按章节', 'per_chapter')],
    'require_label': [('是', True), ('否', False)],
    'bold': [('是', True), ('否', False)],
    'italic': [('是', True), ('否', False)],
    'page_break_before': [('是', True), ('否', False)],
}

TRANSLATIONS = {
    'alignment': '对齐方式', 'space_before': '段前间距', 'space_after': '段后间距',
    'line_spacing': '行间距', 'first_line_indent': '首行缩进', 'chinese_font_name': '中文字体',
    'english_font_name': '英文字体', 'font_size': '字号', 'font_color': '字体颜色',
    'bold': '加粗', 'italic': '斜体', 'underline': '下划线',
    'caption_position': '图注位置', 'caption_prefix': '编号前缀', 'caption_numbering': '编号范围',
    'page_break_before': '段前分页', 'section_title': '章节标题',
    'separator': '关键字分隔符', 'count_min': '最少关键字', 'count_max': '最多关键字',
    'builtin_style_name': '样式名称', 'require_label': '强制图名',
    'level_1': '一级标题', 'level_2': '二级标题', 'level_3': '三级标题',
    'chinese': '中文配置', 'english': '英文配置'
}
