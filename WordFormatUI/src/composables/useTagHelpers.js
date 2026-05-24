// src/composables/useTagHelpers.js

export const CATEGORY_CONFIG = {
    "abstract_chinese_title": "仅当段落是“摘要”或“摘 要”（允许尾随空格或冒号）",
    "abstract_chinese_title_content": "当且仅当摘要和正文在一个段落中",
    "abstract_english_title": "仅当段落是“Abstract”（大小写不敏感，允许尾随空格或冒号）",
    "abstract_english_title_content": "当且仅当摘要和正文在一个段落中",
    "keywords_chinese": "包含“关键词”或“关键字”，后面跟着术语列表",
    "keywords_english": "包含“Keywords”（大小写不敏感），后面跟着英文术语",
    "chinese_title": "论文中文题目（通常简短且无编号）",
    "english_title": "论文英文题目（通常简短且无编号）",
    "heading_level_1": "段落必须以“第X章”或单个阿拉伯数字（如“1”“2”）开头，后接空格和标题文字；仅为名词短语",
    "heading_level_2": "段落必须以“X.Y”格式开头（如“1.1”），仅含一个“.”；后接标题文字无完整句子",
    "heading_level_3": "段落必须以“X.Y.Z”格式开头（如“1.1.1”），仅含两个“.”；后接标题文字无完整句子",
    "heading_mulu": "目录标题：段落等于“目录”或“目 录”（含空格变体）",
    "heading_fulu": "段落等于“附录”",
    "references_title": "段落等于“参考文献”或“References”",
    "acknowledgements_title": "段落和“致谢”或“Acknowledgements”等词意思相近",
    "caption_figure": "以“图 X.Y”或“Figure X.Y”开头的图注",
    "caption_table": "以“表 X.Y”或“Table X.Y”开头的表注",
    "body_text": "包含完整句子、谓语动词/句号；或含“本章/本文”等明确论述的内容",
    "other": "标记跳过：后端处理时直接忽略该节点（防止误删，仅标记）"
};

export const LEVEL_MAP = {
    "chinese_title": 1, "english_title": 1,
    "abstract_chinese_title": 1, "abstract_english_title": 1,
    "keywords_chinese": 2, "keywords_english": 2,
    "references_title": 1, "acknowledgements_title": 1,
    "heading_fulu": 1,
    "heading_mulu": 1,
    "heading_level_1": 1, "heading_level_2": 2, "heading_level_3": 3,
    "abstract_chinese_title_content": 1, "abstract_english_title_content": 1,
    "caption_figure": 5, "caption_table": 5, "body_text": 5, "other": 5
};

export const LEVEL_COLORS = ['#1e40af', '#2563eb', '#3b82f6', '#60a5fa', '#93c5fd', '#9ca3af'];
export const SCORE_THRESHOLD = 0.8;

// 工具函数
export function checkTagError(node, threshold = SCORE_THRESHOLD) {
    return node?.score < threshold;
}

export function getNodeIndent(node, levelMap = LEVEL_MAP) {
    const level = levelMap[node.category] || 6;
    return level === 6 ? '0px' : `${(level - 1) * 24}px`;
}

export function getLevelColor(node, levelMap = LEVEL_MAP, colors = LEVEL_COLORS) {
    const level = levelMap[node.category] || 6;
    return colors[Math.min(level - 1, 5)];
}

export function highlightSearchText(text, searchTerm) {
    if (!searchTerm.trim()) return text;
    const term = searchTerm.trim().toLowerCase();
    const regex = new RegExp(`(${term})`, 'gi');
    return text.split(regex).map(part =>
        part.toLowerCase() === term ? `<span class="search-highlight">${part}</span>` : part
    ).join('');
}
