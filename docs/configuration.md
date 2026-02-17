# 配置文件说明

本文档详细说明 WordFormat 的配置文件格式和自定义配置方法，帮助您根据不同学校/期刊的格式要求定制专属格式规范。

## 配置文件格式

格式规范通过 YAML 文件定义，示例参考项目中：
- `example/undergrad_thesis.yaml`（本科毕业论文模板）
- `example/grad_thesis.yaml`（研究生毕业论文模板）

所有配置项均有清晰注释，便于理解和修改。

## 核心配置项

### 1. 全局页面格式（正文通用）

```yaml
# 全局页面格式（正文通用）
global_format: &global_format
  alignment: '两端对齐'
  space_before: "0.5行"
  space_after: "0.5行"
  line_spacingrule: "1.5倍行距"
  line_spacing: '1.5倍'
  left_indent: "0字符"
  right_indent: "0字符"
  first_line_indent: '20磅'
  builtin_style_name: '正文'
  chinese_font_name: '宋体'
  english_font_name: 'Times New Roman'
  font_size: '小四'
  font_color: '黑色'
  bold: false
  italic: false
  underline: false
```

### 2. 摘要部分（中英文）

```yaml
# 摘要部分（中英文）
abstract:
  chinese:
    chinese_title:
      <<: *global_format
      alignment: '居中对齐'
      first_line_indent: '0字符'
      chinese_font_name: '黑体'
      font_size: '小二'
      bold: true
      section_title_re: '^\s*摘\s+要\s*$'
    chinese_content:
      <<: *global_format
      alignment: '两端对齐'
  english:
    english_title:
      <<: *global_format
      alignment: '居中对齐'
      first_line_indent: '0字符'
      english_font_name: 'Times New Roman'
      font_size: '四号'
      bold: true
      line_spacing: "0倍"
      section_title_re: '^\\s*[Aa]b\\s*[Ss]t\\s*[Rr]a\\s*[Cc]t\\s*$'
    english_content:
      <<: *global_format
      alignment: '两端对齐'
      english_font_name: 'Times New Roman'
      font_size: '小四'
      line_spacing: "0倍"
```

### 3. 关键词部分

```yaml
keywords:
  chinese:
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    font_size: '小四'
    keywords_bold: true  # 规范未要求加粗，仅字体为黑体四号
    count_min: 4
    count_max: 4
    trailing_punct_forbidden: true
    section_title_re: '^\s*关键词\s*[：:,，,；;\\s—-]*.*$'
  english:
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    english_font_name: 'Times New Roman'
    font_size: '小四'
    keywords_bold: true
    count_min: 4
    count_max: 4
    trailing_punct_forbidden: true
    section_title_re: '^\\s*[Kk]ey\\s*[Ww]ords\\s*[：:,，,\\s;—-]*.*$'
```

### 4. 各级标题格式

```yaml
# 各级标题格式（章、节、小节）
headings:
  level_1: # 一级标题（如“第一章 绪论”）
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '小二'
    bold: false
    space_before: "0.5行"
    space_after: "0.5行"
    line_spacingrule: "1.5倍行距"
    line_spacing: '1.5倍'  # 标题通常单倍行距
    builtin_style_name: 'Heading 1'
    section_title_re: '^\d+(?![\.\d])\s*'

  level_2: # 二级标题（如“1.1 研究背景”）
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '三号'
    bold: false
    space_before: "0.5行"
    space_after: "0.5行"
    line_spacingrule: "1.5倍行距"
    line_spacing: '1.5倍'
    builtin_style_name: 'Heading 2'
    section_title_re: '^\d+\.\d+(?![\.\d])\s*'

  level_3: # 三级标题（如“1.1.1 国内现状”）
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '小四'
    bold: false
    space_before: "0.5行"
    space_after: "0.5行"
    builtin_style_name: 'Heading 3'
    line_spacingrule: "1.5倍行距"
    line_spacing: '1.5倍'
    section_title_re: '^\d+\.\d+\.\d+(?!\.)\s*'
```

### 5. 正文段落格式

```yaml
# 正文段落格式
body_text:
  <<: *global_format
  # 已在 global_format 中定义，此处可省略或保留以强调
```

### 6. 插图（Figure）格式

```yaml
# 插图（Figure）格式
figures:
  <<: *global_format
  caption_position: 'below'          # 图名置于图之下（规范明确要求）
  caption_prefix: '图'               # 如“图2.1”
  font_size: '五号'
  section_title_re: '^\s*图\s*\d+(\.\d+)*.*$'
  builtin_style_name: '题注'
  alignment: '居中对齐'
  first_line_indent: '0字符'
```

### 7. 表格（Table）格式

```yaml
# 表格（Table）格式
tables:
  <<: *global_format
  caption_position: 'above'          # 表名置于表之上（规范明确要求）
  caption_prefix: '表'               # 如“表5.1”
  chinese_font_name: '宋体'
  font_size: '五号'
  english_font_name: 'Times New Roman'    # 表格内英文用 Times New Roman 五号
  section_title_re: '^表\s*\d+(\.\d+)*'
  builtin_style_name: '题注'
  alignment: '居中对齐'
  first_line_indent: '0字符'
```

### 8. 参考文献（References）

```yaml
# 参考文献（References）
references:
  title:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    font_size: '三号'
    bold: true
    section_title_re: '参考文献'
  content:
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '五号'
    entry_indent: 0.0                # 顶格（序号[1]顶格）
    entry_ending_punct: '.'          # 每条以句号结束
    numbering_format: '[1], [2], ...'
```

### 9. 致谢（Acknowledgements）

```yaml
# 致谢（Acknowledgements）
acknowledgements:
  title:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    font_size: '小二'
    bold: true
    section_title_re: '^\s*致\s+谢\s*$'
  content:
    <<: *global_format
    alignment: '两端对齐'
    first_line_indent: '2字符'
```

## 自定义配置

### 配置继承体系

YAML 配置文件支持使用锚点（`&`）和引用（`<<:`）来实现配置的继承，减少重复配置。例如：

```yaml
# 定义全局格式锚点
global_format: &global_format
  alignment: '两端对齐'
  # 其他配置...

# 引用全局格式并覆盖部分配置
headings:
  level_1:
    <<: *global_format  # 继承全局格式
    alignment: '居中对齐'  # 覆盖对齐方式
    # 其他配置...
```

### 配置项说明

#### 段落格式
| 配置项 | 说明 | 支持单位 | 示例值 |
|-------|------|--------|--------|
| alignment | 对齐方式 | - | '两端对齐', '居中对齐', '左对齐', '右对齐', '分散对齐' |
| space_before | 段前间距 | 行, 磅, 毫米, 厘米, 英寸 | '0.5行', '12磅', '5mm', '0.5cm', '0.2inch' |
| space_after | 段后间距 | 行, 磅, 毫米, 厘米, 英寸 | '0.5行', '12磅', '5mm', '0.5cm', '0.2inch' |
| line_spacingrule | 行距规则 | - | '单倍行距', '1.5倍行距', '2倍行距', '最小值', '固定值', '多倍行距' |
| line_spacing | 行距 | 倍 | '1倍', '1.5倍', '2倍' |
| left_indent | 左缩进 | 字符, 磅, 毫米, 厘米, 英寸 | '0字符', '2字符', '20磅', '5mm', '0.5cm' |
| right_indent | 右缩进 | 字符, 磅, 毫米, 厘米, 英寸 | '0字符', '2字符', '20磅', '5mm', '0.5cm' |
| first_line_indent | 首行缩进 | 字符, 磅, 毫米, 厘米, 英寸 | '2字符', '20磅', '5mm', '0.5cm' |

#### 字符格式
| 配置项 | 说明 | 支持单位 | 示例值 |
|-------|------|--------|--------|
| chinese_font_name | 中文字体 | - | '宋体', '黑体', '楷体', '仿宋', '微软雅黑', '汉仪小标宋' |
| english_font_name | 英文字体 | - | 'Times New Roman', 'Arial', 'Calibri', 'Courier New', 'Helvetica' |
| font_size | 字号 | 磅, 中文字号 | '小四', '五号', '12pt', '10.5', '16' |
| font_color | 字体颜色 | - | '黑色', '红色', '#FF0000', '#f00', 'blue', '浅蓝色' |
| bold | 是否加粗 | - | true, false |
| italic | 是否斜体 | - | true, false |
| underline | 是否下划线 | - | true, false |

#### 标题格式
| 配置项 | 说明 | 支持单位 | 示例值 |
|-------|------|--------|--------|
| builtin_style_name | Word内置样式名称 | - | 'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Normal', 'Title', 'Subtitle', 'List Paragraph', 'Caption' |
| section_title_re | 标题正则表达式 | - | '^\d+(?![\.\d])\s*' |

#### 特殊段落
| 配置项 | 说明 | 支持单位 | 示例值 |
|-------|------|--------|--------|
| caption_position | 题注位置 | - | 'above', 'below' |
| caption_prefix | 题注前缀 | - | '图', '表' |
| entry_indent | 参考文献条目缩进 | - | 0.0, 0.5 |
| entry_ending_punct | 参考文献条目结束标点 | - | '.', '' |
| numbering_format | 参考文献编号格式 | - | '[1], [2], ...', '1. 2. ...' |

## 配置验证

配置文件加载时会进行验证，确保所有必需的配置项都已提供且格式正确。如果配置加载失败，会显示详细的错误信息。

### 常见配置错误

1. **YAML 语法错误**：
   - 缩进不正确
   - 缺少冒号
   - 字符串引号不匹配

2. **配置项值错误**：
   - 对齐方式值不正确
   - 行距规则值不正确
   - 字体名称不存在

3. **引用错误**：
   - 引用了未定义的锚点
   - 循环引用

## 配置模板管理

### 创建自定义模板

1. 复制现有的模板文件（如 `example/undergrad_thesis.yaml`）
2. 根据需要修改配置项
3. 保存为新的模板文件，放入 `example/` 目录

### 模板推荐

- **本科毕业论文**：使用 `example/undergrad_thesis.yaml`
- **研究生毕业论文**：使用 `example/grad_thesis.yaml`
- **期刊论文**：基于现有模板修改，重点关注摘要、关键词和参考文献格式
- **会议论文**：通常要求更严格的格式，建议参考会议官方模板创建配置

## 最佳实践

1. **从模板开始**：使用现有的模板作为起点，避免从零开始配置
2. **逐步修改**：一次修改少量配置项，测试效果后再继续
3. **文档化配置**：为自定义配置添加注释，说明修改原因和预期效果
4. **版本控制**：将配置文件纳入版本控制系统，跟踪变更历史
5. **共享配置**：将经过验证的配置模板分享给其他用户，共同完善模板库

## 故障排查

### 配置文件加载失败

- 检查 YAML 语法是否正确
- 确保所有必需的配置项都已提供
- 验证配置文件路径是否正确
- 查看日志输出，获取详细的错误信息

### 格式检查无效果

- 确认配置文件中的规则与预期格式一致
- 检查 JSON 结构文件中的段落类别是否正确识别
- 验证配置文件是否被正确加载（查看日志输出）
- 尝试使用默认模板测试，排除配置文件问题
