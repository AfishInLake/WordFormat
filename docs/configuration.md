# 配置文件说明

本文档详细说明 WordFormat 的配置文件格式和自定义配置方法，帮助您根据不同学校/期刊的格式要求定制专属格式规范。

## 配置文件格式

格式规范通过 YAML 文件定义，示例参考项目中：
- `example/undergrad_thesis.yaml`（本科毕业论文模板）
- `example/grad_thesis.yaml`（研究生毕业论文模板）

所有配置项均有清晰注释，便于理解和修改。

## 核心配置项

### 1. 全局格式配置（style_checks_warning）
用于控制格式校验时，哪些属性不满足规范时触发警告提示。
```yaml
style_checks_warning:
  bold: true
  italic: true
  underline: true
  font_size: true
  font_name: false
  font_color: false
  alignment: true
  space_before: true
  space_after: true
  line_spacing: true
  line_spacingrule: true
  left_indent: true
  right_indent: true
  first_line_indent: true
  builtin_style_name: true
```

### 2. 全局基础格式（global_format）
所有段落格式默认继承的基础样式，可通过锚点 `&global_format` 被其他节点复用。
```yaml
global_format: &global_format
  alignment: '两端对齐'
  space_before: "0.5行"
  space_after: "0.5行"
  line_spacingrule: "1.5倍行距"
  line_spacing: '1.5倍'
  left_indent: "0字符"
  right_indent: "0字符"
  first_line_indent: '2字符'
  builtin_style_name: '正文'
  chinese_font_name: '宋体'
  english_font_name: 'Times New Roman'
  font_size: '小四'
  font_color: '黑色'
  bold: false
  italic: false
  underline: false
```

### 3. 摘要及关键词配置（abstract）
包含中文摘要、英文摘要及对应关键词的统一配置节点。
```yaml
abstract:
  chinese:
    chinese_title:
      <<: *global_format
      alignment: '居中对齐'
      first_line_indent: '0字符'
      chinese_font_name: '黑体'
      font_size: '小二'
      bold: true
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
    english_content:
      <<: *global_format
      alignment: '两端对齐'
      english_font_name: 'Times New Roman'
      font_size: '小四'
  keywords:
    chinese:
      <<: *global_format
      alignment: '左对齐'
      first_line_indent: '0字符'
      chinese_font_name: '黑体'
      font_size: '小四'
      keywords_bold: true
      count_min: 4
      count_max: 4
      trailing_punct_forbidden: true
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
```

### 4. 各级标题格式（headings）
配置一级、二级、三级标题的样式，对应 Word 内置标题样式。
```yaml
headings:
  level_1:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '小二'
    bold: false
    builtin_style_name: 'Heading 1'

  level_2:
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '三号'
    bold: false
    builtin_style_name: 'Heading 2'

  level_3:
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    english_font_name: 'Times New Roman'
    font_size: '小四'
    bold: false
    builtin_style_name: 'Heading 3'
```

### 5. 正文段落格式（body_text）
论文正文主体内容格式，直接继承全局格式，无需重复定义。
```yaml
body_text:
  <<: *global_format
```

### 6. 插图格式（figures）
配置图片及其题注的格式，题注默认位于图片下方。
```yaml
figures:
  <<: *global_format
  caption_position: 'below'
  caption_prefix: '图'
  font_size: '五号'
  builtin_style_name: '题注'
  alignment: '居中对齐'
  first_line_indent: '0字符'
```

### 7. 表格格式（tables）
配置表格及其题注的格式，题注默认位于表格上方。
```yaml
tables:
  <<: *global_format
  caption_position: 'above'
  caption_prefix: '表'
  chinese_font_name: '宋体'
  font_size: '五号'
  english_font_name: 'Times New Roman'
  builtin_style_name: '题注'
  alignment: '居中对齐'
  first_line_indent: '0字符'
```

### 8. 参考文献格式（references）
包含参考文献标题与条目内容的格式，支持编号、缩进、结尾标点控制。
```yaml
references:
  title:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    font_size: '三号'
    bold: true
    section_title: '参考文献'
  content:
    <<: *global_format
    alignment: '左对齐'
    first_line_indent: '0字符'
    chinese_font_name: '宋体'
    english_font_name: 'Times New Roman'
    font_size: '五号'
    entry_indent: 0.0
    entry_ending_punct: '.'
    numbering_format: '[1], [2], ...'
```

### 9. 致谢格式（acknowledgements）
配置致谢章节的标题与正文格式。
```yaml
acknowledgements:
  title:
    <<: *global_format
    alignment: '居中对齐'
    first_line_indent: '0字符'
    chinese_font_name: '黑体'
    font_size: '小二'
    bold: true
  content:
    <<: *global_format
    alignment: '两端对齐'
    first_line_indent: '2字符'
```

## 字段详细说明

### 格式警告控制字段（style_checks_warning）
| 配置项 | 说明 | 取值 |
|--------|------|------|
| bold | 是否对加粗不规范进行警告 | true/false |
| italic | 是否对斜体不规范进行警告 | true/false |
| underline | 是否对下划线不规范进行警告 | true/false |
| font_size | 是否对字号不规范进行警告 | true/false |
| font_name | 是否对字体名称不规范进行警告 | true/false |
| font_color | 是否对字体颜色不规范进行警告 | true/false |
| alignment | 是否对齐方式不规范进行警告 | true/false |
| space_before | 是否对段前间距不规范进行警告 | true/false |
| space_after | 是否对段后间距不规范进行警告 | true/false |
| line_spacing | 是否对行距数值不规范进行警告 | true/false |
| line_spacingrule | 是否对行距类型不规范进行警告 | true/false |
| left_indent | 是否对左缩进不规范进行警告 | true/false |
| right_indent | 是否对右缩进不规范进行警告 | true/false |
| first_line_indent | 是否对首行缩进不规范进行警告 | true/false |
| builtin_style_name | 是否对内置样式不规范进行警告 | true/false |

### 段落格式字段
| 配置项 | 说明 | 支持单位 | 示例值 |
|-------|------|--------|--------|
| alignment | 段落对齐方式 | - | 左对齐、居中对齐、右对齐、两端对齐、分散对齐 |
| space_before | 段前间距 | 行/磅/毫米/厘米/英寸 | 0.5行、12磅、0.5cm |
| space_after | 段后间距 | 行/磅/毫米/厘米/英寸 | 0.5行、12磅、0.5cm |
| line_spacingrule | 行距规则 | - | 单倍行距、1.5倍行距、2倍行距、最小值、固定值、多倍行距 |
| line_spacing | 行距数值 | 倍 | 1倍、1.5倍、2倍 |
| left_indent | 左缩进 | 字符/磅/毫米/厘米/英寸 | 0字符、2字符、20磅 |
| right_indent | 右缩进 | 字符/磅/毫米/厘米/英寸 | 0字符、2字符 |
| first_line_indent | 首行缩进 | 字符/磅/毫米/厘米/英寸 | 2字符、20磅 |
| builtin_style_name | Word内置样式名 | - | 正文、Heading 1、Heading 2、题注 |

### 字符格式字段
| 配置项 | 说明 | 可选值 |
|-------|------|--------|
| chinese_font_name | 中文字体 | 宋体、黑体、楷体、仿宋、微软雅黑、汉仪小标宋 |
| english_font_name | 英文字体 | Times New Roman、Arial、Calibri、Courier New、Helvetica |
| font_size | 字号 | 一号、小一、二号、小二、三号、小三、四号、小四、五号、小五、六号、七号，或数值（如 12、14） |
| font_color | 字体颜色 | 黑色、红色、十六进制色值 |
| bold | 是否加粗 | true/false |
| italic | 是否斜体 | true/false |
| underline | 是否下划线 | true/false |

### 关键词专用字段
| 配置项 | 说明 | 取值 |
|-------|------|--------|
| keywords_bold | 关键词是否加粗 | true/false |
| count_min | 最少关键词数量 | 正整数 |
| count_max | 最多关键词数量 | 正整数 |
| trailing_punct_forbidden | 是否禁止末尾标点 | true/false |

### 图表专用字段
| 配置项 | 说明 | 可选值 |
|-------|------|--------|
| caption_position | 题注位置 | above（上方）、below（下方） |
| caption_prefix | 题注前缀 | 图、表 |

### 参考文献专用字段
| 配置项 | 说明 | 示例 |
|-------|------|------|
| section_title | 章节标题 | 参考文献 |
| entry_indent | 条目缩进 | 0.0、0.5 |
| entry_ending_punct | 条目结尾标点（null 表示不限制） | . 、null |
| numbering_format | 编号格式 | [1]、1. |

## 配置继承机制
使用 YAML 锚点 `&` 与引用 `<<:` 实现样式复用：
```yaml
global_format: &global_format
  alignment: '两端对齐'

headings:
  level_1:
    <<: *global_format
    alignment: '居中对齐'
```

## 配置验证
配置加载时会自动校验：
- YAML 语法合法性
- 字段取值范围
- 继承锚点有效性
- 必填字段完整性

## 自定义配置流程
1. 复制 `example/` 下现有模板
2. 修改对应格式节点参数
3. 保存为新 YAML 文件使用
4. 执行命令校验配置有效性

## 故障排查
- 配置加载失败：检查 YAML 缩进、引号、冒号是否规范
- 格式不生效：确认段落类别识别正确、配置路径正确
- 警告异常：调整 `style_checks_warning` 中对应开关
- 样式不匹配：检查内置样式名称 `builtin_style_name` 是否正确