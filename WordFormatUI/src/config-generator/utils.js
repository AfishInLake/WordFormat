// 默认全局段落格式
const defaultParagraph = {
  alignment: "两端对齐",
  space_before: "0行",
  space_after: "0行",
  line_spacingrule: "1.5倍行距",
  line_spacing: "1.5倍",
  left_indent: "0字符",
  right_indent: "0字符",
  first_line_indent: "2字符",
  builtin_style_name: "正文"
}

// 默认全局字体格式
const defaultFont = {
  chinese_font_name: "宋体",
  english_font_name: "Times New Roman",
  font_size: "小四",
  font_color: "黑色",
  bold: false,
  italic: false,
  underline: false
}

// 默认全局格式（平铺，向后兼容）
export const defaultGlobalFormat = {
  ...defaultParagraph,
  ...defaultFont
}

// 创建带全局格式继承的配置对象
export const createConfigWithGlobalInheritance = (overrides = {}) => {
  return { ...defaultGlobalFormat, ...overrides }
}

// 默认配置
export const defaultConfig = {
  template_name: "未知模板",
  style_checks_warning: {
    bold: true,
    italic: false,
    underline: false,
    font_size: true,
    font_name_cn: true,
    font_name_en: true,
    font_color: false,
    alignment: true,
    space_before: true,
    space_after: true,
    line_spacing: true,
    line_spacing_rule: true,
    left_indent: true,
    right_indent: true,
    first_line_indent: true,
    builtin_style_name: true
  },
  global_format: {
    paragraph: { ...defaultParagraph },
    font: { ...defaultFont }
  },
  abstract: {
    chinese: {
      title: createConfigWithGlobalInheritance({
        alignment: "居中对齐",
        first_line_indent: "0字符",
        chinese_font_name: "黑体",
        font_size: "四号"
      }),
      body: createConfigWithGlobalInheritance({
        alignment: "两端对齐"
      }),
      keywords: createConfigWithGlobalInheritance({
        alignment: "两端对齐",
        first_line_indent: "2字符",
        font_size: "小四",
        label: createConfigWithGlobalInheritance({
          chinese_font_name: '黑体',
          font_size: '四号',
          bold: false
        }),
        rules: {
          keyword_count: { enabled: true, count_min: 3, count_max: 5 },
          trailing_punctuation: { enabled: true, forbidden_chars: "；，。、" }
        }
      })
    },
    english: {
      title: createConfigWithGlobalInheritance({
        alignment: "居中对齐",
        first_line_indent: "0字符",
        font_size: "四号",
        bold: false
      }),
      body: createConfigWithGlobalInheritance({
        alignment: "两端对齐"
      }),
      keywords: createConfigWithGlobalInheritance({
        alignment: "两端对齐",
        first_line_indent: "2字符",
        font_size: "小四",
        label: createConfigWithGlobalInheritance({
          font_size: '四号',
          bold: false
        }),
        rules: {
          keyword_count: { enabled: true, count_min: 3, count_max: 5 }
        }
      })
    }
  },
  headings: {
    level_1: createConfigWithGlobalInheritance({
      alignment: "居中对齐",
      first_line_indent: "0字符",
      chinese_font_name: "黑体",
      font_size: "小二",
      space_before: "0.5行",
      space_after: "0.5行",
      builtin_style_name: "Heading 1"
    }),
    level_2: createConfigWithGlobalInheritance({
      alignment: "左对齐",
      first_line_indent: "0字符",
      chinese_font_name: "黑体",
      font_size: "三号",
      space_before: "0行",
      space_after: "0行",
      builtin_style_name: "Heading 2"
    }),
    level_3: createConfigWithGlobalInheritance({
      alignment: "左对齐",
      first_line_indent: "0字符",
      chinese_font_name: "黑体",
      font_size: "小四",
      space_before: "0行",
      space_after: "0行",
      builtin_style_name: "Heading 3"
    })
  },
  body: {
    text: createConfigWithGlobalInheritance({
      rules: { punctuation: { enabled: true } }
    })
  },
  figures: {
    caption: createConfigWithGlobalInheritance({
      alignment: "居中对齐",
      first_line_indent: "0字符",
      font_size: "五号",
      builtin_style_name: "题注",
      caption_prefix: "图",
      rules: {
        caption_numbering: { enabled: true, separator: '.', label_number_space: false }
      }
    }),
    image: createConfigWithGlobalInheritance({
      alignment: "居中对齐",
      first_line_indent: "0字符"
    })
  },
  tables: {
    caption: createConfigWithGlobalInheritance({
      alignment: "居中对齐",
      first_line_indent: "0字符",
      font_size: "五号",
      builtin_style_name: "题注",
      caption_prefix: "表",
      rules: {
        caption_numbering: { enabled: true, separator: '.', label_number_space: false }
      }
    }),
    object: createConfigWithGlobalInheritance({
      alignment: "居中对齐",
      first_line_indent: "0字符"
    })
  },
  references: {
    title: createConfigWithGlobalInheritance({
      alignment: "居中对齐",
      first_line_indent: "0字符",
      chinese_font_name: "黑体",
      font_size: "三号"
    }),
    entry: createConfigWithGlobalInheritance({
      alignment: "两端对齐",
      first_line_indent: "-2.2字符",
      left_indent: "0.26字符",
      chinese_font_name: "宋体",
      font_size: "五号"
    })
  },
  acknowledgements: {
    title: createConfigWithGlobalInheritance({
      alignment: "居中对齐",
      first_line_indent: "0字符",
      chinese_font_name: "黑体",
      font_size: "小二"
    }),
    body: createConfigWithGlobalInheritance({
      alignment: "两端对齐",
      font_size: "五号"
    })
  },
  numbering: {
    enabled: true,
    level_1: {
      enabled: true,
      template: '%1',
      suffix: 'space',
      numbering_indent: null,
      text_indent: null
    },
    level_2: {
      enabled: true,
      template: '%1.%2',
      suffix: 'space',
      numbering_indent: null,
      text_indent: null
    },
    level_3: {
      enabled: true,
      template: '%1.%2.%3',
      suffix: 'space',
      numbering_indent: null,
      text_indent: null
    },
    references: {
      enabled: true,
      template: '[%1]',
      suffix: 'space',
      numbering_indent: null,
      text_indent: null
    }
  }
}

// 对齐方式选项
export const alignmentOptions = ["左对齐", "居中对齐", "右对齐", "两端对齐", "分散对齐"]

// 行距类型选项
export const lineSpacingRuleOptions = ["单倍行距", "1.5倍行距", "2倍行距", "最小值", "固定值", "多倍行距"]

// 中文字体选项
export const chineseFontOptions = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑", "汉仪小标宋"]

// 英文字体选项
export const englishFontOptions = ["Times New Roman", "Arial", "Calibri", "Courier New", "Helvetica"]

// 字号选项
export const fontSizeOptions = ["一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号", "小五", "六号", "七号"]

// 深度合并：用 defaults 填充 config 中缺失的字段
export const mergeWithDefaults = (config, defaults) => {
  const result = { ...config }
  for (const key of Object.keys(defaults)) {
    if (!(key in result)) {
      result[key] = JSON.parse(JSON.stringify(defaults[key]))
    } else if (typeof defaults[key] === 'object' && defaults[key] !== null && !Array.isArray(defaults[key])) {
      result[key] = mergeWithDefaults(result[key], defaults[key])
    }
  }
  return result
}

// 应用全局格式到所有局部配置
export const applyGlobalFormatToAll = (userConfig) => {
  const gf = userConfig.global_format
  const flat = { ...(gf.paragraph || {}), ...(gf.font || {}) }

  // 摘要
  const abs = userConfig.abstract
  abs.chinese.title = { ...flat, ...abs.chinese.title }
  abs.chinese.body = { ...flat, ...abs.chinese.body }
  abs.chinese.keywords = { ...flat, label: abs.chinese.keywords.label, rules: abs.chinese.keywords.rules }
  abs.english.title = { ...flat, ...abs.english.title }
  abs.english.body = { ...flat, ...abs.english.body }
  abs.english.keywords = { ...flat, label: abs.english.keywords.label, rules: abs.english.keywords.rules }

  // 标题
  userConfig.headings.level_1 = { ...flat, ...userConfig.headings.level_1 }
  userConfig.headings.level_2 = { ...flat, ...userConfig.headings.level_2 }
  userConfig.headings.level_3 = { ...flat, ...userConfig.headings.level_3 }

  // 正文
  userConfig.body.text = { ...flat, ...userConfig.body.text }

  // 图/表
  userConfig.figures.caption = { ...flat, caption_prefix: userConfig.figures.caption.caption_prefix, rules: userConfig.figures.caption.rules }
  userConfig.figures.image = { ...flat, ...userConfig.figures.image }
  userConfig.tables.caption = { ...flat, caption_prefix: userConfig.tables.caption.caption_prefix, rules: userConfig.tables.caption.rules }
  userConfig.tables.object = { ...flat, ...userConfig.tables.object }

  // 参考文献
  userConfig.references.title = { ...flat, ...userConfig.references.title }
  userConfig.references.entry = { ...flat, ...userConfig.references.entry }

  // 致谢
  userConfig.acknowledgements.title = { ...flat, ...userConfig.acknowledgements.title }
  userConfig.acknowledgements.body = { ...flat, ...userConfig.acknowledgements.body }
}
