// 默认全局格式
export const defaultGlobalFormat = {
  alignment: "左对齐",
  space_before: "0.5行",
  space_after: "0.5行",
  line_spacingrule: "单倍行距",
  line_spacing: "1.5倍",
  left_indent: "0字符",
  right_indent: "0字符",
  first_line_indent: "2字符",
  builtin_style_name: "正文",
  chinese_font_name: "宋体",
  english_font_name: "Times New Roman",
  font_size: "小四",
  font_color: "黑色",
  bold: false,
  italic: false,
  underline: false
}

// 创建带全局格式继承的配置对象
export const createConfigWithGlobalInheritance = (baseConfig = {}) => {
  return {
    ...defaultGlobalFormat,
    ...baseConfig
  }
}

// 默认配置
export const defaultConfig = {
  style_checks_warning: {
    bold: true,
    italic: true,
    underline: true,
    font_size: true,
    font_name: false,
    font_color: false,
    alignment: true,
    space_before: true,
    space_after: true,
    line_spacing: true,
    line_spacingrule: true,
    left_indent: true,
    right_indent: true,
    first_line_indent: true,
    builtin_style_name: true
  },
  global_format: {
    ...defaultGlobalFormat
  },
  abstract: {
    chinese: {
      chinese_title: createConfigWithGlobalInheritance(),
      chinese_content: createConfigWithGlobalInheritance()
    },
    english: {
      english_title: createConfigWithGlobalInheritance(),
      english_content: createConfigWithGlobalInheritance()
    },
    keywords: {
      english: createConfigWithGlobalInheritance({
        label: createConfigWithGlobalInheritance({
          bold: true,
          font_size: '四号'
        }),
        count_min: 4,
        count_max: 4,
        trailing_punct_forbidden: true
      }),
      chinese: createConfigWithGlobalInheritance({
        label: createConfigWithGlobalInheritance({
          chinese_font_name: '黑体',
          font_size: '四号',
          bold: true
        }),
        count_min: 4,
        count_max: 4,
        trailing_punct_forbidden: true
      })
    }
  },
  headings: {
    level_1: createConfigWithGlobalInheritance(),
    level_2: createConfigWithGlobalInheritance(),
    level_3: createConfigWithGlobalInheritance()
  },
  body_text: createConfigWithGlobalInheritance(),
  figures: createConfigWithGlobalInheritance({
    caption_position: "below",
    caption_prefix: "图"
  }),
  tables: createConfigWithGlobalInheritance({
    caption_position: "above",
    caption_prefix: "表",
    content: createConfigWithGlobalInheritance({
      font_size: '五号',
      line_spacingrule: '单倍行距',
      alignment: '居中对齐',
      first_line_indent: '0字符',
      space_before: "0行",
      space_after: "0行"
    })
  }),
  references: {
    title: createConfigWithGlobalInheritance({
      section_title: "参考文献"
    }),
    content: createConfigWithGlobalInheritance({
      numbering_format: null,
      entry_indent: 0.0,
      entry_ending_punct: null
    })
  },
  acknowledgements: {
    title: createConfigWithGlobalInheritance(),
    content: createConfigWithGlobalInheritance()
  },
  numbering: {
    enabled: true,
    level_1: {
      enabled: true,
      template: '%1',
      suffix: 'space'
    },
    level_2: {
      enabled: true,
      template: '%1.%2',
      suffix: 'space'
    },
    level_3: {
      enabled: true,
      template: '%1.%2.%3',
      suffix: 'space'
    },
    references: {
      enabled: true,
      template: '[%1]',
      suffix: 'space'
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

// 图注位置选项
export const captionPositionOptions = ["above", "below"]

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
  // 摘要配置
  userConfig.abstract.chinese.chinese_title = {
    ...userConfig.global_format
  }
  userConfig.abstract.chinese.chinese_content = {
    ...userConfig.global_format
  }
  userConfig.abstract.english.english_title = {
    ...userConfig.global_format
  }
  userConfig.abstract.english.english_content = {
    ...userConfig.global_format
  }
  userConfig.abstract.keywords.english = {
    ...userConfig.global_format,
    label: userConfig.abstract.keywords.english.label,
    count_min: userConfig.abstract.keywords.english.count_min,
    count_max: userConfig.abstract.keywords.english.count_max,
    trailing_punct_forbidden: userConfig.abstract.keywords.english.trailing_punct_forbidden
  }
  userConfig.abstract.keywords.chinese = {
    ...userConfig.global_format,
    label: userConfig.abstract.keywords.chinese.label,
    count_min: userConfig.abstract.keywords.chinese.count_min,
    count_max: userConfig.abstract.keywords.chinese.count_max,
    trailing_punct_forbidden: userConfig.abstract.keywords.chinese.trailing_punct_forbidden
  }

  // 标题配置
  userConfig.headings.level_1 = {
    ...userConfig.global_format
  }
  userConfig.headings.level_2 = {
    ...userConfig.global_format
  }
  userConfig.headings.level_3 = {
    ...userConfig.global_format
  }

  // 正文配置
  userConfig.body_text = {
    ...userConfig.global_format
  }

  // 插图配置
  userConfig.figures = {
    ...userConfig.global_format,
    caption_position: userConfig.figures.caption_position,
    caption_prefix: userConfig.figures.caption_prefix
  }

  // 表格配置
  userConfig.tables = {
    ...userConfig.global_format,
    caption_position: userConfig.tables.caption_position,
    caption_prefix: userConfig.tables.caption_prefix,
    content: userConfig.tables.content
  }

  // 参考文献配置
  userConfig.references.title = {
    ...userConfig.global_format,
    section_title: userConfig.references.title.section_title
  }
  userConfig.references.content = {
    ...userConfig.global_format,
    numbering_format: userConfig.references.content.numbering_format,
    entry_indent: userConfig.references.content.entry_indent,
    entry_ending_punct: userConfig.references.content.entry_ending_punct
  }

  // 致谢配置
  userConfig.acknowledgements.title = {
    ...userConfig.global_format
  }
  userConfig.acknowledgements.content = {
    ...userConfig.global_format
  }
}
