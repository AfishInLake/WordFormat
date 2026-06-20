<template>
  <div class="abstract-config">
    <h3>中文摘要</h3>
    <h4>中文标题</h4>
    <div class="grid grid-cols-2 gap-4 mb-4">
    </div>
    <FormatConfig :config="config.chinese.chinese_title" />

    <h4 class="mt-4">中文内容</h4>
    <FormatConfig :config="config.chinese.chinese_content" />

    <h3 class="mt-4">英文摘要</h3>
    <h4>英文标题</h4>
    <div class="grid grid-cols-2 gap-4 mb-4">
    </div>
    <FormatConfig :config="config.english.english_title" />

    <h4 class="mt-4">英文内容</h4>
    <FormatConfig :config="config.english.english_content" />

    <h3 class="mt-4">关键词配置</h3>
    <h4>中文关键词</h4>
    <div class="grid grid-cols-2 gap-4 mb-4">
      <div class="form-item">
        <label><input type="checkbox" v-model="config.keywords.chinese.rules.keyword_count.enabled"> 启用数量检查</label>
      </div>
      <div class="form-item">
        <label>最少数量:</label>
        <input type="number" v-model.number="config.keywords.chinese.rules.keyword_count.count_min">
      </div>
      <div class="form-item">
        <label>最大数量:</label>
        <input type="number" v-model.number="config.keywords.chinese.rules.keyword_count.count_max">
      </div>
      <div class="form-item">
        <label><input type="checkbox" v-model="config.keywords.chinese.rules.trailing_punctuation.enabled"> 禁止末尾标点</label>
      </div>
    </div>
    <h5 class="subsection-title">中文关键词内容格式</h5>
    <FormatConfig :config="config.keywords.chinese" />

    <h5 class="subsection-title mt-3">中文关键词标签格式（"关键词："）</h5>
    <FormatConfig :config="config.keywords.chinese.label" />

    <h4 class="mt-4">英文关键词</h4>
    <div class="grid grid-cols-2 gap-4 mb-4">
      <div class="form-item">
        <label><input type="checkbox" v-model="config.keywords.english.rules.keyword_count.enabled"> 启用数量检查</label>
      </div>
      <div class="form-item">
        <label>最少数量:</label>
        <input type="number" v-model.number="config.keywords.english.rules.keyword_count.count_min">
      </div>
      <div class="form-item">
        <label>最大数量:</label>
        <input type="number" v-model.number="config.keywords.english.rules.keyword_count.count_max">
      </div>
    </div>
    <h5 class="subsection-title">英文关键词内容格式</h5>
    <FormatConfig :config="config.keywords.english" />

    <h5 class="subsection-title mt-3">英文关键词标签格式（"Keywords:"）</h5>
    <FormatConfig :config="config.keywords.english.label" />
  </div>
</template>

<script setup>
import FormatConfig from './FormatConfig.vue'

const props = defineProps({
  config: {
    type: Object,
    required: true
  }
})
</script>

<style scoped>
.form-item {
  margin-bottom: 5px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.form-item label {
  font-weight: 500;
  color: #555;
  font-size: 13px;
  white-space: nowrap;
  min-width: 120px;
}

.form-item input[type="text"],
.form-item input[type="number"],
.form-item select {
  width: 90px;
  min-width: 80px;
  padding: 6px 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  font-size: 13px;
}

.form-item input[type="checkbox"] {
  margin-right: 5px;
}

/* 复选框项特殊处理 */
.form-item:has(input[type="checkbox"]) {
  width: auto;
}

.form-item:has(input[type="checkbox"]) label {
  min-width: auto;
}

.grid {
  display: grid;
}

.grid-cols-2 {
  grid-template-columns: repeat(4, 1fr);
}

/* 大屏幕显示更多列 */
@media (min-width: 1024px) {
  .grid-cols-2 {
    grid-template-columns: repeat(5, 1fr);
  }
}

@media (min-width: 1200px) {
  .grid-cols-2 {
    grid-template-columns: repeat(6, 1fr);
  }
}

.gap-4 {
  gap: 10px;
}

.mb-4 {
  margin-bottom: 12px;
}

.subsection-title {
  font-size: 12px;
  font-weight: 600;
  color: #94a3b8;
  margin-bottom: 4px;
}

.mt-3 {
  margin-top: 8px;
}

.mt-4 {
  margin-top: 12px;
}

@media (max-width: 768px) {
  .grid-cols-2 {
    grid-template-columns: 1fr;
  }

  .form-item {
    flex-direction: column;
    align-items: flex-start;
  }

  .form-item label {
    min-width: auto;
  }

  .form-item input[type="text"],
  .form-item input[type="number"],
  .form-item select {
    width: 120px;
  }
}
</style>
