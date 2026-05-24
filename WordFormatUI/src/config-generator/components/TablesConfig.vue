<template>
  <div class="tables-config">
    <div class="grid grid-cols-2 gap-4 mb-4">
      <div class="form-item">
        <label>表注位置:</label>
        <select v-model="config.caption_position">
          <option v-for="option in captionPositionOptions" :key="option" :value="option">{{ option }}</option>
        </select>
      </div>
      <div class="form-item">
        <label>表注编号前缀:</label>
        <input type="text" v-model="config.caption_prefix">
      </div>
    </div>
    <FormatConfig :config="config" />

    <h3 class="mt-4">表格内容格式（单元格内文字）</h3>
    <FormatConfig :config="config.content" />
  </div>
</template>

<script setup>
import { captionPositionOptions } from '../utils'
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
  min-width: 100px;
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
