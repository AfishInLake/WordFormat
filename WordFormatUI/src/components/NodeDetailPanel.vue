<!-- src/components/NodeDetailPanel.vue -->
<template>
  <div class="node-detail-section">
    <div class="card detail-card">
      <h2 class="detail-title">节点详情</h2>
      <div v-if="!currentNode" class="no-select-tip">
        <svg class="no-select-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5"
                d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path>
        </svg>
        <p>生成节点数据后，点击左侧节点查看/修改详情</p>
      </div>

      <div v-else class="node-detail-content">
        <!-- 固定信息 -->
        <div class="property-card">
          <div class="property-card-header">固定信息（不可修改）</div>
          <div class="property-card-body">
            <div class="property-item">
              <div class="property-label">节点序号</div>
              <div class="property-value">{{ nodeIndex }}/{{ totalNodes }}</div>
            </div>
            <div class="property-item">
              <div class="property-label">节点状态</div>
              <div class="property-value status-value" :class="statusClass">
                {{ statusText }}
              </div>
            </div>
            <div class="property-item">
              <div class="property-label">节点层级</div>
              <div class="property-value flex items-center">
                <div class="level-dot mr-2" :style="{ backgroundColor: getLevelColor(currentNode) }"></div>
                {{ levelText }}
              </div>
            </div>
            <div class="property-item">
              <div class="property-label">节点内容</div>
              <div class="property-value content-value">{{ currentNode.paragraph }}</div>
            </div>
            <div class="property-item" v-if="currentNode.replace">
              <div class="property-label replace-label">替换内容</div>
              <div class="property-value content-value replace-value">{{ currentNode.replace }}</div>
            </div>
            <div class="property-item">
              <div class="property-label">置信度得分</div>
              <div class="property-value">{{ currentNode.score.toFixed(4) }}（判定阈值：{{ scoreThreshold }}）</div>
            </div>
            <div class="property-item">
              <div class="property-label">节点指纹</div>
              <div class="property-value fingerprint-value">{{ currentNode.fingerprint || '[无指纹]' }}</div>
            </div>
            <div class="property-item">
              <div class="property-label">节点注释</div>
              <div class="property-value">{{ currentNode.comment || '[无注释]' }}</div>
            </div>
          </div>
        </div>

        <!-- 可变信息 -->
        <div class="property-card mt-4">
          <div class="property-card-header">可变信息（仅标签可修改）</div>
          <div class="property-card-body">
            <div class="property-item">
              <div class="property-label">
                当前分类标签
                <span class="label-tip">(hover查看标签说明)</span>
              </div>
              <select
                  class="category-select"
                  v-model="localCategory"
                  @change="onCategoryChange"
              >
                <option
                    v-for="(desc, key) in categoryConfig"
                    :key="key"
                    :value="key"
                    :title="desc"
                >
                  {{ key.replace(/_/g, ' ') }}
                </option>
              </select>
              <div class="category-desc">{{ categoryConfig[localCategory] || '' }}</div>
            </div>
            <div class="check-result" :class="resultClass">
              {{ resultText }}
            </div>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { computed, ref, watch } from 'vue'
import { checkTagError, getLevelColor, LEVEL_MAP } from '@/composables/useTagHelpers'

const props = defineProps({
  currentNode: Object,
  categoryConfig: Object,
  scoreThreshold: Number,
  totalNodes: Number,
  nodeIndex: Number
})

const emit = defineEmits(['update-category'])

const localCategory = ref(props.currentNode?.category || '')

watch(() => props.currentNode, (newNode) => {
  if (newNode) localCategory.value = newNode.category
})

const onCategoryChange = () => {
  emit('update-category', localCategory.value)
}

const isError = computed(() => checkTagError(props.currentNode, props.scoreThreshold))
const isOther = computed(() => props.currentNode?.category === 'other')

const statusClass = computed(() => {
  if (isOther.value) return 'status-other'
  return isError.value ? 'status-error' : 'status-normal'
})

const statusText = computed(() => {
  if (isOther.value) return '🔖 标记跳过（后端将忽略）'
  return isError.value ? '❌ 疑似标签错误（得分低于阈值）' : '✅ 正常节点（得分高于阈值）'
})

const levelText = computed(() => {
  if (isOther.value) return '标记层级（无缩进）'
  const level = LEVEL_MAP[props.currentNode.category] || 6
  return `层级 ${level} (配置映射)`
})

const resultClass = computed(() => {
  if (isOther.value) return 'result-other'
  return isError.value ? 'result-error' : 'result-normal'
})

const resultText = computed(() => {
  if (isOther.value) return '📌 已标记为跳过：后端处理时直接忽略该节点（防止误删）'
  const score = props.currentNode.score.toFixed(4)
  if (isError.value) return `❌ 疑似标签错误：得分(${score}) ＜ 阈值(${props.scoreThreshold})`
  return `✅ 标签匹配正常：得分(${score}) ≥ 阈值(${props.scoreThreshold})`
})
</script>

<style scoped>
.node-detail-section { margin-left: auto; }
.detail-card {
  position: sticky;
  top: 84px;
  max-height: calc(100vh - 120px);
  min-height: auto;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  background-color: #1e293b;
  border: 1px solid #334155;
  border-radius: 10px;
  padding: 1rem;
}
.detail-title {
  font-size: 14px;
  font-weight: 600;
  color: #e2e8f0;
  padding-bottom: 0.5rem;
  border-bottom: 1px solid #334155;
  margin-bottom: 1rem;
}
.no-select-tip {
  display: flex; flex-direction: column; align-items: center;
  justify-content: center; flex: 1; color: #64748b; gap: 0.75rem; padding: 6rem 0;
}
.no-select-icon { width: 48px; height: 48px; color: #334155; }
.property-card {
  border: 1px solid #334155; border-radius: 6px; overflow: hidden; margin-bottom: 8px;
}
.property-card-header {
  padding: 6px 8px; font-size: 12px; font-weight: 500;
  background-color: #0f172a; color: #94a3b8;
  border-bottom: 1px solid #334155;
}
.property-card-body {
  padding: 8px; display: flex; flex-direction: column; gap: 0.75rem;
}
.property-item { display: flex; flex-direction: column; gap: 2px; }
.property-label { font-size: 12px; font-weight: 500; color: #64748b; }
.label-tip { font-size: 11px; font-weight: 400; color: #475569; }
.property-value {
  font-size: 13px; color: #e2e8f0;
  padding: 6px 8px; border-radius: 6px;
  border: 1px solid #334155; background-color: #0f172a;
  line-height: 1.4; word-break: break-all;
}
.content-value { max-height: 8rem; overflow-y: auto; }
.replace-label { color: #4ade80; font-weight: 600; }
.replace-value { border-color: #166534; background-color: #052e16; color: #6ee7b7; }
.status-value { border: none; padding: 6px 8px; }
.status-normal { background-color: #052e16; color: #4ade80; }
.status-error { background-color: #450a0a; color: #fca5a5; }
.status-other { background-color: #1e293b; color: #94a3b8; }
.category-select {
  width: 100%; font-size: 13px; padding: 6px 8px;
  border-radius: 6px; border: 1px solid #475569;
  background: #0f172a; color: #e2e8f0; cursor: pointer;
  margin: 4px 0; box-sizing: border-box;
}
.category-select:focus { outline: 1px solid #22c55e; border-color: #22c55e; }
.category-desc { font-size: 11px; color: #64748b; line-height: 1.4; padding: 4px 0 0; }
.check-result { padding: 6px 8px; border-radius: 6px; font-size: 13px; text-align: center; margin-top: 8px; }
.result-normal { background-color: #052e16; color: #4ade80; }
.result-error { background-color: #450a0a; color: #fca5a5; }
.result-other { background-color: #1e293b; color: #94a3b8; }
.flex { display: flex; }
.items-center { align-items: center; }
.mr-2 { margin-right: 0.5rem; }
.mt-4 { margin-top: 1rem; }
.fingerprint-value { font-family: 'SFMono-Regular', Menlo, Monaco, Consolas, monospace; font-size: 11px; }
</style>
