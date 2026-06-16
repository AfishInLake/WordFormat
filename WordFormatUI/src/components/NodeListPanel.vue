<!-- src/components/NodeListPanel.vue -->
<template>
  <div class="node-list-section">
    <div class="card">
      <div v-if="isLoading" class="loading-tip">
        <div class="loading-spinner"></div>
        <p>{{ loadingText }}</p>
      </div>

      <div v-else-if="!isFileLoaded" class="init-tip">
        <svg class="init-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5"
                d="M19 20H5a2 2 0 01-2-2V6a2 2 0 012-2h10a2 2 0 012 2v1m2 13a2 2 0 01-2-2V7m2 13a2 2 0 002-2V9a2 2 0 00-2-2h-2m-4-3H9m-5 13h14a2 2 0 002-2V9a2 2 0 00-2-2h-2.586a1 1 0 01-.707-.293l-2.414-2.414a1 1 0 00-.707-.293H11a2 2 0 00-2 2v1m2 13a2 2 0 01-2-2V7m2 13c3.314 0 6-2.686 6-6V9a6 6 0 00-6-6H9a6 6 0 00-6 6v8a6 6 0 006 6z"></path>
        </svg>
        <p>请上传docx+yaml配置文件，点击「生成节点JSON」获取节点数据</p>
        <p class="init-sub-tip">生成后可修改标签，再执行「格式校验」或「自动格式化」，支持直接下载结果文件</p>
      </div>

      <div v-else class="node-list">
        <div v-if="nodeData.length === 0" class="empty-tip">无匹配的节点</div>
        <div
            v-for="(node, index) in nodeData"
            :key="index"
            class="node-item"
            :class="{
            error: checkTagError(node),
            other: node.category === 'other',
            selected: selectedIndex === index
          }"
            :style="{ marginLeft: getNodeIndent(node) }"
            @click="$emit('select-node', index)"
        >
          <div class="level-dot" :style="{ backgroundColor: getLevelColor(node) }"></div>
          <div class="node-tag" :class="node.category" :title="CATEGORY_CONFIG[node.category] || ''">
            {{ node.category.replace(/_/g, ' ') }}
          </div>
          <div class="node-score">{{ node.score.toFixed(4) }}</div>
          <div class="node-content" v-html="highlightSearchText(node.paragraph, globalSearchTerm)"></div>
          <span v-if="node.replace" class="replace-badge" title="有替换内容">R</span>
          <div class="node-meta">
            <span class="node-comment">{{ node.comment || '无注释' }}</span>
            <span class="node-fingerprint">{{ node.fingerprint?.slice(0, 8) || '无' }}</span>
          </div>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { checkTagError, getNodeIndent, getLevelColor, highlightSearchText, CATEGORY_CONFIG } from '@/composables/useTagHelpers'

defineProps({
  nodeData: Array,
  selectedIndex: Number,
  isFileLoaded: Boolean,
  isLoading: Boolean,
  loadingText: String,
  globalSearchTerm: String
})

defineEmits(['select-node'])
</script>

<style scoped>
.node-list-section { width: 100%; flex: 1; }
.card {
  background-color: #1e293b;
  border: 1px solid #334155;
  border-radius: 10px;
  padding: 1rem;
  display: flex;
  flex-direction: column;
  height: 100%;
}
.node-list {
  display: flex;
  flex-direction: column;
  gap: 0.125rem;
  overflow-y: auto;
  flex: 1;
  min-height: calc(100vh - 110px);
}
.node-item {
  position: relative;
  margin: 2px 0;
  border-radius: 6px;
  transition: background-color 0.2s ease;
  display: flex;
  align-items: center;
  min-height: 36px;
  padding: 4px 8px;
  cursor: pointer;
  gap: 0.5rem;
  flex-wrap: wrap;
}
.node-item:hover { background-color: #334155; }
.node-item.error { background-color: #450a0a; border-left: 3px solid #ef4444; }
.node-item.other { background-color: #1e293b; border-left: 3px solid #475569; opacity: 0.7; }
.node-item.selected { background-color: #1e3a5f; border-left: 3px solid #3b82f6; opacity: 1; }
.level-dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }
.node-tag {
  min-width: 140px;
  font-size: 12px;
  padding: 2px 6px;
  border-radius: 3px;
  text-align: center;
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  flex-shrink: 0;
}
.node-score {
  min-width: 60px;
  font-size: 12px;
  padding: 2px 6px;
  border-radius: 3px;
  background: #064e3b;
  color: #6ee7b7;
  text-align: center;
  flex-shrink: 0;
}
.node-content {
  flex: 1;
  font-size: 13px;
  color: #e2e8f0;
  word-break: break-all;
  line-height: 1.4;
}
.node-meta {
  font-size: 11px;
  color: #64748b;
  display: flex;
  gap: 10px;
  align-items: center;
  flex-shrink: 0;
  white-space: nowrap;
}
.node-fingerprint { font-family: 'SFMono-Regular', Menlo, Monaco, Consolas, monospace; }
.init-tip, .loading-tip, .empty-tip {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  flex: 1;
  color: #64748b;
  gap: 0.75rem;
  padding: 6rem 0;
}
.init-icon { width: 48px; height: 48px; color: #334155; }
.init-sub-tip { font-size: 12px; color: #475569; margin-top: 0.25rem; text-align: center; line-height: 1.5; }
.loading-spinner {
  width: 20px; height: 20px;
  border: 3px solid #334155;
  border-radius: 50%;
  border-top-color: #22c55e;
  animation: spin 1s ease-in-out infinite;
}
.empty-tip { font-size: 13px; }
@keyframes spin { to { transform: rotate(360deg); } }
.node-tag.other { background: #334155; color: #94a3b8; font-weight: 500; }
.node-tag.abstract_chinese_title { background: #064e3b; color: #6ee7b7; }
.node-tag.abstract_chinese_title_content { background: #0c4a6e; color: #7dd3fc; }
.node-tag.abstract_english_title { background: #4c1d95; color: #c4b5fd; }
.node-tag.abstract_english_title_content { background: #581c87; color: #d8b4fe; }
.node-tag.keywords_chinese { background: #713f12; color: #fde68a; }
.node-tag.keywords_english { background: #7f1d1d; color: #fecaca; }
.node-tag.chinese_title { background: #0c4a6e; color: #7dd3fc; }
.node-tag.english_title { background: #064e3b; color: #6ee7b7; }
.node-tag.heading_level_1 { background: #312e81; color: #c7d2fe; }
.node-tag.heading_level_2 { background: #4c1d95; color: #d8b4fe; }
.node-tag.heading_level_3 { background: #581c87; color: #e9d5ff; }
.node-tag.heading_mulu { background: #713f12; color: #fde68a; }
.node-tag.heading_fulu { background: #1e293b; color: #94a3b8; }
.node-tag.references_title { background: #164e63; color: #67e8f9; }
.node-tag.acknowledgements_title { background: #064e3b; color: #6ee7b7; }
.node-tag.caption_figure { background: #1e293b; color: #94a3b8; }
.node-tag.caption_table { background: #1e293b; color: #94a3b8; }
.node-tag.body_text { background: #0f172a; color: #94a3b8; }
.node-tag.image { background: #0f172a; color: #94a3b8; }
.node-tag.table { background: #0f172a; color: #94a3b8; }
.node-tag.formula { background: #0f172a; color: #94a3b8; }
.search-highlight { background-color: #713f12; color: #fde68a; padding: 0 2px; border-radius: 2px; }
.replace-badge {
  display: inline-flex; align-items: center; justify-content: center;
  width: 20px; height: 20px; border-radius: 50%;
  background: #22c55e; color: #052e16; font-size: 11px; font-weight: 700;
  flex-shrink: 0; cursor: default;
}
</style>
