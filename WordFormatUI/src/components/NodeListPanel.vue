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
.node-list-section { min-height: 0; display: flex; flex-direction: column; }
.card {
  background-color: var(--paper);
  border: 1px solid var(--border);
  border-radius: 10px;
  padding: 1rem;
  display: flex;
  flex-direction: column;
  flex: 1;
  min-height: 0;
}
.node-list {
  display: flex;
  flex-direction: column;
  gap: 0.125rem;
  overflow-y: auto;
  flex: 1;
}
.node-item {
  display: flex;
  align-items: flex-start;
  gap: 0.5rem;
  margin: 2px 0;
  border-radius: 6px;
  padding: 4px 8px;
  cursor: pointer;
  transition: background-color 0.2s ease;
}
.node-item:hover { background-color: var(--surface); }
.node-item.error { background-color: var(--red-dim); border-left: 3px solid var(--red); }
.node-item.other { background-color: var(--paper); border-left: 3px solid var(--border-hover); opacity: 0.7; }
.node-item.selected { background-color: var(--raised); border-left: 3px solid var(--brass); opacity: 1; }
.level-dot { width: 8px; height: 8px; margin-top: 0.4em; border-radius: 50%; flex-shrink: 0; }
.node-tag {
  width: 140px;
  max-width: 140px;
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
  width: 60px;
  font-size: 12px;
  padding: 2px 6px;
  border-radius: 3px;
  background: var(--green-dim);
  color: var(--green);
  text-align: center;
  flex-shrink: 0;
}
.node-content {
  flex: 1;
  min-width: 0;
  font-size: 13px;
  color: var(--text);
  word-break: break-all;
  overflow-wrap: break-word;
  line-height: 1.5;
}
.node-meta {
  font-size: 11px;
  color: var(--text-muted);
  flex-shrink: 1;
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.init-tip, .loading-tip, .empty-tip {
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  flex: 1;
  color: var(--text-muted);
  gap: 0.75rem;
  padding: 6rem 0;
}
.init-icon { width: 48px; height: 48px; color: var(--border); }
.init-sub-tip { font-size: 12px; color: var(--text-muted); margin-top: 0.25rem; text-align: center; line-height: 1.5; }
.loading-spinner {
  width: 20px; height: 20px;
  border: 3px solid var(--border);
  border-radius: 50%;
  border-top-color: var(--brass);
  animation: spin 1s ease-in-out infinite;
}
.empty-tip { font-size: 13px; }
@keyframes spin { to { transform: rotate(360deg); } }
.node-tag.other { background: var(--surface); color: var(--text-muted); }
.node-tag.abstract_chinese_title { background: #1a3d2e; color: #6ee7b7; }
.node-tag.abstract_chinese_title_content { background: #1a3540; color: #7dd3fc; }
.node-tag.abstract_english_title { background: #2d1a4a; color: #c4b5fd; }
.node-tag.abstract_english_title_content { background: #341a4a; color: #d8b4fe; }
.node-tag.keywords_chinese { background: #3d2a12; color: #fde68a; }
.node-tag.keywords_english { background: #3d1a1a; color: #fecaca; }
.node-tag.chinese_title { background: #1a3540; color: #7dd3fc; }
.node-tag.english_title { background: #1a3d2e; color: #6ee7b7; }
.node-tag.heading_level_1 { background: #1f1d3d; color: #c7d2fe; }
.node-tag.heading_level_2 { background: #2d1a4a; color: #d8b4fe; }
.node-tag.heading_level_3 { background: #341a4a; color: #e9d5ff; }
.node-tag.heading_mulu { background: #3d2a12; color: #fde68a; }
.node-tag.heading_fulu { background: var(--surface); color: var(--text-muted); }
.node-tag.references_title { background: #1a3038; color: #67e8f9; }
.node-tag.acknowledgements_title { background: #1a3d2e; color: #6ee7b7; }
.node-tag.caption_figure { background: var(--surface); color: var(--text-muted); }
.node-tag.caption_table { background: var(--surface); color: var(--text-muted); }
.node-tag.body_text { background: var(--ink); color: var(--text-muted); }
.search-highlight { background-color: var(--brass-dim); color: var(--text); padding: 0 2px; border-radius: 2px; }
.replace-badge {
  display: inline-flex; align-items: center; justify-content: center;
  width: 20px; height: 20px; border-radius: 50%;
  background: var(--brass); color: var(--ink); font-size: 11px; font-weight: 700;
  flex-shrink: 0;
}
</style>
