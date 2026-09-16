<!-- src/components/DocTagChecker.vue -->
<template>
  <div class="doc-tag-check-container">
    <!-- 顶部工具栏 -->
    <div class="header-bar">
      <div class="header-content">
        <div class="header-left">
          <h1 class="tool-title">文档标签核对工具</h1>
          <div class="stats-info" v-if="isFileLoaded">
            共 <span>{{ nodeCount }}</span> 个节点 |
            疑似错误 <span>{{ errorCount }}</span> 个 |
            标记跳过 <span>{{ otherCount }}</span> 个
          </div>
        </div>

        <div class="header-right">
          <!-- 文件上传面板 -->
          <FileUploadPanel
              v-model:docx-file="docxFile"
              v-model:yaml-file="yamlFile"
              :generated-config="generatedConfig"
              :is-loading="isLoading"
              @generate-json="callGenerateJsonApi"
          />

          <!-- 格式操作面板（仅在加载后显示） -->
          <FormatActionPanel
              v-if="isFileLoaded"
              :is-loading="isLoading"
              @check-format="callCheckFormatApi"
              @apply-format="callApplyFormatApi"
          />

          <!-- 搜索框 -->
          <div class="search-box">
            <input
                type="text"
                v-model="searchTerm"
                placeholder="搜索段落内容..."
                class="search-input"
            />
            <svg class="search-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
                    d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"/>
            </svg>
          </div>

          <!-- 辅助按钮 -->
          <button class="btn secondary-btn" @click="checkAllTags" :disabled="!isFileLoaded || isLoading">
            核对所有标签
          </button>
          <button class="btn secondary-btn" @click="importJsonFile" :disabled="!isFileLoaded || isLoading">
            导入JSON
          </button>
          <button class="btn primary-btn" @click="exportModifiedJson" :disabled="!isFileLoaded">
            导出JSON
          </button>
        </div>
      </div>
    </div>

    <!-- 隐藏的 file input 用于导入 JSON -->
    <input
      ref="jsonFileInput"
      type="file"
      accept=".json"
      style="display: none"
      @change="handleJsonImport"
    />

    <!-- 主体内容区 -->
    <div class="main-content">
      <!-- 节点列表 -->
      <NodeListPanel
          :node-data="filteredNodeData"
          :selected-index="selectedDisplayIndex"
          :is-file-loaded="isFileLoaded"
          :is-loading="isLoading"
          :loading-text="loadingText"
          :global-search-term="searchTerm"
          @select-node="selectNode"
      />

      <!-- 节点详情 -->
      <NodeDetailPanel
          :current-node="currentNode"
          :category-config="CATEGORY_CONFIG"
          :score-threshold="thresholdFor(currentNode?.category)"
          :total-nodes="visibleNodeData.length"
          :node-index="selectedDisplayIndex + 1"
          @update-category="handleCategoryChange"
      />
    </div>
  </div>
</template>

<script setup>
import {ref, computed, watch} from 'vue'
import FileUploadPanel from './FileUploadPanel.vue'
import FormatActionPanel from './FormatActionPanel.vue'
import NodeListPanel from './NodeListPanel.vue'
import NodeDetailPanel from './NodeDetailPanel.vue'
import {
  CATEGORY_CONFIG,
  thresholdFor,
  checkTagError
} from '../composables/useTagHelpers'

import request from "../utils/request.js";
import { backendBaseUrl } from '../utils/settings';

// ====== Props ======
const props = defineProps({
  generatedConfig: {
    type: Object,
    default: null
  }
});

// 创建响应式引用，确保能够使用最新的props值
const currentConfig = ref(props.generatedConfig);

// 监听props变化，更新响应式引用
watch(
  () => props.generatedConfig,
  (newConfig) => {
    currentConfig.value = newConfig;
  },
  { deep: true }
);

// ====== 响应式状态 ======
const docxFile = ref(null)
const yamlFile = ref(null)
const nodeData = ref([]) // 原始节点数据（可修改）
const isFileLoaded = ref(false)
const isLoading = ref(false)
const loadingText = ref('')
const selectedNodeIndex = ref(-1)
const searchTerm = ref('')
const jsonFileInput = ref(null)

// ====== 计算属性 ======
const currentNode = computed(() => {
  return selectedNodeIndex.value >= 0 ? nodeData.value[selectedNodeIndex.value] : null
})

// 可见节点：figure_image 占位节点仅在显示层隐藏，绝不能从 nodeData 删除——
// 否则节点数组与文档段落失去 1:1 对应，后端按位置 zip 时整段错位，
// 会把 keywords_chinese/caption_figure/heading_level_* 的规则写到错误段落上。
const visibleNodeData = computed(() =>
    nodeData.value.filter(node => node.category !== 'figure_image')
)

const nodeCount = computed(() => visibleNodeData.value.length)

const errorCount = computed(() =>
    visibleNodeData.value.filter(node => checkTagError(node)).length
)

const otherCount = computed(() =>
    visibleNodeData.value.filter(node => node.category === 'other').length
)

const filteredNodeData = computed(() => {
  const term = searchTerm.value.trim().toLowerCase()
  const base = visibleNodeData.value
  if (!term) return base
  return base.filter(node =>
      node.paragraph.toLowerCase().includes(term)
  )
})

// 列表渲染用的是过滤后的索引，当前选中节点需要换算回过滤列表中的位置
// （selectedNodeIndex 保存的是 nodeData 中的原始索引）
const selectedDisplayIndex = computed(() => {
  if (selectedNodeIndex.value < 0) return -1
  const node = nodeData.value[selectedNodeIndex.value]
  if (!node) return -1
  return filteredNodeData.value.indexOf(node)
})

// ====== 方法 ======
// 选择节点：index 来自列表渲染（filteredNodeData）的索引，需换算回 nodeData 原始索引
const selectNode = (index) => {
  const node = filteredNodeData.value[index]
  selectedNodeIndex.value = node ? nodeData.value.indexOf(node) : -1
}

// 修改当前节点的分类标签
const handleCategoryChange = (newCategory) => {
  if (currentNode.value) {
    nodeData.value[selectedNodeIndex.value].category = newCategory
    // 自动重新计算置信度？这里保留原始 score（因无模型），仅改标签
    // 若需重新打分，需调用后端接口
  }
}

// 生成节点 JSON（上传文件）
const callGenerateJsonApi = async () => {
  if (!docxFile.value) return

  const formData = new FormData()
  formData.append('docx_file', docxFile.value)

  // 优先使用生成的配置，否则使用选择的yaml文件
  if (currentConfig.value) {
    // 将生成的配置转换为YAML字符串
    const yaml = await import('js-yaml');
    const yamlContent = yaml.dump(currentConfig.value, { indent: 2, skipInvalid: true });
    const blob = new Blob([yamlContent], { type: 'application/yaml' });
    const configFile = new File([blob], 'generated-config.yaml', { type: 'application/yaml' });
    formData.append('config_file', configFile);
  } else if (yamlFile.value) {
    formData.append('config_file', yamlFile.value)
  } else {
    return
  }

  isLoading.value = true
  loadingText.value = '正在解析文档并生成节点...'

  try {
    const res = await request.post(
        '/generate-json',
        formData,
        {
          timeout: 300_000,
          headers: {'Content-Type': 'multipart/form-data'}
        }
    )

    if (res.data && Array.isArray(res.data.json_data)) {
      // 保留全部节点（含 figure_image 占位），保证与文档段落 1:1 对应；
      // 隐藏 figure_image 交给显示层 filteredNodeData 处理
      nodeData.value = res.data.json_data
        .map((node, idx) => ({
          ...node,
          id: idx
        }))
      isFileLoaded.value = true
      selectedNodeIndex.value = -1
    } else {
      alert('❌ 后端返回数据格式异常')
    }
  } catch (error) {
    console.error('生成JSON失败:', error)
    alert('❌ 生成节点失败：' + (error.response?.data?.detail || error.message))
  } finally {
    isLoading.value = false
  }
}

// 执行格式校验
const callCheckFormatApi = async () => {
  if (!isFileLoaded.value) return

  isLoading.value = true
  loadingText.value = '正在执行格式校验...'
  if (!docxFile.value) return

  const formData = new FormData()
  formData.append('docx_file', docxFile.value)

  // 优先使用生成的配置，否则使用选择的yaml文件
  if (currentConfig.value) {
    // 将生成的配置转换为YAML字符串
    const yaml = await import('js-yaml');
    const yamlContent = yaml.dump(currentConfig.value, { indent: 2, skipInvalid: true });
    const blob = new Blob([yamlContent], { type: 'application/yaml' });
    const configFile = new File([blob], 'generated-config.yaml', { type: 'application/yaml' });
    formData.append('config_file', configFile);
  } else if (yamlFile.value) {
    formData.append('config_file', yamlFile.value)
  } else {
    isLoading.value = false
    return
  }

  formData.append('json_data', JSON.stringify(nodeData.value))

  try {
    const res = await request.post(
        '/check-format',
        formData,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': undefined
          }
        }
    )
    downloadFileFromResponse(res)
  } catch (error) {
    console.error('格式校验失败:', error)
    alert('❌ 格式校验失败：' + (error.response?.data?.detail || error.message))
  } finally {
    isLoading.value = false
  }
}

// 执行自动格式化
const callApplyFormatApi = async () => {
  if (!isFileLoaded.value) return

  // 低置信节点（建议人工复核）会被按当前标签强制执行，先提醒用户
  const reviewCount = visibleNodeData.value.filter(n => n.needs_review).length
  if (reviewCount > 0) {
    const ok = window.confirm(
        `有 ${reviewCount} 个节点置信度低于阈值（已标记“建议人工复核”），` +
        '格式化将按当前标签强制应用。建议先核对并修正这些节点，是否仍要继续？'
    )
    if (!ok) return
  }

  const ok = window.confirm('此操作将根据当前标签生成新文档，是否继续？');
  if (!ok) return;

  isLoading.value = true
  loadingText.value = '正在生成格式化后的文档...'
  if (!docxFile.value) return

  const formData = new FormData()
  formData.append('docx_file', docxFile.value)

  // 优先使用生成的配置，否则使用选择的yaml文件
  if (currentConfig.value) {
    // 将生成的配置转换为YAML字符串
    const yaml = await import('js-yaml');
    const yamlContent = yaml.dump(currentConfig.value, { indent: 2, skipInvalid: true });
    const blob = new Blob([yamlContent], { type: 'application/yaml' });
    const configFile = new File([blob], 'generated-config.yaml', { type: 'application/yaml' });
    formData.append('config_file', configFile);
  } else if (yamlFile.value) {
    formData.append('config_file', yamlFile.value)
  } else {
    isLoading.value = false
    return
  }

  formData.append('json_data', JSON.stringify(nodeData.value))

  try {
    const res = await request.post(
        '/apply-format',
        formData,
        {
          headers: {
            'Accept': 'application/json',
            'Content-Type': undefined
          }
        })

    // 触发下载
    downloadFileFromResponse(res)
  } catch (error) {
    console.error('格式化失败:', error)
    alert('❌ 文档格式化失败：' + (error.response?.data?.detail || error.message))
  } finally {
    isLoading.value = false
  }
}

// 下载文件（浏览器原生方式）
async function downloadFileFromResponse(response) {
  try {
    isLoading.value = false;

    // 兼容后端不同的响应结构：直接返回数据 或 包在 data 字段里
    const payload = response?.data ?? response;
    if (!payload?.download_url) {
      alert("操作完成！");
      return;
    }

    const { download_url, final_filename } = payload;

    // 相对路径要拼接后端地址，否则 fetch 会打到前端 dev server 拿到 HTML
    const fullUrl = download_url.startsWith('http')
      ? download_url
      : `${backendBaseUrl.value}${download_url}`;

    const downloadResp = await fetch(fullUrl);
    if (!downloadResp.ok) throw new Error(`下载失败: ${downloadResp.status}`);
    const blob = await downloadResp.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = final_filename || 'document.docx';
    document.body.appendChild(a);
    a.click();
    URL.revokeObjectURL(url);
    document.body.removeChild(a);
  } catch (err) {
    console.error("下载失败", err);
    alert("处理完成！文件下载失败：" + err.message);
  }
}
// 核对所有标签（前端模拟：高亮低分+other）
const checkAllTags = () => {
  if (!isFileLoaded.value) return

  const errors = visibleNodeData.value
      .map((node, idx) => ({...node, idx}))
      .filter(n => checkTagError(n) || n.category === 'other')

  if (errors.length === 0) {
    alert('✅ 所有节点标签均通过阈值校验！')
  } else {
    let msg = `🔍 发现 ${errors.length} 个需关注的节点：\n\n`
    msg += errors.slice(0, 10).map(e => `• [${e.idx + 1}] ${e.category}: ${e.paragraph.substring(0, 50)}...`).join('\n')
    if (errors.length > 10) msg += `\n... 还有 ${errors.length - 10} 个`
    alert(msg)
  }
}

// 导出修改后的 JSON
const exportModifiedJson = () => {
  if (!isFileLoaded.value) return

  const dataStr = JSON.stringify(nodeData.value, null, 2)
  const blob = new Blob([dataStr], {type: 'application/json;charset=utf-8'})
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = 'modified_nodes.json'
  document.body.appendChild(a)
  a.click()
  URL.revokeObjectURL(url)
  document.body.removeChild(a)
}

// 触发导入 JSON 的文件选择
const importJsonFile = () => {
  jsonFileInput.value?.click()
}

// 处理 JSON 文件导入
const handleJsonImport = async (e) => {
  const file = e.target.files?.[0]
  if (!file) return
  try {
    const text = await file.text()
    const json = JSON.parse(text)
    if (!Array.isArray(json)) {
      alert('JSON 格式错误：期望一个节点数组')
      return
    }
    // 补全缺失的 id 字段；保留全部节点（含 figure_image），
    // 保证与文档段落 1:1 对应，避免后端按位置 zip 时错位
    nodeData.value = json.map((node, idx) => ({
      ...node,
      id: node.id ?? idx
    }))
    isFileLoaded.value = true
    selectedNodeIndex.value = -1
    alert(`导入成功：${nodeData.value.length} 个节点`)
  } catch (err) {
    console.error('导入 JSON 失败:', err)
    alert('导入失败：' + err.message)
  } finally {
    e.target.value = ''
  }
}

// 清除选中当数据重置
watch(isFileLoaded, (loaded) => {
  if (!loaded) selectedNodeIndex.value = -1
})
</script>

<style scoped>
.doc-tag-check-container { width: 100%; height: 100vh; display: flex; flex-direction: column; background-color: var(--ink); }
.header-bar { background-color: var(--paper); border-bottom: 1px solid var(--border); padding: 0.75rem 1.5rem; }
.header-content { display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem; max-width: 1280px; margin: 0 auto; width: 100%; }
.header-left { display: flex; flex-direction: column; gap: 0.25rem; }
.tool-title { font-size: 1.15rem; font-weight: 600; color: var(--text); margin: 0; font-family: 'Georgia, Times New Roman', serif; }
.stats-info { font-size: 0.8125rem; color: var(--text-muted); }
.stats-info span { font-weight: 600; color: var(--green); }
.header-right { display: flex; align-items: center; gap: 0.75rem; flex-wrap: wrap; }
.search-box { position: relative; display: flex; align-items: center; }
.search-input { padding: 6px 28px 6px 10px; font-size: 13px; border: 1px solid var(--border); border-radius: 6px; width: 180px; outline: none; background: var(--ink); color: var(--text); font-family: inherit; }
.search-input:focus { border-color: var(--brass); box-shadow: 0 0 0 2px rgba(184,153,59,.15); }
.search-icon { position: absolute; right: 8px; width: 16px; height: 16px; color: var(--text-muted); pointer-events: none; }
.btn { padding: 6px 12px; font-size: 12px; border-radius: 6px; border: 1px solid transparent; cursor: pointer; transition: all .2s; white-space: nowrap; font-family: inherit; }
.btn:disabled { opacity: 0.4; cursor: not-allowed; }
.primary-btn { background-color: var(--brass); color: var(--ink); font-weight: 600; }
.primary-btn:hover:not(:disabled) { background-color: var(--brass-dim); }
.secondary-btn { background-color: var(--surface); color: var(--text-secondary); border: 1px solid var(--border); }
.secondary-btn:hover:not(:disabled) { background-color: var(--raised); border-color: var(--border-hover); color: var(--text); }
.main-content { flex: 1; min-height: 0; display: grid; grid-template-columns: 1fr 400px; gap: 1.5rem; padding: 1.5rem; width: 100%; max-width: 1280px; margin: 0 auto; }
@media (max-width: 992px) { .main-content { grid-template-columns: 1fr; } }
</style>
