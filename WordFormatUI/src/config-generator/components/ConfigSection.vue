<template>
  <div class="config-section">
    <div class="section-header" @click="toggleSection">
      <div class="section-header-left">
        <span class="section-rule"></span>
        <h2>{{ title }}</h2>
      </div>
      <svg class="toggle-arrow" :class="{ expanded: isExpanded }" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div v-if="isExpanded" class="section-content">
      <slot></slot>
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue'
const props = defineProps({
  title: { type: String, required: true },
  initialExpanded: { type: Boolean, default: true }
})
const isExpanded = ref(props.initialExpanded)
const toggleSection = () => { isExpanded.value = !isExpanded.value }
</script>

<style scoped>
.config-section {
  margin-bottom: 12px;
  border: 1px solid var(--border);
  border-radius: 10px;
  overflow: hidden;
  background: var(--paper);
  transition: border-color .15s;
}
.config-section:hover { border-color: var(--border-hover); }

.section-header {
  display: flex; justify-content: space-between; align-items: center;
  padding: 14px 18px;
  background: var(--surface);
  cursor: pointer; user-select: none;
  transition: background .12s;
}
.section-header:hover { background: var(--raised); }

.section-header-left { display: flex; align-items: center; gap: 10px; }
.section-rule {
  width: 3px; height: 18px; border-radius: 2px;
  background: var(--brass); flex-shrink: 0;
}

.section-header h2 {
  margin: 0; font-size: 13px; font-weight: 600; color: var(--text); letter-spacing: -0.01em;
}

.toggle-arrow {
  color: var(--text-muted); transition: transform .2s cubic-bezier(.4,0,.2,1); flex-shrink: 0;
}
.toggle-arrow.expanded { transform: rotate(180deg); }

.section-content {
  padding: 16px 18px;
  border-top: 1px solid var(--border);
}
</style>
