<template>
  <div class="config-section">
    <div class="section-header" @click="toggleSection">
      <h2>{{ title }}</h2>
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
  margin-bottom: 18px;
  border: 1px solid #334155;
  border-radius: 10px;
  overflow: hidden;
  background-color: #1e293b;
  transition: border-color 0.2s;
}
.config-section:hover { border-color: #475569; }
.section-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 14px 20px;
  background-color: #0f172a;
  cursor: pointer;
  transition: background-color 0.2s;
}
.section-header:hover { background-color: #1e293b; }
.section-header h2 {
  margin: 0;
  font-size: 15px;
  font-weight: 600;
  color: #e2e8f0;
}
.toggle-arrow {
  color: #64748b;
  transition: transform 0.2s;
  flex-shrink: 0;
}
.toggle-arrow.expanded { transform: rotate(180deg); }
.section-content {
  padding: 18px 20px;
  border-top: 1px solid #334155;
}
.section-content h3 {
  margin-top: 0;
  margin-bottom: 14px;
  font-size: 14px;
  color: #94a3b8;
}
</style>
