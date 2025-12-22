<template>
  <nav class="examples-nav">
    <div class="nav-brand">SharePoint Plugin Examples</div>
    <div class="nav-items">
      <button
        v-for="item in items"
        :key="item.id"
        class="nav-btn"
        :class="{ active: modelValue === item.id }"
        @click="$emit('update:modelValue', item.id)"
      >
        {{ item.label }}
      </button>
    </div>
  </nav>
</template>

<script setup lang="ts">
interface NavItem {
  id: string
  label: string
}

const items: NavItem[] = [
  { id: 'search', label: 'CV Search' },
  { id: 'upload', label: 'Drop Off Library' },
  { id: 'todo', label: 'To-Do List (CRUD)' },
  { id: 'mydocs', label: 'My Documents' },
  { id: 'myteam', label: 'Team Documents' },
  { id: 'system', label: 'System Info' }
]

defineProps<{
  modelValue: string
}>()

defineEmits<{
  (e: 'update:modelValue', value: string): void
}>()
</script>

<style scoped>
.examples-nav {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 1rem 2rem;
  background-color: #fff;
  border-bottom: 1px solid #e5e7eb;
  margin-bottom: 2rem;
  box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
}

.nav-brand {
  font-weight: 600;
  font-size: 1.125rem;
  color: #111827;
}

.nav-items {
  display: flex;
  gap: 0.5rem;
}

.nav-btn {
  padding: 0.5rem 1rem;
  border-radius: 0.375rem;
  border: 1px solid transparent;
  background-color: transparent;
  color: #4b5563;
  font-weight: 500;
  font-size: 0.875rem;
  cursor: pointer;
  transition: all 0.2s;
}

.nav-btn:hover {
  background-color: #f3f4f6;
  color: #111827;
}

.nav-btn.active {
  background-color: #eff6ff;
  color: #2563eb;
  border-color: #dbeafe;
}
</style>
