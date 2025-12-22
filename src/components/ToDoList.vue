<template>
  <div class="todo-list-container">
    <div class="header-section">
      <h1>Simple To-Do List</h1>
      <p class="subtitle">Demonstrates CRUD operations on SharePoint List Items.</p>
    </div>

    <!-- Error/Status Message -->
    <div v-if="error" class="error-message">
      {{ error }}
    </div>

    <!-- Add New Task -->
    <div class="add-task-row">
      <input
        v-model="newTaskTitle"
        @keyup.enter="addTask"
        type="text"
        placeholder="What needs to be done?"
        :disabled="loading"
        class="task-input"
      />
      <button @click="addTask" :disabled="!newTaskTitle || loading" class="add-btn">
        {{ loading ? '...' : 'Add' }}
      </button>
    </div>

    <!-- Task List -->
    <div class="tasks-wrapper">
      <div v-if="loading && tasks.length === 0" class="loading-state">
        Loading tasks...
      </div>

      <div v-else-if="tasks.length === 0" class="empty-state">
        No tasks found. Add one above!
      </div>

      <transition-group name="list" tag="ul" class="task-list">
        <li v-for="task in tasks" :key="task.Id" class="task-item" :class="{ completed: task.IsComplete }">
          <label class="checkbox-wrapper">
            <input
              type="checkbox"
              :checked="task.IsComplete"
              @change="toggleTask(task)"
            />
            <span class="checkmark"></span>
          </label>

          <span class="task-title">{{ task.Title }}</span>

          <button @click="deleteTask(task.Id)" class="delete-btn" title="Delete Task">
            ✕
          </button>
        </li>
      </transition-group>
    </div>
  </div>
</template>

<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { useSharePoint } from 'vue3-sharepoint-plugin'

interface TaskItem {
  Id: number
  Title: string
  IsComplete: boolean
}

// Config
const LIST_NAME = 'Tasks' // Assumes a generic 'Tasks' list exists or will be mocked

// State
const { createListItem, getListItems, updateListItem, deleteListItem } = useSharePoint()
const tasks = ref<TaskItem[]>([])
const newTaskTitle = ref('')
const loading = ref(false)
const error = ref('')

// Actions
const loadTasks = async () => {
  loading.value = true
  error.value = ''
  try {
    // We assume the list has 'Title' and 'IsComplete' (Boolean)
    // Note: Standard SP Task list uses 'Status' (Choice) or 'PercentComplete'.
    // For this demo, we assume a custom boolean or we map it in a real scenario.
    // In Mock, we just store what we send.
    const items = await getListItems<TaskItem>(LIST_NAME)
    // Sort by ID desc (newest first)
    tasks.value = items.sort((a, b) => b.Id - a.Id)
  } catch (e) {
    console.error(e)
    error.value = `Failed to load list '${LIST_NAME}'. Ensure it exists.`
  } finally {
    loading.value = false
  }
}

const addTask = async () => {
  if (!newTaskTitle.value) return
  loading.value = true
  try {
    const payload = {
      Title: newTaskTitle.value,
      IsComplete: false
    }
    const newItem = await createListItem<TaskItem>(LIST_NAME, payload)
    tasks.value.unshift(newItem)
    newTaskTitle.value = ''
  } catch (e) {
    console.error(e)
    error.value = 'Failed to create task.'
  } finally {
    loading.value = false
  }
}

const toggleTask = async (task: TaskItem) => {
  // Optimistic UI update
  const originalState = task.IsComplete
  task.IsComplete = !task.IsComplete

  try {
    await updateListItem(LIST_NAME, task.Id, {
      IsComplete: task.IsComplete
    })
  } catch (e) {
    console.error(e)
    // Revert
    task.IsComplete = originalState
    error.value = 'Failed to update task.'
  }
}

const deleteTask = async (id: number) => {
  if (!confirm('Are you sure?')) return

  // Optimistic UI removal
  const originalList = [...tasks.value]
  tasks.value = tasks.value.filter(t => t.Id !== id)

  try {
    await deleteListItem(LIST_NAME, id)
  } catch (e) {
    console.error(e)
    tasks.value = originalList
    error.value = 'Failed to delete task.'
  }
}

onMounted(() => {
  loadTasks()
})
</script>

<style scoped>
.todo-list-container {
  max-width: 600px;
  margin: 0 auto;
  font-family: 'Segoe UI', sans-serif;
  color: #333;
  padding: 20px;
  background: white;
  border-radius: 8px;
  box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
}

.header-section {
  text-align: center;
  margin-bottom: 30px;
}

.header-section h1 {
  margin: 0 0 5px 0;
  color: #2563eb;
}

.subtitle {
  margin: 0;
  color: #6b7280;
  font-size: 0.9rem;
}

.error-message {
  background-color: #fee2e2;
  color: #b91c1c;
  padding: 10px;
  border-radius: 6px;
  margin-bottom: 20px;
  text-align: center;
  font-size: 0.9rem;
}

/* ADD TASK ROW */
.add-task-row {
  display: flex;
  gap: 10px;
  margin-bottom: 20px;
}

.task-input {
  flex: 1;
  padding: 12px 15px;
  border: 1px solid #d1d5db;
  border-radius: 6px;
  font-size: 1rem;
  outline: none;
  transition: border-color 0.2s;
}

.task-input:focus {
  border-color: #2563eb;
  box-shadow: 0 0 0 2px rgba(37, 99, 235, 0.1);
}

.add-btn {
  padding: 0 24px;
  background-color: #2563eb;
  color: white;
  border: none;
  border-radius: 6px;
  font-weight: 600;
  cursor: pointer;
  transition: background-color 0.2s;
}

.add-btn:hover:not(:disabled) {
  background-color: #1d4ed8;
}

.add-btn:disabled {
  background-color: #93c5fd;
  cursor: not-allowed;
}

/* TASK LIST */
.task-list {
  list-style: none;
  padding: 0;
  margin: 0;
}

.task-item {
  display: flex;
  align-items: center;
  padding: 12px 10px;
  border-bottom: 1px solid #f3f4f6;
  transition: background-color 0.2s;
}

.task-item:last-child {
  border-bottom: none;
}

.task-item:hover {
  background-color: #f9fafb;
}

.task-item.completed .task-title {
  text-decoration: line-through;
  color: #9ca3af;
}

/* CHECKBOX */
.checkbox-wrapper {
  position: relative;
  display: inline-block;
  width: 20px;
  height: 20px;
  margin-right: 15px;
  cursor: pointer;
}

.checkbox-wrapper input {
  opacity: 0;
  width: 0;
  height: 0;
}

.checkmark {
  position: absolute;
  top: 0;
  left: 0;
  height: 20px;
  width: 20px;
  background-color: white;
  border: 2px solid #d1d5db;
  border-radius: 4px;
  transition: all 0.2s;
}

.checkbox-wrapper input:checked ~ .checkmark {
  background-color: #2563eb;
  border-color: #2563eb;
}

.checkmark:after {
  content: "";
  position: absolute;
  display: none;
}

.checkbox-wrapper input:checked ~ .checkmark:after {
  display: block;
}

.checkbox-wrapper .checkmark:after {
  left: 6px;
  top: 2px;
  width: 4px;
  height: 10px;
  border: solid white;
  border-width: 0 2px 2px 0;
  transform: rotate(45deg);
}

.task-title {
  flex: 1;
  font-size: 1rem;
  color: #374151;
  word-break: break-word;
}

.delete-btn {
  background: none;
  border: none;
  color: #d1d5db;
  font-size: 1.2rem;
  cursor: pointer;
  padding: 0 8px;
  transition: color 0.2s;
}

.delete-btn:hover {
  color: #ef4444;
}

.empty-state, .loading-state {
  text-align: center;
  padding: 40px;
  color: #9ca3af;
  font-style: italic;
}

/* Transitions */
.list-enter-active,
.list-leave-active {
  transition: all 0.3s ease;
}
.list-enter-from,
.list-leave-to {
  opacity: 0;
  transform: translateX(-30px);
}
</style>
