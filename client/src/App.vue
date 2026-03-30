<script setup lang="ts">
import { ref } from 'vue'
import Spreadsheet from './components/Spreadsheet.vue'
import CollaboratorList from './components/CollaboratorList.vue'

const documentId = ref('default')
const connected = ref(false)
const collaborators = ref<any[]>([])

const onStatusChange = (status: { connected: boolean; collaborators: any[] }) => {
  connected.value = status.connected
  collaborators.value = status.collaborators
}
</script>

<template>
  <div class="app">
    <header class="header">
      <div class="logo">
        <svg viewBox="0 0 24 24" width="28" height="28">
          <path fill="currentColor" d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14zM7 7h2v2H7zm0 4h2v2H7zm0 4h2v2H7zm4-8h6v2h-6zm0 4h6v2h-6zm0 4h6v2h-6z"/>
        </svg>
        <span>在线协作表格</span>
      </div>
      <div class="header-right">
        <div class="connection-status" :class="{ connected }">
          <span class="status-dot"></span>
          {{ connected ? '已连接' : '连接中...' }}
        </div>
        <CollaboratorList :collaborators="collaborators" />
      </div>
    </header>
    <main class="main">
      <Spreadsheet :document-id="documentId" @status-change="onStatusChange" />
    </main>
  </div>
</template>

<style scoped>
.app {
  display: flex;
  flex-direction: column;
  height: 100vh;
  background: #f5f5f5;
}

.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 16px;
  height: 56px;
  background: white;
  border-bottom: 1px solid #e0e0e0;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.08);
}

.logo {
  display: flex;
  align-items: center;
  gap: 10px;
  font-size: 18px;
  font-weight: 600;
  color: #1a73e8;
}

.header-right {
  display: flex;
  align-items: center;
  gap: 16px;
}

.connection-status {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 13px;
  color: #666;
}

.status-dot {
  width: 8px;
  height: 8px;
  border-radius: 50%;
  background: #f44336;
}

.connection-status.connected .status-dot {
  background: #4caf50;
}

.main {
  flex: 1;
  padding: 16px;
  overflow: hidden;
}
</style>
