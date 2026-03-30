<script setup lang="ts">
import { Collaborator } from '../composables/useCollaboration'

defineProps<{
  collaborators: Collaborator[]
}>()
</script>

<template>
  <div class="collaborator-list">
    <div
      v-for="collab in collaborators.slice(0, 5)"
      :key="collab.id"
      class="avatar"
      :style="{ backgroundColor: collab.color }"
      :title="collab.name"
    >
      {{ collab.name.charAt(0) }}
    </div>
    <div v-if="collaborators.length > 5" class="avatar more">
      +{{ collaborators.length - 5 }}
    </div>
  </div>
</template>

<style scoped>
.collaborator-list {
  display: flex;
  align-items: center;
  gap: 4px;
}

.avatar {
  width: 32px;
  height: 32px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  color: white;
  font-size: 13px;
  font-weight: 500;
  border: 2px solid white;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.15);
  margin-left: -4px;
}

.avatar:first-child {
  margin-left: 0;
}

.avatar.more {
  background: #9e9e9e;
  font-size: 11px;
}
</style>
