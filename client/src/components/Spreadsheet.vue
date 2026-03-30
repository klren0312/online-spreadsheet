<script setup lang="ts">
import { ref, reactive, onMounted, onUnmounted, nextTick, computed } from 'vue'
import * as Y from 'yjs'
import { WebsocketProvider } from 'y-websocket'

const props = defineProps<{
  documentId: string
}>()

const emit = defineEmits<{
  statusChange: [status: { connected: boolean; collaborators: any[] }]
}>()

const ROWS = 100
const COLS = 26

const colLetters = Array.from({ length: COLS }, (_, i) => String.fromCharCode(65 + i))

const ydoc = new Y.Doc()
const ycells = ydoc.getMap<Y.Map<string>>('cells')
const ymerges = ydoc.getArray<string>('merges')
let provider: WebsocketProvider | null = null

const cells = reactive(new Map<string, string>())
const merges = reactive<Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>>([])

const selectedCell = ref<{ row: number; col: number } | null>(null)
const selectionStart = ref<{ row: number; col: number } | null>(null)
const selectionEnd = ref<{ row: number; col: number } | null>(null)
const isSelecting = ref(false)

const editingCell = ref<{ row: number; col: number } | null>(null)
const editValue = ref('')

const otherCursors = reactive(new Map<number, { row: number; col: number; color: string; name: string }>())

const inputProxy = ref<HTMLDivElement | null>(null)

const getCellKey = (row: number, col: number) => `${row}-${col}`

const getCellValue = (row: number, col: number): string => {
  return cells.get(getCellKey(row, col)) || ''
}

const getMergeInfo = (row: number, col: number) => {
  return merges.find(m =>
    row >= m.startRow && row <= m.endRow &&
    col >= m.startCol && col <= m.endCol
  )
}

const isMergeStart = (row: number, col: number) => {
  const merge = getMergeInfo(row, col)
  return merge && merge.startRow === row && merge.startCol === col
}

const shouldHideCell = (row: number, col: number) => {
  const merge = getMergeInfo(row, col)
  if (!merge) return false
  return !isMergeStart(row, col)
}

const getRowspan = (row: number, col: number) => {
  const merge = getMergeInfo(row, col)
  if (!merge) return 1
  if (merge.startRow !== row || merge.startCol !== col) return 0
  return merge.endRow - merge.startRow + 1
}

const getColspan = (row: number, col: number) => {
  const merge = getMergeInfo(row, col)
  if (!merge) return 1
  if (merge.startRow !== row || merge.startCol !== col) return 0
  return merge.endCol - merge.startCol + 1
}

const selectionRange = computed(() => {
  if (!selectionStart.value || !selectionEnd.value) return null

  const startRow = Math.min(selectionStart.value.row, selectionEnd.value.row)
  const endRow = Math.max(selectionStart.value.row, selectionEnd.value.row)
  const startCol = Math.min(selectionStart.value.col, selectionEnd.value.col)
  const endCol = Math.max(selectionStart.value.col, selectionEnd.value.col)

  return { startRow, endRow, startCol, endCol }
})

const isInSelection = (row: number, col: number) => {
  if (!selectionRange.value) return false
  const { startRow, endRow, startCol, endCol } = selectionRange.value
  return row >= startRow && row <= endRow && col >= startCol && col <= endCol
}

const canMerge = computed(() => {
  if (!selectionRange.value) return false
  const { startRow, endRow, startCol, endCol } = selectionRange.value
  if (startRow === endRow && startCol === endCol) return false

  for (const merge of merges) {
    const overlap = !(merge.endRow < startRow || merge.startRow > endRow ||
      merge.endCol < startCol || merge.startCol > endCol)
    if (overlap) return false
  }

  return true
})

const canUnmerge = computed(() => {
  if (!selectedCell.value) return false
  return !!getMergeInfo(selectedCell.value.row, selectedCell.value.col)
})

const mergeCells = () => {
  if (!selectionRange.value) return

  const { startRow, endRow, startCol, endCol } = selectionRange.value

  const mergeKey = `${startRow}-${startCol}-${endRow}-${endCol}`

  const values: { row: number; col: number; value: string }[] = []
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) {
      const value = getCellValue(r, c)
      if (value) {
        values.push({ row: r, col: c, value })
      }
    }
  }

  ydoc.transact(() => {
    ymerges.push([mergeKey])

    const topLeftValue = getCellValue(startRow, startCol)
    if (!topLeftValue && values.length > 0) {
      const firstValue = values[0]
      setCellValue(startRow, startCol, firstValue.value)

      if (firstValue.row !== startRow || firstValue.col !== startCol) {
        setCellValue(firstValue.row, firstValue.col, '')
      }
    }

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        if (r !== startRow || c !== startCol) {
          const val = getCellValue(r, c)
          if (val) {
            setCellValue(r, c, '')
          }
        }
      }
    }
  })

  syncMergesFromYjs()
}

const unmergeCells = () => {
  if (!selectedCell.value) return

  const merge = getMergeInfo(selectedCell.value.row, selectedCell.value.col)
  if (!merge) return

  const mergeKey = `${merge.startRow}-${merge.startCol}-${merge.endRow}-${merge.endCol}`

  ydoc.transact(() => {
    for (let i = 0; i < ymerges.length; i++) {
      if (ymerges.get(i) === mergeKey) {
        ymerges.delete(i, 1)
        break
      }
    }
  })

  syncMergesFromYjs()
}

const updateCollaborators = () => {
  if (!provider) return

  const states = provider.awareness.getStates()
  const collabs: any[] = []
  otherCursors.clear()

  states.forEach((state, clientId) => {
    if (state.user) {
      collabs.push({
        id: state.user.id,
        name: state.user.name,
        color: state.user.color
      })

      if (state.cursor && clientId !== provider!.awareness.clientID) {
        otherCursors.set(clientId, {
          row: state.cursor.row,
          col: state.cursor.col,
          color: state.user.color,
          name: state.user.name
        })
      }
    }
  })

  emit('statusChange', {
    connected: provider.wsconnected,
    collaborators: collabs
  })
}

const syncCellsFromYjs = () => {
  const newCells = new Map<string, string>()
  ycells.forEach((cellMap, key) => {
    const value = cellMap.get('value') || ''
    newCells.set(key, value)
  })
  cells.clear()
  newCells.forEach((v, k) => cells.set(k, v))
}

const syncMergesFromYjs = () => {
  merges.length = 0
  for (let i = 0; i < ymerges.length; i++) {
    const key = ymerges.get(i)
    const [startRow, startCol, endRow, endCol] = key.split('-').map(Number)
    merges.push({ startRow, startCol, endRow, endCol })
  }
}

const setCellValue = (row: number, col: number, value: string) => {
  const key = getCellKey(row, col)

  ydoc.transact(() => {
    let cellMap = ycells.get(key)
    if (!cellMap) {
      cellMap = new Y.Map<string>()
      ycells.set(key, cellMap)
    }
    cellMap.set('value', value)
  })
}

const selectCell = (row: number, col: number) => {
  selectedCell.value = { row, col }
  selectionStart.value = { row, col }
  selectionEnd.value = { row, col }

  if (provider) {
    provider.awareness.setLocalStateField('cursor', { row, col })
  }
}

const handleMouseDown = (row: number, col: number, e: MouseEvent) => {
  if (editingCell.value) {
    finishEditing()
  }

  if (e.shiftKey && selectedCell.value) {
    selectionStart.value = { ...selectedCell.value }
    selectionEnd.value = { row, col }
  } else {
    selectCell(row, col)
    isSelecting.value = true
  }

  nextTick(() => {
    focusInputProxy()
  })
}

const handleMouseMove = (row: number, col: number) => {
  if (isSelecting.value) {
    selectionEnd.value = { row, col }
  }
}

const handleMouseUp = () => {
  isSelecting.value = false
  nextTick(() => {
    focusInputProxy()
  })
}

const handleCellDblClick = (row: number, col: number) => {
  startEditing(row, col)
}

const focusInputProxy = () => {
  nextTick(() => {
    inputProxy.value?.focus()
  })
}

const clearInputProxy = () => {
  if (inputProxy.value) {
    inputProxy.value.textContent = ''
  }
}

const startEditing = (row: number, col: number, initialValue?: string) => {
  const merge = getMergeInfo(row, col)
  if (merge) {
    row = merge.startRow
    col = merge.startCol
  }

  editingCell.value = { row, col }
  editValue.value = initialValue ?? getCellValue(row, col)
  selectCell(row, col)

  nextTick(() => {
    const input = document.querySelector('.cell-input') as HTMLInputElement
    if (input) {
      input.focus()
      if (initialValue !== undefined) {
        const len = initialValue.length
        input.setSelectionRange(len, len)
      }
    }
  })
}

const finishEditing = () => {
  if (editingCell.value) {
    const { row, col } = editingCell.value
    if (editValue.value !== getCellValue(row, col)) {
      setCellValue(row, col, editValue.value)
    }
    editingCell.value = null
    editValue.value = ''
    focusInputProxy()
  }
}

const handleInputProxyBeforeInput = (e: InputEvent) => {
  if (editingCell.value) return
  if (!selectedCell.value) return
  if (!e.data) return

  e.preventDefault()

  const { row, col } = selectedCell.value
  startEditing(row, col, e.data)

  clearInputProxy()
}

const handleInputProxyKeyDown = (e: KeyboardEvent) => {
  if (editingCell.value) return
  if (!selectedCell.value) return

  const { row, col } = selectedCell.value

  switch (e.key) {
    case 'ArrowUp':
      e.preventDefault()
      if (row > 1) selectCell(row - 1, col)
      break
    case 'ArrowDown':
      e.preventDefault()
      if (row < ROWS) selectCell(row + 1, col)
      break
    case 'ArrowLeft':
      e.preventDefault()
      if (col > 0) selectCell(row, col - 1)
      break
    case 'ArrowRight':
      e.preventDefault()
      if (col < COLS - 1) selectCell(row, col + 1)
      break
    case 'Enter':
      e.preventDefault()
      startEditing(row, col)
      break
    case 'Delete':
    case 'Backspace':
      e.preventDefault()
      if (selectionRange.value) {
        for (let r = selectionRange.value.startRow; r <= selectionRange.value.endRow; r++) {
          for (let c = selectionRange.value.startCol; c <= selectionRange.value.endCol; c++) {
            setCellValue(r, c, '')
          }
        }
      } else {
        setCellValue(row, col, '')
      }
      break
    case 'Tab':
      e.preventDefault()
      if (e.shiftKey) {
        if (col > 0) selectCell(row, col - 1)
      } else {
        if (col < COLS - 1) selectCell(row, col + 1)
      }
      break
  }
}

const handleEditingKeyDown = (e: KeyboardEvent) => {
  if (e.key === 'Enter') {
    e.preventDefault()
    const currentRow = editingCell.value!.row
    const currentCol = editingCell.value!.col
    finishEditing()
    if (currentRow < ROWS) {
      selectCell(currentRow + 1, currentCol)
    }
  } else if (e.key === 'Tab') {
    e.preventDefault()
    const currentCol = editingCell.value!.col
    const currentRow = editingCell.value!.row
    finishEditing()
    if (e.shiftKey) {
      if (currentCol > 0) selectCell(currentRow, currentCol - 1)
    } else {
      if (currentCol < COLS - 1) selectCell(currentRow, currentCol + 1)
    }
  } else if (e.key === 'Escape') {
    editingCell.value = null
    editValue.value = ''
    focusInputProxy()
  }
}

const getCellClass = (row: number, col: number) => {
  const classes = ['cell']

  if (selectedCell.value?.row === row && selectedCell.value?.col === col) {
    classes.push('selected')
  }

  if (isInSelection(row, col)) {
    classes.push('in-selection')
  }

  if (editingCell.value?.row === row && editingCell.value?.col === col) {
    classes.push('editing')
  }

  return classes.join(' ')
}

const getCellStyle = (row: number, col: number) => {
  const style: Record<string, string> = {}

  for (const [clientId, cursor] of otherCursors) {
    if (cursor.row === row && cursor.col === col) {
      style.boxShadow = `inset 0 0 0 2px ${cursor.color}`
    }
  }

  return style
}

const getCursorLabel = (row: number, col: number) => {
  for (const [clientId, cursor] of otherCursors) {
    if (cursor.row === row && cursor.col === col) {
      return { name: cursor.name, color: cursor.color }
    }
  }
  return null
}

onMounted(() => {
  const wsUrl = `ws://${window.location.hostname}:3001`
  provider = new WebsocketProvider(wsUrl, props.documentId, ydoc)

  const userId = Math.floor(Math.random() * 1000000)
  const userColors = ['#f44336', '#2196f3', '#4caf50', '#ff9800', '#9c27b0', '#00bcd4']
  const userColor = userColors[Math.floor(Math.random() * userColors.length)]
  const userName = `用户${userId % 1000}`

  provider.awareness.setLocalStateField('user', {
    id: userId,
    name: userName,
    color: userColor
  })

  provider.on('status', () => {
    updateCollaborators()
  })

  provider.awareness.on('change', () => {
    updateCollaborators()
  })

  ycells.observe(syncCellsFromYjs)
  ymerges.observe(syncMergesFromYjs)
  syncCellsFromYjs()
  syncMergesFromYjs()

  focusInputProxy()
  document.addEventListener('mouseup', handleMouseUp)
})

onUnmounted(() => {
  document.removeEventListener('mouseup', handleMouseUp)
  if (provider) {
    provider.disconnect()
    provider.destroy()
  }
  ydoc.destroy()
})
</script>

<template>
  <div class="spreadsheet-wrapper">
    <div
      ref="inputProxy"
      class="input-proxy"
      contenteditable="true"
      @beforeinput="handleInputProxyBeforeInput"
      @keydown="handleInputProxyKeyDown"
    ></div>
    <div class="toolbar">
      <button
        class="toolbar-btn"
        :class="{ disabled: !canMerge }"
        :disabled="!canMerge"
        @click="mergeCells"
        title="合并单元格"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M3 5v14h18V5H3zm16 12H5V7h14v10zM8 13h3v-2H8V8l-4 4 4 4v-3zm8-2h-3v2h3v3l4-4-4-4v3z"/>
        </svg>
        <span>合并单元格</span>
      </button>
      <button
        class="toolbar-btn"
        :class="{ disabled: !canUnmerge }"
        :disabled="!canUnmerge"
        @click="unmergeCells"
        title="取消合并"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M3 5v14h18V5H3zm16 12H5V7h14v10zM8 13h3v-2H8V8l-4 4 4 4v-3zm5-2h3v2h-3v3l-4-4 4-4v3z"/>
        </svg>
        <span>取消合并</span>
      </button>
      <div class="toolbar-divider"></div>
      <button
        class="toolbar-btn"
        title="清除选中内容"
        @click="() => { if(selectionRange) { for(let r = selectionRange.startRow; r <= selectionRange.endRow; r++) { for(let c = selectionRange.startCol; c <= selectionRange.endCol; c++) { setCellValue(r, c, '') } } } }"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
        </svg>
        <span>清除</span>
      </button>
    </div>
    <div class="spreadsheet" tabindex="0" @click="focusInputProxy" @focus="focusInputProxy">
      <div class="spreadsheet-container">
        <table class="spreadsheet-table">
          <thead>
            <tr>
              <th class="corner-cell"></th>
              <th v-for="col in colLetters" :key="col" class="col-header">
                {{ col }}
              </th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="row in ROWS" :key="row">
              <td class="row-header">{{ row }}</td>
              <template v-for="colIdx in COLS" :key="`${row}-${colIdx}`">
                <td
                  v-if="!shouldHideCell(row, colIdx)"
                  :class="getCellClass(row, colIdx)"
                  :style="getCellStyle(row, colIdx)"
                  :rowspan="getRowspan(row, colIdx) || undefined"
                  :colspan="getColspan(row, colIdx) || undefined"
                  @mousedown="handleMouseDown(row, colIdx, $event)"
                  @mousemove="handleMouseMove(row, colIdx)"
                  @dblclick="handleCellDblClick(row, colIdx)"
                >
                  <div
                    v-if="getCursorLabel(row, colIdx)"
                    class="cursor-label"
                    :style="{ backgroundColor: getCursorLabel(row, colIdx)?.color }"
                  >
                    {{ getCursorLabel(row, colIdx)?.name }}
                  </div>
                  <input
                    v-if="editingCell?.row === row && editingCell?.col === colIdx"
                    v-model="editValue"
                    class="cell-input"
                    @keydown="handleEditingKeyDown"
                    @blur="finishEditing"
                    @mousedown.stop
                  />
                  <span v-else class="cell-value">{{ getCellValue(row, colIdx) }}</span>
                </td>
              </template>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
</template>

<style scoped>
.spreadsheet-wrapper {
  display: flex;
  flex-direction: column;
  height: 100%;
  background: white;
  border-radius: 8px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
  overflow: hidden;
  position: relative;
}

.input-proxy {
  position: absolute;
  left: -100px;
  top: -100px;
  width: 1px;
  height: 1px;
  opacity: 0;
  overflow: hidden;
  font-size: 16px;
}

.toolbar {
  display: flex;
  align-items: center;
  gap: 4px;
  padding: 8px 12px;
  border-bottom: 1px solid var(--border-color);
  background: var(--header-bg);
}

.toolbar-btn {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 6px 12px;
  border: 1px solid var(--border-color);
  border-radius: 4px;
  background: white;
  cursor: pointer;
  font-size: 13px;
  color: #333;
  transition: all 0.15s;
}

.toolbar-btn:hover:not(.disabled) {
  background: #e8f0fe;
  border-color: var(--primary-color);
  color: var(--primary-color);
}

.toolbar-btn.disabled {
  opacity: 0.5;
  cursor: not-allowed;
}

.toolbar-btn svg {
  flex-shrink: 0;
}

.toolbar-divider {
  width: 1px;
  height: 24px;
  background: var(--border-color);
  margin: 0 8px;
}

.spreadsheet {
  flex: 1;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  outline: none;
}

.spreadsheet-container {
  flex: 1;
  overflow: auto;
  position: relative;
}

.spreadsheet-table {
  border-collapse: collapse;
  width: max-content;
  min-width: 100%;
  table-layout: fixed;
}

.corner-cell {
  width: 50px;
  min-width: 50px;
  height: 28px;
  background: var(--header-bg);
  border: 1px solid var(--border-color);
  position: sticky;
  top: 0;
  left: 0;
  z-index: 3;
}

.col-header {
  width: 100px;
  min-width: 100px;
  height: 28px;
  background: var(--header-bg);
  border: 1px solid var(--border-color);
  text-align: center;
  font-weight: 500;
  color: #666;
  position: sticky;
  top: 0;
  z-index: 2;
  user-select: none;
}

.row-header {
  width: 50px;
  min-width: 50px;
  height: var(--cell-height);
  background: var(--header-bg);
  border: 1px solid var(--border-color);
  text-align: center;
  font-weight: 500;
  color: #666;
  position: sticky;
  left: 0;
  z-index: 1;
  user-select: none;
}

.cell {
  width: 100px;
  min-width: 100px;
  height: var(--cell-height);
  border: 1px solid var(--border-color);
  padding: 0 4px;
  position: relative;
  cursor: cell;
  transition: background-color 0.1s;
  box-sizing: border-box;
  vertical-align: middle;
}

.cell:hover {
  background: var(--bg-hover);
}

.cell.selected {
  background: var(--selection-bg);
  outline: 2px solid var(--selection-border);
  outline-offset: -2px;
  z-index: 2;
}

.cell.in-selection {
  background: var(--selection-bg);
}

.cell.editing {
  padding: 0;
  outline: 2px solid var(--selection-border);
  outline-offset: -2px;
  z-index: 2;
}

.cell-value {
  display: block;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  line-height: var(--cell-height);
}

.cell-input {
  width: 100%;
  height: 100%;
  border: none;
  outline: none;
  padding: 0 4px;
  font-size: inherit;
  font-family: inherit;
  background: white;
  box-sizing: border-box;
}

.cursor-label {
  position: absolute;
  top: -18px;
  left: -2px;
  padding: 2px 6px;
  font-size: 11px;
  color: white;
  border-radius: 3px 3px 3px 0;
  white-space: nowrap;
  z-index: 10;
  pointer-events: none;
}
</style>
