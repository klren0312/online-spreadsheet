<script setup lang="ts">
import { ref, reactive, onMounted, onUnmounted, nextTick, computed } from 'vue'
import * as Y from 'yjs'
import { WebsocketProvider } from 'y-websocket'

interface CellStyle {
  color?: string
  bgColor?: string
  fontSize?: number
  fontWeight?: string
  textDecoration?: string
  textAlign?: string
}

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

const cells = ref(new Map<string, string>())
const cellStyles = ref(new Map<string, CellStyle>())
const forceUpdate = ref(0)
const merges = reactive<Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>>([])

const fontSizes = [10, 12, 14, 16, 18, 20, 24, 28, 32]
const defaultFontSize = 14
const fontSizeDropdownOpen = ref(false)

// 待应用的颜色值（用于节流）
let pendingColorValue: { styleKey: string; value: string } | null = null
let colorThrottleTimer: ReturnType<typeof setTimeout> | null = null

const flushColorChange = () => {
  if (pendingColorValue) {
    applyStyle(pendingColorValue.styleKey, pendingColorValue.value)
    pendingColorValue = null
  }
}

const handleColorInput = (styleKey: string, value: string) => {
  pendingColorValue = { styleKey, value }

  if (!colorThrottleTimer) {
    colorThrottleTimer = setTimeout(() => {
      colorThrottleTimer = null
      flushColorChange()
    }, 150)
  }
}

const handleColorChange = (styleKey: string, value: string) => {
  if (colorThrottleTimer) {
    clearTimeout(colorThrottleTimer)
    colorThrottleTimer = null
  }
  pendingColorValue = null
  applyStyle(styleKey, value)
}

const selectedCell = ref<{ row: number; col: number } | null>(null)
const selectionStart = ref<{ row: number; col: number } | null>(null)
const selectionEnd = ref<{ row: number; col: number } | null>(null)
const isSelecting = ref(false)

const editingCell = ref<{ row: number; col: number } | null>(null)
const editValue = ref('')

const otherCursors = reactive(new Map<number, { row: number; col: number; color: string; name: string }>())

const inputProxy = ref<HTMLDivElement | null>(null)
const fontSizeSelect = ref<HTMLSelectElement | null>(null)

const clipboard = ref<string[][]>([])
const clipboardStartCell = ref<{ row: number; col: number } | null>(null)
const isCut = ref(false)

const contextMenu = ref<{ visible: boolean; x: number; y: number }>({
  visible: false,
  x: 0,
  y: 0
})

const getCellKey = (row: number, col: number) => `${row}-${col}`

const getCellValue = (row: number, col: number): string => {
  forceUpdate.value
  return cells.value.get(getCellKey(row, col)) || ''
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
  const newStyles = new Map<string, CellStyle>()
  ycells.forEach((cellMap, key) => {
    const value = cellMap.get('value') || ''
    newCells.set(key, value)

    const style: CellStyle = {}
    const color = cellMap.get('color')
    const bgColor = cellMap.get('bgColor')
    const fontSize = cellMap.get('fontSize')
    const fontWeight = cellMap.get('fontWeight')
    const textDecoration = cellMap.get('textDecoration')
    const textAlign = cellMap.get('textAlign')

    if (color) style.color = color
    if (bgColor) style.bgColor = bgColor
    if (fontSize) style.fontSize = Number(fontSize)
    if (fontWeight) style.fontWeight = fontWeight
    if (textDecoration) style.textDecoration = textDecoration
    if (textAlign) style.textAlign = textAlign

    if (Object.keys(style).length > 0) {
      newStyles.set(key, style)
    }
  })
  cells.value = newCells
  cellStyles.value = newStyles
  forceUpdate.value++
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

const setCellStyle = (row: number, col: number, styleKey: string, value: any) => {
  const key = getCellKey(row, col)

  ydoc.transact(() => {
    let cellMap = ycells.get(key)
    if (!cellMap) {
      cellMap = new Y.Map<string>()
      ycells.set(key, cellMap)
    }
    if (value) {
      cellMap.set(styleKey, String(value))
    } else {
      cellMap.delete(styleKey)
    }
  })
}

const applyStyle = (styleKey: string, value: any) => {
  if (!selectionRange.value) return

  const { startRow, endRow, startCol, endCol } = selectionRange.value

  ydoc.transact(() => {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        setCellStyle(r, c, styleKey, value)
      }
    }
  })
}

const handleFontSizeChange = (e: Event) => {
  const select = e.target as HTMLSelectElement
  const value = Number(select.value)
  setTimeout(() => {
    applyStyle('fontSize', value)
    nextTick(() => {
      fontSizeSelect.value?.focus()
    })
  }, 10)
}

const toggleFontSizeDropdown = () => {
  fontSizeDropdownOpen.value = !fontSizeDropdownOpen.value
}

const selectFontSize = (size: number) => {
  applyStyle('fontSize', size)
  fontSizeDropdownOpen.value = false
}

const applyStyleLater = (styleKey: string, value: any) => {
  setTimeout(() => {
    applyStyle(styleKey, value)
  }, 0)
}

const getCurrentCellStyle = (): CellStyle => {
  if (!selectedCell.value) return {}
  const key = getCellKey(selectedCell.value.row, selectedCell.value.col)
  return cellStyles.value.get(key) || {}
}

const toggleBold = () => {
  const current = getCurrentCellStyle()
  const newValue = current.fontWeight === 'bold' ? '' : 'bold'
  applyStyle('fontWeight', newValue)
}

const toggleUnderline = () => {
  const current = getCurrentCellStyle()
  const newValue = current.textDecoration === 'underline' ? '' : 'underline'
  applyStyle('textDecoration', newValue)
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
  if (e.button === 2) return

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

  hideContextMenu()

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

const handleContextMenu = (e: MouseEvent, row: number, col: number) => {
  e.preventDefault()

  if (!isInSelection(row, col)) {
    selectCell(row, col)
  }

  contextMenu.value = {
    visible: true,
    x: e.clientX,
    y: e.clientY
  }
}

const hideContextMenu = () => {
  contextMenu.value.visible = false
  fontSizeDropdownOpen.value = false
}

const handleContextAction = (action: string) => {
  switch (action) {
    case 'cut':
      cutSelection()
      break
    case 'copy':
      copySelection()
      break
    case 'paste':
      pasteSelection()
      break
    case 'delete':
      deleteSelection()
      break
  }
  hideContextMenu()
}

const copySelection = () => {
  if (!selectionRange.value) return

  const { startRow, endRow, startCol, endCol } = selectionRange.value
  const data: string[][] = []

  for (let r = startRow; r <= endRow; r++) {
    const row: string[] = []
    for (let c = startCol; c <= endCol; c++) {
      row.push(getCellValue(r, c))
    }
    data.push(row)
  }

  clipboard.value = data
  clipboardStartCell.value = { row: startRow, col: startCol }

  const text = data.map(row => row.join('\t')).join('\n')
  navigator.clipboard.writeText(text)

  isCut.value = false
}

const cutSelection = () => {
  copySelection()
  isCut.value = true
}

const pasteSelection = () => {
  if (!selectedCell.value) return
  if (clipboard.value.length === 0) return

  const { row: startRow, col: startCol } = selectedCell.value

  ydoc.transact(() => {
    for (let r = 0; r < clipboard.value.length; r++) {
      for (let c = 0; c < clipboard.value[r].length; c++) {
        const targetRow = startRow + r
        const targetCol = startCol + c

        if (targetRow <= ROWS && targetCol < COLS) {
          setCellValue(targetRow, targetCol, clipboard.value[r][c])
        }
      }
    }

    if (isCut.value && clipboardStartCell.value) {
      for (let r = 0; r < clipboard.value.length; r++) {
        for (let c = 0; c < clipboard.value[r].length; c++) {
          const sourceRow = clipboardStartCell.value.row + r
          const sourceCol = clipboardStartCell.value.col + c

          if (sourceRow !== startRow + r || sourceCol !== startCol + c) {
            if (sourceRow <= ROWS && sourceCol < COLS) {
              setCellValue(sourceRow, sourceCol, '')
            }
          }
        }
      }
      isCut.value = false
    }
  })
}

const deleteSelection = () => {
  if (selectionRange.value) {
    const { startRow, endRow, startCol, endCol } = selectionRange.value

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        setCellValue(r, c, '')
      }
    }
  } else if (selectedCell.value) {
    setCellValue(selectedCell.value.row, selectedCell.value.col, '')
  }
}

const selectAll = () => {
  selectionStart.value = { row: 1, col: 0 }
  selectionEnd.value = { row: ROWS, col: COLS - 1 }
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

  if (e.ctrlKey || e.metaKey) {
    switch (e.key.toLowerCase()) {
      case 'c':
        e.preventDefault()
        copySelection()
        return
      case 'x':
        e.preventDefault()
        cutSelection()
        return
      case 'v':
        e.preventDefault()
        pasteSelection()
        return
      case 'a':
        e.preventDefault()
        selectAll()
        return
      case 'b':
        e.preventDefault()
        toggleBold()
        return
      case 'u':
        e.preventDefault()
        toggleUnderline()
        return
    }
  }

  switch (e.key) {
    case 'ArrowUp':
      e.preventDefault()
      if (e.shiftKey) {
        if (selectionEnd.value && selectionEnd.value.row > 1) {
          selectionEnd.value = { row: selectionEnd.value.row - 1, col: selectionEnd.value.col }
        }
      } else {
        if (row > 1) selectCell(row - 1, col)
      }
      break
    case 'ArrowDown':
      e.preventDefault()
      if (e.shiftKey) {
        if (selectionEnd.value && selectionEnd.value.row < ROWS) {
          selectionEnd.value = { row: selectionEnd.value.row + 1, col: selectionEnd.value.col }
        }
      } else {
        if (row < ROWS) selectCell(row + 1, col)
      }
      break
    case 'ArrowLeft':
      e.preventDefault()
      if (e.shiftKey) {
        if (selectionEnd.value && selectionEnd.value.col > 0) {
          selectionEnd.value = { row: selectionEnd.value.row, col: selectionEnd.value.col - 1 }
        }
      } else {
        if (col > 0) selectCell(row, col - 1)
      }
      break
    case 'ArrowRight':
      e.preventDefault()
      if (e.shiftKey) {
        if (selectionEnd.value && selectionEnd.value.col < COLS - 1) {
          selectionEnd.value = { row: selectionEnd.value.row, col: selectionEnd.value.col + 1 }
        }
      } else {
        if (col < COLS - 1) selectCell(row, col + 1)
      }
      break
    case 'Enter':
      e.preventDefault()
      startEditing(row, col)
      break
    case 'Delete':
    case 'Backspace':
      e.preventDefault()
      deleteSelection()
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
  const key = getCellKey(row, col)
  const cellStyle = cellStyles.value.get(key)

  if (cellStyle) {
    if (cellStyle.color) style.color = cellStyle.color
    if (cellStyle.bgColor) style.backgroundColor = cellStyle.bgColor
    if (cellStyle.fontSize) style.fontSize = cellStyle.fontSize + 'px'
    if (cellStyle.fontWeight) style.fontWeight = cellStyle.fontWeight
    if (cellStyle.textDecoration) style.textDecoration = cellStyle.textDecoration
    if (cellStyle.textAlign) style.textAlign = cellStyle.textAlign
  }

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

  ycells.observeDeep(syncCellsFromYjs)
  ymerges.observeDeep(syncMergesFromYjs)
  syncCellsFromYjs()
  syncMergesFromYjs()

  focusInputProxy()
  document.addEventListener('mouseup', handleMouseUp)
  document.addEventListener('click', hideContextMenu)
})

onUnmounted(() => {
  document.removeEventListener('mouseup', handleMouseUp)
  document.removeEventListener('click', hideContextMenu)
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
        @click="copySelection"
        title="复制 (Ctrl+C)"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/>
        </svg>
      </button>
      <button
        class="toolbar-btn"
        @click="cutSelection"
        title="剪切 (Ctrl+X)"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M9.64 7.64c.23-.5.36-1.05.36-1.64 0-2.21-1.79-4-4-4S2 3.79 2 6s1.79 4 4 4c.59 0 1.14-.13 1.64-.36L10 12l-2.36 2.36C7.14 14.13 6.59 14 6 14c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4c0-.59-.13-1.14-.36-1.64L12 14l7 7h3v-1L9.64 7.64zM6 8c-1.1 0-2-.89-2-2s.9-2 2-2 2 .89 2 2-.9 2-2 2zm0 12c-1.1 0-2-.89-2-2s.9-2 2-2 2 .89 2 2-.9 2-2 2zm6-7.5c-.28 0-.5-.22-.5-.5s.22-.5.5-.5.5.22.5.5-.22.5-.5.5zM19 3l-6 6 2 2 7-7V3h-3z"/>
        </svg>
      </button>
      <button
        class="toolbar-btn"
        @click="pasteSelection"
        title="粘贴 (Ctrl+V)"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M19 2h-4.18C14.4.84 13.3 0 12 0c-1.3 0-2.4.84-2.82 2H5c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1-1-.45-1-1 .45-1 1-1zm7 18H5V4h2v3h10V4h2v16z"/>
        </svg>
      </button>
      <div class="toolbar-divider"></div>
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
        <span>合并</span>
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
        @click="deleteSelection"
        title="删除 (Delete)"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
        </svg>
        <span>清除</span>
      </button>
      <div class="toolbar-divider"></div>
      <!-- 粗体 -->
      <button
        class="toolbar-btn"
        :class="{ active: getCurrentCellStyle().fontWeight === 'bold' }"
        @click="toggleBold"
        title="粗体 (Ctrl+B)"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M15.6 10.79c.97-.67 1.65-1.77 1.65-2.79 0-2.26-1.75-4-4-4H7v14h7.04c2.09 0 3.71-1.7 3.71-3.79 0-1.52-.86-2.82-2.15-3.42zM10 6.5h3c.83 0 1.5.67 1.5 1.5s-.67 1.5-1.5 1.5h-3v-3zm3.5 9H10v-3h3.5c.83 0 1.5.67 1.5 1.5s-.67 1.5-1.5 1.5z"/>
        </svg>
      </button>
      <!-- 下划线 -->
      <button
        class="toolbar-btn"
        :class="{ active: getCurrentCellStyle().textDecoration === 'underline' }"
        @click="toggleUnderline"
        title="下划线 (Ctrl+U)"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M12 17c3.31 0 6-2.69 6-6V3h-2.5v8c0 1.93-1.57 3.5-3.5 3.5S8.5 12.93 8.5 11V3H6v8c0 3.31 2.69 6 6 6zm-7 2v2h14v-2H5z"/>
        </svg>
      </button>
      <div class="toolbar-divider"></div>
      <!-- 文字颜色 -->
      <div class="color-picker-wrapper">
        <button class="toolbar-btn color-btn" title="文字颜色">
          <svg viewBox="0 0 24 24" width="18" height="18">
            <path fill="currentColor" d="M11 2L5.5 16h2.25l1.12-3h6.25l1.13 3h2.25L13 2h-2zm-1.38 9L12 4.67 14.38 11H9.62z"/>
          </svg>
          <span class="color-indicator" :style="{ backgroundColor: getCurrentCellStyle().color || '#333' }"></span>
        </button>
        <input
          type="color"
          class="color-input"
          :value="getCurrentCellStyle().color || '#333333'"
          @input="handleColorInput('color', ($event.target as HTMLInputElement).value)"
          @change="handleColorChange('color', ($event.target as HTMLInputElement).value)"
        />
      </div>
      <!-- 背景颜色 -->
      <div class="color-picker-wrapper">
        <button class="toolbar-btn color-btn" title="背景颜色">
          <svg viewBox="0 0 24 24" width="18" height="18">
            <path fill="currentColor" d="M16.56 8.94L7.62 0 6.21 1.41l2.38 2.38-5.15 5.15c-.59.59-.59 1.54 0 2.12l5.5 5.5c.29.29.68.44 1.06.44s.77-.15 1.06-.44l5.5-5.5c.59-.58.59-1.53 0-2.12zM5.21 10L10 5.21 14.79 10H5.21zM19 11.5s-2 2.17-2 3.5c0 1.1.9 2 2 2s2-.9 2-2c0-1.33-2-3.5-2-3.5z"/>
          </svg>
          <span class="color-indicator" :style="{ backgroundColor: getCurrentCellStyle().bgColor || '#fff' }"></span>
        </button>
        <input
          type="color"
          class="color-input"
          :value="getCurrentCellStyle().bgColor || '#ffffff'"
          @input="handleColorInput('bgColor', ($event.target as HTMLInputElement).value)"
          @change="handleColorChange('bgColor', ($event.target as HTMLInputElement).value)"
        />
      </div>
      <div class="toolbar-divider"></div>
      <!-- 字号 -->
      <div class="font-size-dropdown">
        <button
          class="toolbar-btn font-size-btn"
          @click.stop="toggleFontSizeDropdown"
        >
          <span>{{ getCurrentCellStyle().fontSize || defaultFontSize }}</span>
          <svg viewBox="0 0 24 24" width="14" height="14">
            <path fill="currentColor" d="M7 10l5 5 5-5z"/>
          </svg>
        </button>
        <div v-if="fontSizeDropdownOpen" class="font-size-menu">
          <div
            v-for="size in fontSizes"
            :key="size"
            class="font-size-option"
            :class="{ active: getCurrentCellStyle().fontSize === size || (!getCurrentCellStyle().fontSize && size === defaultFontSize) }"
            @click.stop="selectFontSize(size)"
          >
            {{ size }}
          </div>
        </div>
      </div>
      <div class="toolbar-divider"></div>
      <!-- 对齐 -->
      <button
        class="toolbar-btn"
        :class="{ active: getCurrentCellStyle().textAlign === 'left' }"
        @click="applyStyle('textAlign', 'left')"
        title="左对齐"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M15 15H3v2h12v-2zm0-8H3v2h12V7zM3 13h18v-2H3v2zm0 8h18v-2H3v2zM3 3v2h18V3H3z"/>
        </svg>
      </button>
      <button
        class="toolbar-btn"
        :class="{ active: getCurrentCellStyle().textAlign === 'center' || !getCurrentCellStyle().textAlign }"
        @click="applyStyle('textAlign', 'center')"
        title="居中"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M7 15v2h10v-2H7zm-4 6h18v-2H3v2zm0-8h18v-2H3v2zm4-6v2h10V7H7zM3 3v2h18V3H3z"/>
        </svg>
      </button>
      <button
        class="toolbar-btn"
        :class="{ active: getCurrentCellStyle().textAlign === 'right' }"
        @click="applyStyle('textAlign', 'right')"
        title="右对齐"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M3 21h18v-2H3v2zm6-4h12v-2H9v2zm-6-4h18v-2H3v2zm6-4h12V7H9v2zM3 3v2h18V3H3z"/>
        </svg>
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
                  @contextmenu="handleContextMenu($event, row, colIdx)"
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
    <div
      v-if="contextMenu.visible"
      class="context-menu"
      :style="{ left: contextMenu.x + 'px', top: contextMenu.y + 'px' }"
      @mousedown.stop
    >
      <div class="context-menu-item" @mousedown="handleContextAction('cut')">
        <svg viewBox="0 0 24 24" width="16" height="16">
          <path fill="currentColor" d="M9.64 7.64c.23-.5.36-1.05.36-1.64 0-2.21-1.79-4-4-4S2 3.79 2 6s1.79 4 4 4c.59 0 1.14-.13 1.64-.36L10 12l-2.36 2.36C7.14 14.13 6.59 14 6 14c-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4c0-.59-.13-1.14-.36-1.64L12 14l7 7h3v-1L9.64 7.64zM6 8c-1.1 0-2-.89-2-2s.9-2 2-2 2 .89 2 2-.9 2-2 2zm0 12c-1.1 0-2-.89-2-2s.9-2 2-2 2 .89 2 2-.9 2-2 2zm6-7.5c-.28 0-.5-.22-.5-.5s.22-.5.5-.5.5.22.5.5-.22.5-.5.5zM19 3l-6 6 2 2 7-7V3h-3z"/>
        </svg>
        <span>剪切</span>
        <span class="shortcut">Ctrl+X</span>
      </div>
      <div class="context-menu-item" @mousedown="handleContextAction('copy')">
        <svg viewBox="0 0 24 24" width="16" height="16">
          <path fill="currentColor" d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/>
        </svg>
        <span>复制</span>
        <span class="shortcut">Ctrl+C</span>
      </div>
      <div class="context-menu-item" @mousedown="handleContextAction('paste')">
        <svg viewBox="0 0 24 24" width="16" height="16">
          <path fill="currentColor" d="M19 2h-4.18C14.4.84 13.3 0 12 0c-1.3 0-2.4.84-2.82 2H5c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1-1-.45-1-1 .45-1 1-1zm7 18H5V4h2v3h10V4h2v16z"/>
        </svg>
        <span>粘贴</span>
        <span class="shortcut">Ctrl+V</span>
      </div>
      <div class="context-menu-divider"></div>
      <div class="context-menu-item" @mousedown="handleContextAction('delete')">
        <svg viewBox="0 0 24 24" width="16" height="16">
          <path fill="currentColor" d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
        </svg>
        <span>清除内容</span>
        <span class="shortcut">Delete</span>
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
  padding: 6px 10px;
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

.toolbar-btn.active {
  background: #e8f0fe;
  border-color: var(--primary-color);
  color: var(--primary-color);
}

.color-picker-wrapper {
  position: relative;
  display: flex;
  align-items: center;
}

.color-btn {
  position: relative;
  padding: 6px 8px;
}

.color-indicator {
  position: absolute;
  bottom: 2px;
  left: 50%;
  transform: translateX(-50%);
  width: 14px;
  height: 3px;
  border-radius: 1px;
  border: 1px solid #ccc;
}

.color-input {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  opacity: 0;
  cursor: pointer;
}

.font-size-dropdown {
  position: relative;
}

.font-size-btn {
  min-width: 50px;
  justify-content: space-between;
}

.font-size-menu {
  position: absolute;
  top: 100%;
  left: 0;
  margin-top: 4px;
  background: white;
  border: 1px solid var(--border-color);
  border-radius: 4px;
  box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
  z-index: 100;
  min-width: 60px;
}

.font-size-option {
  padding: 8px 12px;
  cursor: pointer;
  font-size: 13px;
}

.font-size-option:hover {
  background: #f0f0f0;
}

.font-size-option.active {
  background: #e8f0fe;
  color: var(--primary-color);
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
  height: 100%;
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

.context-menu {
  position: fixed;
  background: white;
  border: 1px solid #e0e0e0;
  border-radius: 6px;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
  padding: 4px 0;
  min-width: 180px;
  z-index: 1000;
}

.context-menu-item {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 12px;
  cursor: pointer;
  font-size: 13px;
  color: #333;
}

.context-menu-item:hover {
  background: #f0f0f0;
}

.context-menu-item svg {
  flex-shrink: 0;
  color: #666;
}

.context-menu-item .shortcut {
  margin-left: auto;
  font-size: 11px;
  color: #999;
}

.context-menu-divider {
  height: 1px;
  background: #e0e0e0;
  margin: 4px 0;
}
</style>
