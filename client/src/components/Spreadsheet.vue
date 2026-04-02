<script setup lang="ts">
import { ref, reactive, onMounted, onUnmounted, nextTick, computed } from 'vue'
import * as Y from 'yjs'
import { WebsocketProvider } from 'y-websocket'
import * as XLSX from 'xlsx-js-style'
import ExcelJS from 'exceljs'

interface CellStyle {
  color?: string
  bgColor?: string
  fontSize?: number
  fontWeight?: string
  textDecoration?: string
  textAlign?: string
  image?: string  // base64 图片数据
  dropdown?: string  // 下拉配置 key
}

const props = defineProps<{
  documentId: string
}>()

const emit = defineEmits<{
  statusChange: [status: { connected: boolean; collaborators: any[] }]
}>()

const ROWS = 1000  // 支持更多行
const COLS = 52   // 支持更多列 (A-Z, AA-AZ)

const DEFAULT_COL_WIDTH = 100
const DEFAULT_ROW_HEIGHT = 28

// 虚拟滚动配置
const BUFFER_ROWS = 5  // 上下缓冲行数
const BUFFER_COLS = 2  // 左右缓冲列数
const visibleStartRow = ref(1)
const visibleEndRow = ref(50)
const visibleStartCol = ref(0)
const visibleEndCol = ref(26)
const spreadsheetContainer = ref<HTMLElement | null>(null)

// 生成列字母 A-Z, AA-AZ, BA-BZ, ...
const colLetters = (() => {
  const letters: string[] = []
  for (let i = 0; i < COLS; i++) {
    if (i < 26) {
      letters.push(String.fromCharCode(65 + i))
    } else {
      const first = Math.floor((i - 26) / 26)
      const second = (i - 26) % 26
      letters.push(String.fromCharCode(65 + first) + String.fromCharCode(65 + second))
    }
  }
  return letters
})()

const ydoc = new Y.Doc()
const ycells = ydoc.getMap<Y.Map<string>>('cells')
const ymerges = ydoc.getArray<string>('merges')
const ycolWidths = ydoc.getMap<number>('colWidths')
const yrowHeights = ydoc.getMap<number>('rowHeights')
const ydropdowns = ydoc.getMap<string>('dropdowns')  // 下拉配置
let provider: WebsocketProvider | null = null

const cells = ref(new Map<string, string>())
const cellStyles = ref(new Map<string, CellStyle>())
const forceUpdate = ref(0)
const merges = reactive<Array<{ startRow: number; startCol: number; endRow: number; endCol: number }>>([])
const colWidths = ref(new Map<number, number>())
const rowHeights = ref(new Map<number, number>())
const dropdowns = ref(new Map<string, string[]>())  // key: "row,col" or "row" for row, "col" for column

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

// 列宽/行高调整
const isResizingCol = ref(false)
const resizingCol = ref<number | null>(null)
const resizingStartX = ref(0)
const resizingStartWidth = ref(0)

const isResizingRow = ref(false)
const resizingRow = ref<number | null>(null)
const resizingStartY = ref(0)
const resizingStartHeight = ref(0)

// 虚拟滚动：计算可见区域
const updateVisibleRange = () => {
  const container = spreadsheetContainer.value
  if (!container) return

  const scrollTop = container.scrollTop
  const scrollLeft = container.scrollLeft
  const clientHeight = container.clientHeight
  const clientWidth = container.clientWidth

  // 计算可见行范围
  let rowHeightSum = 0
  let startRow = 1
  for (let r = 1; r <= ROWS; r++) {
    rowHeightSum += getRowHeight(r)
    if (rowHeightSum > scrollTop) {
      startRow = r
      break
    }
  }

  let endRow = startRow
  rowHeightSum = 0
  for (let r = startRow; r <= ROWS; r++) {
    rowHeightSum += getRowHeight(r)
    if (rowHeightSum > clientHeight + scrollTop) {
      endRow = r
      break
    }
    endRow = r
  }

  // 计算可见列范围
  let colWidthSum = 50 // 行头宽度
  let startCol = 0
  for (let c = 0; c < COLS; c++) {
    colWidthSum += getColWidth(c)
    if (colWidthSum > scrollLeft) {
      startCol = c
      break
    }
  }

  let endCol = startCol
  colWidthSum = 50
  for (let c = 0; c < COLS; c++) {
    colWidthSum += getColWidth(c)
    if (colWidthSum > clientWidth + scrollLeft) {
      endCol = c
      break
    }
    endCol = c
  }

  // 添加缓冲区
  visibleStartRow.value = Math.max(1, startRow - BUFFER_ROWS)
  visibleEndRow.value = Math.min(ROWS, endRow + BUFFER_ROWS)
  visibleStartCol.value = Math.max(0, startCol - BUFFER_COLS)
  visibleEndCol.value = Math.min(COLS - 1, endCol + BUFFER_COLS)
}

// 计算表格总高度
const getTotalHeight = computed(() => {
  let height = 0
  for (let r = 1; r <= ROWS; r++) {
    height += getRowHeight(r)
  }
  return height
})

// 计算表格总宽度
const getTotalWidth = computed(() => {
  let width = 50 // 行头宽度
  for (let c = 0; c < COLS; c++) {
    width += getColWidth(c)
  }
  return width
})

// 计算上方占位高度
const getTopPadding = computed(() => {
  let height = 0
  for (let r = 1; r < visibleStartRow.value; r++) {
    height += getRowHeight(r)
  }
  return height
})

// 计算左侧占位宽度
const getLeftPadding = computed(() => {
  let width = 0
  for (let c = 0; c < visibleStartCol.value; c++) {
    width += getColWidth(c)
  }
  return width
})

// 获取可见行范围
const visibleRows = computed(() => {
  const rows: number[] = []
  for (let r = visibleStartRow.value; r <= visibleEndRow.value; r++) {
    rows.push(r)
  }
  return rows
})

// 获取可见列范围
const visibleCols = computed(() => {
  const cols: number[] = []
  for (let c = visibleStartCol.value; c <= visibleEndCol.value; c++) {
    cols.push(c)
  }
  return cols
})

// 节流的滚动处理
let scrollThrottleTimer: ReturnType<typeof setTimeout> | null = null
const handleScroll = () => {
  if (scrollThrottleTimer) return
  scrollThrottleTimer = setTimeout(() => {
    scrollThrottleTimer = null
    updateVisibleRange()
  }, 16) // 约 60fps
}

const handleColResizeStart = (e: MouseEvent, col: number) => {
  e.preventDefault()
  e.stopPropagation()
  isResizingCol.value = true
  resizingCol.value = col
  resizingStartX.value = e.clientX
  resizingStartWidth.value = getColWidth(col)
  document.addEventListener('mousemove', handleColResizing)
  document.addEventListener('mouseup', handleColResizeEnd)
}

const handleColResizing = (e: MouseEvent) => {
  if (!isResizingCol.value || resizingCol.value === null) return
  const delta = e.clientX - resizingStartX.value
  const newWidth = Math.max(30, resizingStartWidth.value + delta)
  setColWidth(resizingCol.value, newWidth)
}

const handleColResizeEnd = () => {
  isResizingCol.value = false
  resizingCol.value = null
  document.removeEventListener('mousemove', handleColResizing)
  document.removeEventListener('mouseup', handleColResizeEnd)
}

const handleRowResizeStart = (e: MouseEvent, row: number) => {
  e.preventDefault()
  e.stopPropagation()
  isResizingRow.value = true
  resizingRow.value = row
  resizingStartY.value = e.clientY
  resizingStartHeight.value = getRowHeight(row)
  document.addEventListener('mousemove', handleRowResizing)
  document.addEventListener('mouseup', handleRowResizeEnd)
}

const handleRowResizing = (e: MouseEvent) => {
  if (!isResizingRow.value || resizingRow.value === null) return
  const delta = e.clientY - resizingStartY.value
  const newHeight = Math.max(20, resizingStartHeight.value + delta)
  setRowHeight(resizingRow.value, newHeight)
}

const handleRowResizeEnd = () => {
  isResizingRow.value = false
  resizingRow.value = null
  document.removeEventListener('mousemove', handleRowResizing)
  document.removeEventListener('mouseup', handleRowResizeEnd)
}

// 图片处理
const imageInput = ref<HTMLInputElement | null>(null)
const excelInput = ref<HTMLInputElement | null>(null)

const handleImageUpload = (e: Event) => {
  const input = e.target as HTMLInputElement
  if (!input.files || input.files.length === 0) return

  const file = input.files[0]
  if (!file.type.startsWith('image/')) return

  const reader = new FileReader()
  reader.onload = () => {
    if (selectedCell.value && typeof reader.result === 'string') {
      applyStyle('image', reader.result)
    }
  }
  reader.readAsDataURL(file)

  // 清空 input 以便重复选择同一文件
  input.value = ''
}

const openImageUpload = () => {
  if (!selectedCell.value) return
  imageInput.value?.click()
}

// 导入 Excel
const triggerImportExcel = () => {
  excelInput.value?.click()
}

const handleImportExcel = (e: Event) => {
  const input = e.target as HTMLInputElement
  if (!input.files || input.files.length === 0) return

  const file = input.files[0]
  const reader = new FileReader()

  reader.onload = (event) => {
    try {
      const data = event.target?.result
      const workbook = XLSX.read(data, { type: 'binary' })

      // 读取第一个工作表
      const sheetName = workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]

      // 转换为二维数组
      const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as string[][]

      // 填充到表格中
      ydoc.transact(() => {
        for (let r = 0; r < jsonData.length && r < ROWS; r++) {
          const row = jsonData[r]
          if (!Array.isArray(row)) continue

          for (let c = 0; c < row.length && c < COLS; c++) {
            const value = row[c]
            if (value !== undefined && value !== null) {
              setCellValue(r + 1, c, String(value))
            }
          }
        }
      })
    } catch (err) {
      console.error('导入 Excel 失败:', err)
      alert('导入 Excel 失败，请检查文件格式')
    }
  }

  reader.readAsBinaryString(file)

  // 清空 input 以便重复选择同一文件
  input.value = ''
}

// 导出 Excel
// 将列索引转换为 Excel 列名 (0->A, 1->B, ...)
const colToExcelName = (col: number): string => {
  let name = ''
  let c = col
  while (c >= 0) {
    name = String.fromCharCode(65 + (c % 26)) + name
    c = Math.floor(c / 26) - 1
  }
  return name
}

// 将 (row, col) 转换为 Excel 单元格地址
const cellAddress = (row: number, col: number): string => {
  return colToExcelName(col) + row
}

const exportExcel = async () => {
  // 找出有数据的范围
  let maxRow = 0
  let maxCol = 0

  for (let r = 1; r <= ROWS; r++) {
    for (let c = 0; c < COLS; c++) {
      const value = getCellValue(r, c)
      const style = cellStyles.value.get(getCellKey(r, c))
      if (value || style?.image) {
        maxRow = Math.max(maxRow, r)
        maxCol = Math.max(maxCol, c)
      }
    }
  }

  // 如果没有数据，提示用户
  if (maxRow === 0) {
    alert('表格中没有数据可导出')
    return
  }

  // 使用 ExcelJS 创建工作簿
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('Sheet1')

  // 设置列宽
  for (let c = 0; c <= maxCol; c++) {
    worksheet.getColumn(c + 1).width = Math.max(10, Math.floor(getColWidth(c) / 8))
  }

  // 添加数据和样式
  for (let r = 1; r <= maxRow; r++) {
    const row = worksheet.getRow(r)
    row.height = getRowHeight(r) * 0.75

    for (let c = 0; c <= maxCol; c++) {
      const cell = row.getCell(c + 1)
      cell.value = getCellValue(r, c)

      const style = cellStyles.value.get(getCellKey(r, c))
      if (style) {
        // 字体样式
        cell.font = {
          bold: style.fontWeight === 'bold',
          underline: style.textDecoration === 'underline' ? 'single' : undefined,
          color: style.color ? { argb: 'FF' + style.color.replace('#', '') } : undefined,
          size: style.fontSize || 11
        }

        // 背景颜色
        if (style.bgColor) {
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FF' + style.bgColor.replace('#', '') }
          }
        }

        // 对齐
        if (style.textAlign) {
          cell.alignment = {
            horizontal: style.textAlign as 'left' | 'center' | 'right'
          }
        }
      }
    }
  }

  // 合并单元格
  for (const merge of merges) {
    if (merge.startRow <= maxRow && merge.startCol <= maxCol) {
      worksheet.mergeCells(
        merge.startRow,
        merge.startCol + 1,
        Math.min(merge.endRow, maxRow),
        Math.min(merge.endCol, maxCol) + 1
      )
    }
  }

  // 添加图片
  for (let r = 1; r <= maxRow; r++) {
    for (let c = 0; c <= maxCol; c++) {
      const style = cellStyles.value.get(getCellKey(r, c))
      if (style?.image) {
        const imageData = style.image
        const matches = imageData.match(/^data:image\/(\w+);base64,(.+)$/)
        if (matches) {
          const ext = matches[1] as 'png' | 'jpeg' | 'gif' | 'bmp' | 'webp'
          const base64Data = matches[2]

          // 将 base64 转换为 ArrayBuffer
          const binaryString = atob(base64Data)
          const bytes = new Uint8Array(binaryString.length)
          for (let i = 0; i < binaryString.length; i++) {
            bytes[i] = binaryString.charCodeAt(i)
          }

          // 添加图片到工作簿
          const imageId = workbook.addImage({
            buffer: bytes.buffer,
            extension: ext
          })

          // 添加图片到工作表
          worksheet.addImage(imageId, {
            tl: { col: c, row: r - 1 },
            br: { col: c + 1, row: r },
            editAs: 'oneCell'
          })
        }
      }
    }
  }

  // 导出文件
  const fileName = `spreadsheet_${new Date().toISOString().slice(0, 10)}.xlsx`
  const buffer = await workbook.xlsx.writeBuffer()

  // 下载文件
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  link.click()
  URL.revokeObjectURL(url)
}

const handleDocumentKeyDown = (e: KeyboardEvent) => {
  if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'v') {
    e.preventDefault()
    handleCtrlV()
  }
}

const handleCtrlV = async () => {
  if (!selectedCell.value) return

  try {
    // 尝试读取图片
    const items = await navigator.clipboard.read()
    for (const item of items) {
      for (const type of item.types) {
        if (type.startsWith('image/')) {
          const blob = await item.getType(type)
          const reader = new FileReader()
          reader.onload = () => {
            if (typeof reader.result === 'string') {
              applyStyle('image', reader.result)
            }
          }
          reader.readAsDataURL(blob)
          return
        }
      }
    }
  } catch (err) {
    // Clipboard API 不可用
  }

  // 尝试读取文字
  try {
    const text = await navigator.clipboard.readText()
    if (text) {
      // 解析制表符分隔的文本
      const rows = text.split('\n').map(row => row.split('\t'))
      const { row: startRow, col: startCol } = selectedCell.value

      ydoc.transact(() => {
        for (let r = 0; r < rows.length; r++) {
          for (let c = 0; c < rows[r].length; c++) {
            const targetRow = startRow + r
            const targetCol = startCol + c
            if (targetRow <= ROWS && targetCol < COLS) {
              setCellValue(targetRow, targetCol, rows[r][c])
            }
          }
        }
      })
      return
    }
  } catch (err) {
    // 无法读取剪贴板文字
  }

  // 使用内部剪贴板
  pasteSelection()
}

const removeImage = () => {
  if (!selectedCell.value) return
  const { row, col } = selectedCell.value
  const key = getCellKey(row, col)

  ydoc.transact(() => {
    let cellMap = ycells.get(key)
    if (cellMap) {
      cellMap.delete('image')
    }
  })
}

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
    const image = cellMap.get('image')
    const dropdown = cellMap.get('dropdown')

    if (color) style.color = color
    if (bgColor) style.bgColor = bgColor
    if (fontSize) style.fontSize = Number(fontSize)
    if (fontWeight) style.fontWeight = fontWeight
    if (textDecoration) style.textDecoration = textDecoration
    if (textAlign) style.textAlign = textAlign
    if (image) style.image = image
    if (dropdown) style.dropdown = dropdown

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

const syncDimensionsFromYjs = () => {
  const newColWidths = new Map<number, number>()
  const newRowHeights = new Map<number, number>()

  ycolWidths.forEach((width, col) => {
    newColWidths.set(Number(col), width)
  })

  yrowHeights.forEach((height, row) => {
    newRowHeights.set(Number(row), height)
  })

  colWidths.value = newColWidths
  rowHeights.value = newRowHeights
}

const syncDropdownsFromYjs = () => {
  const newDropdowns = new Map<string, string[]>()
  ydropdowns.forEach((optionsJson, key) => {
    try {
      const options = JSON.parse(optionsJson)
      newDropdowns.set(key, options)
    } catch (e) {
      console.error('Failed to parse dropdown options:', e)
    }
  })
  dropdowns.value = newDropdowns
}

const getColWidth = (col: number): number => {
  return colWidths.value.get(col) || DEFAULT_COL_WIDTH
}

const getRowHeight = (row: number): number => {
  return rowHeights.value.get(row) || DEFAULT_ROW_HEIGHT
}

const setColWidth = (col: number, width: number) => {
  ydoc.transact(() => {
    ycolWidths.set(col, width)
  })
}

const setRowHeight = (row: number, height: number) => {
  ydoc.transact(() => {
    yrowHeights.set(row, height)
  })
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
  let startRow: number, endRow: number, startCol: number, endCol: number

  if (selectionRange.value) {
    const range = selectionRange.value
    startRow = range.startRow
    endRow = range.endRow
    startCol = range.startCol
    endCol = range.endCol
  } else if (selectedCell.value) {
    startRow = endRow = selectedCell.value.row
    startCol = endCol = selectedCell.value.col
  } else {
    return
  }

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

  // 确保选中的单元格在可见区域内
  scrollToCell(row, col)
}

// 滚动到指定单元格
const scrollToCell = (row: number, col: number) => {
  const container = spreadsheetContainer.value
  if (!container) return

  // 计算目标位置
  let targetTop = 0
  for (let r = 1; r < row; r++) {
    targetTop += getRowHeight(r)
  }

  let targetLeft = 50 // 行头宽度
  for (let c = 0; c < col; c++) {
    targetLeft += getColWidth(c)
  }

  const cellHeight = getRowHeight(row)
  const cellWidth = getColWidth(col)

  // 垂直滚动
  if (targetTop < container.scrollTop) {
    container.scrollTop = targetTop
  } else if (targetTop + cellHeight > container.scrollTop + container.clientHeight) {
    container.scrollTop = targetTop + cellHeight - container.clientHeight
  }

  // 水平滚动
  if (targetLeft < container.scrollLeft) {
    container.scrollLeft = targetLeft
  } else if (targetLeft + cellWidth > container.scrollLeft + container.clientWidth) {
    container.scrollLeft = targetLeft + cellWidth - container.clientWidth
  }
}

const selectRow = (row: number) => {
  selectCell(row, 0)
  selectionStart.value = { row, col: 0 }
  selectionEnd.value = { row, col: COLS - 1 }
}

const selectCol = (col: number) => {
  selectionStart.value = { row: 1, col }
  selectionEnd.value = { row: ROWS, col }
  selectedCell.value = { row: 1, col }

  if (provider) {
    provider.awareness.setLocalStateField('cursor', { row: 1, col })
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

const previewImage = ref<string | null>(null)

// 下拉配置弹窗
const showDropdownConfig = ref(false)
const dropdownConfigType = ref<'row' | 'col'>('row')
const dropdownConfigIndex = ref(0)
const dropdownOptions = ref('')

// 判断是否选中了整行或整列
const isEntireRowSelected = computed(() => {
  if (!selectionRange.value) return false
  return selectionRange.value.startCol === 0 &&
         selectionRange.value.endCol === COLS - 1 &&
         selectionRange.value.startRow === selectionRange.value.endRow
})

const isEntireColSelected = computed(() => {
  if (!selectionRange.value) return false
  return selectionRange.value.startRow === 1 &&
         selectionRange.value.endRow === ROWS &&
         selectionRange.value.startCol === selectionRange.value.endCol
})

const getSelectedRow = computed(() => {
  if (!selectionRange.value || !isEntireRowSelected.value) return null
  return selectionRange.value.startRow
})

const getSelectedCol = computed(() => {
  if (!selectionRange.value || !isEntireColSelected.value) return null
  return selectionRange.value.startCol
})

const canAddDropdown = computed(() => {
  return isEntireRowSelected.value || isEntireColSelected.value
})

const handleDropdownConfig = () => {
  if (!selectionRange.value) return

  // 判断是行还是列
  const startRow = selectionRange.value.startRow
  const endRow = selectionRange.value.endRow
  const startCol = selectionRange.value.startCol
  const endCol = selectionRange.value.endCol

  // 整行：起始列为0，结束列为COLS-1
  if (startCol === 0 && endCol === COLS - 1) {
    openDropdownConfig('row', startRow)
  }
  // 整列：起始行为1，结束行为ROWS
  else if (startRow === 1 && endRow === ROWS) {
    openDropdownConfig('col', startCol)
  }
}

const dropdownTextarea = ref<HTMLTextAreaElement | null>(null)

const openDropdownConfig = (type: 'row' | 'col', index: number) => {
  dropdownConfigType.value = type
  dropdownConfigIndex.value = index

  // 获取已有的配置
  const key = type === 'row' ? `row-${index}` : `col-${index}`
  const existingOptions = dropdowns.value.get(key) || []
  dropdownOptions.value = existingOptions.join('\n')

  showDropdownConfig.value = true

  // 聚焦输入框
  nextTick(() => {
    dropdownTextarea.value?.focus()
  })
}

const saveDropdownConfig = () => {
  const options = dropdownOptions.value
    .split('\n')
    .map(o => o.trim())
    .filter(o => o.length > 0)

  if (options.length === 0) {
    showDropdownConfig.value = false
    return
  }

  const key = dropdownConfigType.value === 'row'
    ? `row-${dropdownConfigIndex.value}`
    : `col-${dropdownConfigIndex.value}`

  ydoc.transact(() => {
    ydropdowns.set(key, JSON.stringify(options))
  })

  // 对该行/列的所有单元格应用下拉
  ydoc.transact(() => {
    if (dropdownConfigType.value === 'row') {
      for (let c = 0; c < COLS; c++) {
        setCellStyle(dropdownConfigIndex.value, c, 'dropdown', key)
      }
    } else {
      for (let r = 1; r <= ROWS; r++) {
        setCellStyle(r, dropdownConfigIndex.value, 'dropdown', key)
      }
    }
  })

  showDropdownConfig.value = false
}

const cancelDropdownConfig = () => {
  showDropdownConfig.value = false
}

const getCellDropdownOptions = (row: number, col: number): string[] | null => {
  const key = getCellKey(row, col)
  const style = cellStyles.value.get(key)
  if (style?.dropdown) {
    return dropdowns.value.get(style.dropdown) || null
  }

  // 检查行/列级别
  const rowKey = `row-${row}`
  if (dropdowns.value.has(rowKey)) {
    return dropdowns.value.get(rowKey) || null
  }

  const colKey = `col-${col}`
  if (dropdowns.value.has(colKey)) {
    return dropdowns.value.get(colKey) || null
  }

  return null
}

const handleCellDblClick = (row: number, col: number) => {
  const key = getCellKey(row, col)
  const style = cellStyles.value.get(key)

  // 如果有图片，显示预览
  if (style?.image) {
    previewImage.value = style.image
    return
  }

  // 否则进入编辑模式
  startEditing(row, col)
}

const closePreview = () => {
  previewImage.value = null
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

  // 复制图片到剪贴板
  copyImageToClipboard()

  isCut.value = false
}

const copyImageToClipboard = async () => {
  if (!selectionRange.value) return

  const { startRow, startCol } = selectionRange.value

  // 如果只有一个单元格包含图片，复制该图片
  if (startRow === selectionRange.value.endRow && startCol === selectionRange.value.endCol) {
    const key = getCellKey(startRow, startCol)
    const style = cellStyles.value.get(key)
    if (style?.image) {
      try {
        // 将 base64 转换为 blob
        const response = await fetch(style.image)
        const blob = await response.blob()
        await navigator.clipboard.write([
          new ClipboardItem({ [blob.type]: blob })
        ])
      } catch (err) {
        console.error('Failed to copy image:', err)
      }
    }
  }
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
        handleCtrlV()
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

  // 添加列宽和行高
  style.width = getColWidth(col) + 'px'
  style.minWidth = getColWidth(col) + 'px'
  style.height = getRowHeight(row) + 'px'
  style.minHeight = getRowHeight(row) + 'px'

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
  ycolWidths.observeDeep(syncDimensionsFromYjs)
  yrowHeights.observeDeep(syncDimensionsFromYjs)
  ydropdowns.observeDeep(syncDropdownsFromYjs)
  syncCellsFromYjs()
  syncMergesFromYjs()
  syncDimensionsFromYjs()
  syncDropdownsFromYjs()

  focusInputProxy()
  document.addEventListener('mouseup', handleMouseUp)
  document.addEventListener('click', hideContextMenu)
  document.addEventListener('keydown', handleDocumentKeyDown)

  // 初始化可见区域
  nextTick(() => {
    updateVisibleRange()
  })

  // 添加滚动监听
  if (spreadsheetContainer.value) {
    spreadsheetContainer.value.addEventListener('scroll', handleScroll)
  }
})

onUnmounted(() => {
  document.removeEventListener('mouseup', handleMouseUp)
  document.removeEventListener('click', hideContextMenu)
  document.removeEventListener('keydown', handleDocumentKeyDown)

  // 移除滚动监听
  if (spreadsheetContainer.value) {
    spreadsheetContainer.value.removeEventListener('scroll', handleScroll)
  }

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
        :class="{ active: !!getCurrentCellStyle().image }"
        @click="openImageUpload"
        title="插入图片"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M21 19V5c0-1.1-.9-2-2-2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2zM8.5 13.5l2.5 3.01L14.5 12l4.5 6H5l3.5-4.5z"/>
        </svg>
      </button>
      <button
        v-if="getCurrentCellStyle().image"
        class="toolbar-btn"
        @click="removeImage"
        title="删除图片"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
        </svg>
      </button>
      <div class="toolbar-divider"></div>
      <button
        class="toolbar-btn"
        :class="{ disabled: !canAddDropdown }"
        :disabled="!canAddDropdown"
        @click="handleDropdownConfig"
        title="配置下拉"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M3 17v2h6v-2H3zM3 5v2h10V5H3zm10 16v-2h8v-2h-8v-2h-2v6h2zM7 9v2H3v2h4v2h2V9H7zm14 4v-2H11v2h10zm-6-4h2V7h4V5h-4V3h-2v6z"/>
        </svg>
        <span>下拉</span>
      </button>
      <div class="toolbar-divider"></div>
      <button
        class="toolbar-btn"
        @click="triggerImportExcel"
        title="导入 Excel"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/>
        </svg>
        <span>导入</span>
      </button>
      <button
        class="toolbar-btn"
        @click="exportExcel"
        title="导出 Excel"
      >
        <svg viewBox="0 0 24 24" width="18" height="18">
          <path fill="currentColor" d="M19 12v7H5v-7H3v7c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2v-7h-2zm-6 .67l2.59-2.58L17 11.5l-5 5-5-5 1.41-1.41L11 12.67V3h2z"/>
        </svg>
        <span>导出</span>
      </button>
    </div>
    <input
      ref="imageInput"
      type="file"
      accept="image/*"
      style="display: none"
      @change="handleImageUpload"
    />
    <input
      ref="excelInput"
      type="file"
      accept=".xlsx,.xls"
      style="display: none"
      @change="handleImportExcel"
    />
    <div class="spreadsheet" tabindex="0" @click="focusInputProxy" @focus="focusInputProxy">
      <div
        ref="spreadsheetContainer"
        class="spreadsheet-container"
      >
        <table class="spreadsheet-table">
          <thead>
            <tr>
              <th class="corner-cell"></th>
              <th
                v-for="c in visibleCols"
                :key="c"
                class="col-header"
                :style="{ width: getColWidth(c) + 'px', minWidth: getColWidth(c) + 'px' }"
                @mousedown.stop="selectCol(c)"
              >
                {{ colLetters[c] }}
                <div
                  class="col-resize-handle"
                  @mousedown.stop="handleColResizeStart($event, c)"
                ></div>
              </th>
            </tr>
          </thead>
          <tbody>
            <!-- 上方占位 -->
            <tr v-if="getTopPadding > 0" :style="{ height: getTopPadding + 'px' }">
              <td :colspan="visibleCols.length + 1" style="padding: 0"></td>
            </tr>
            <!-- 可见行 -->
            <tr v-for="row in visibleRows" :key="row">
              <td
                class="row-header"
                :style="{ height: getRowHeight(row) + 'px', minHeight: getRowHeight(row) + 'px' }"
                @mousedown.stop="selectRow(row)"
              >
                {{ row }}
                <div
                  class="row-resize-handle"
                  @mousedown.stop="handleRowResizeStart($event, row)"
                ></div>
              </td>
              <template v-for="col in visibleCols" :key="`${row}-${col}`">
                <td
                  v-if="!shouldHideCell(row, col)"
                  :class="getCellClass(row, col)"
                  :style="getCellStyle(row, col)"
                  :rowspan="getRowspan(row, col) || undefined"
                  :colspan="getColspan(row, col) || undefined"
                  @mousedown="handleMouseDown(row, col, $event)"
                  @mousemove="handleMouseMove(row, col)"
                  @dblclick="handleCellDblClick(row, col)"
                  @contextmenu="handleContextMenu($event, row, col)"
                >
                  <div
                    v-if="getCursorLabel(row, col)"
                    class="cursor-label"
                    :style="{ backgroundColor: getCursorLabel(row, col)?.color }"
                  >
                    {{ getCursorLabel(row, col)?.name }}
                  </div>
                  <input
                    v-if="editingCell?.row === row && editingCell?.col === col"
                    v-model="editValue"
                    class="cell-input"
                    @keydown="handleEditingKeyDown"
                    @blur="finishEditing"
                    @mousedown.stop
                  />
                  <template v-else>
                    <img
                      v-if="cellStyles.get(getCellKey(row, col))?.image"
                      :src="cellStyles.get(getCellKey(row, col))?.image"
                      class="cell-image"
                    />
                    <select
                      v-else-if="getCellDropdownOptions(row, col)"
                      class="cell-dropdown"
                      :value="getCellValue(row, col)"
                      @change="setCellValue(row, col, ($event.target as HTMLSelectElement).value)"
                      @mousedown.stop
                      @mouseup.stop
                      @click.stop
                    >
                      <option value="">请选择</option>
                      <option
                        v-for="option in getCellDropdownOptions(row, col)"
                        :key="option"
                        :value="option"
                      >
                        {{ option }}
                      </option>
                    </select>
                    <span v-else class="cell-value">{{ getCellValue(row, col) }}</span>
                  </template>
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
    <!-- 图片预览弹窗 -->
    <div
      v-if="previewImage"
      class="image-preview-overlay"
      @click="closePreview"
      @mousedown.stop
      @mouseup.stop
    >
      <div class="image-preview-content" @click.stop @mouseup.stop>
        <img :src="previewImage" alt="预览图片" />
        <button class="preview-close" @click.stop="closePreview" @mouseup.stop>
          <svg viewBox="0 0 24 24" width="24" height="24">
            <path fill="currentColor" d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
          </svg>
        </button>
      </div>
    </div>
    <!-- 下拉配置弹窗 -->
    <div
      v-if="showDropdownConfig"
      class="dropdown-config-overlay"
      @click="cancelDropdownConfig"
      @mousedown.stop
      @mouseup.stop
    >
      <div class="dropdown-config-modal" @click.stop @mouseup.stop>
        <div class="modal-header">
          <h3>配置下拉选项 - {{ dropdownConfigType === 'row' ? '第 ' + dropdownConfigIndex + ' 行' : '列 ' + colLetters[dropdownConfigIndex] }}</h3>
          <button class="modal-close" @click.stop="cancelDropdownConfig" @mouseup.stop>
            <svg viewBox="0 0 24 24" width="20" height="20">
              <path fill="currentColor" d="M19 6.41L17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z"/>
            </svg>
          </button>
        </div>
        <div class="modal-body">
          <p>请输入下拉选项，每行一个选项：</p>
          <textarea
            ref="dropdownTextarea"
            v-model="dropdownOptions"
            placeholder="选项1&#10;选项2&#10;选项3"
            rows="8"
            @mousedown.stop
            @mouseup.stop
            @click.stop
          ></textarea>
        </div>
        <div class="modal-footer">
          <button class="btn-cancel" @click.stop="cancelDropdownConfig" @mouseup.stop>取消</button>
          <button class="btn-confirm" @click.stop="saveDropdownConfig" @mouseup.stop>确定</button>
        </div>
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
  position: relative;
}

.col-header .col-resize-handle {
  position: absolute;
  top: 0;
  bottom: 0;
  right: 0;
  width: 4px;
  cursor: ew-resize;
  background: transparent;
}

.col-header .col-resize-handle:hover {
  background: var(--primary-color);
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

.row-header .row-resize-handle {
  position: absolute;
  bottom: 0;
  left: 0;
  right: 0;
  height: 4px;
  cursor: ns-resize;
  background: transparent;
}

.row-header .row-resize-handle:hover {
  background: var(--primary-color);
}

.cell {
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

.cell-image {
  max-width: 100%;
  max-height: 100%;
  object-fit: contain;
}

.cell-dropdown {
  width: 100%;
  height: 100%;
  border: none;
  background: transparent;
  font-size: inherit;
  font-family: inherit;
  cursor: pointer;
  padding: 0 2px;
}

.cell-dropdown:focus {
  outline: none;
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

.image-preview-overlay {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(0, 0, 0, 0.8);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 9999;
}

.image-preview-content {
  position: relative;
  max-width: 90vw;
  max-height: 90vh;
}

.image-preview-content img {
  max-width: 100%;
  max-height: 90vh;
  object-fit: contain;
  border-radius: 4px;
}

.preview-close {
  position: absolute;
  top: -40px;
  right: 0;
  background: none;
  border: none;
  cursor: pointer;
  color: white;
  padding: 8px;
}

.preview-close:hover {
  color: #ccc;
}

.dropdown-config-overlay {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  bottom: 0;
  background: rgba(0, 0, 0, 0.5);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 9999;
}

.dropdown-config-modal {
  background: white;
  border-radius: 8px;
  width: 400px;
  max-width: 90vw;
  box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
}

.modal-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 16px 20px;
  border-bottom: 1px solid #e0e0e0;
}

.modal-header h3 {
  margin: 0;
  font-size: 16px;
  font-weight: 500;
}

.modal-close {
  background: none;
  border: none;
  cursor: pointer;
  color: #666;
  padding: 4px;
}

.modal-close:hover {
  color: #333;
}

.modal-body {
  padding: 20px;
}

.modal-body p {
  margin: 0 0 12px;
  color: #666;
  font-size: 14px;
}

.modal-body textarea {
  width: 100%;
  padding: 12px;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  font-size: 14px;
  resize: vertical;
  font-family: inherit;
  box-sizing: border-box;
}

.modal-body textarea:focus {
  outline: none;
  border-color: var(--primary-color);
}

.modal-footer {
  display: flex;
  justify-content: flex-end;
  gap: 12px;
  padding: 16px 20px;
  border-top: 1px solid #e0e0e0;
}

.btn-cancel,
.btn-confirm {
  padding: 8px 20px;
  border-radius: 4px;
  font-size: 14px;
  cursor: pointer;
}

.btn-cancel {
  background: white;
  border: 1px solid #e0e0e0;
  color: #666;
}

.btn-cancel:hover {
  background: #f5f5f5;
}

.btn-confirm {
  background: var(--primary-color);
  border: 1px solid var(--primary-color);
  color: white;
}

.btn-confirm:hover {
  background: #1976d2;
}
</style>
