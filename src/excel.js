import fs from 'fs'
import path from 'path'
import XLSX from 'xlsx'

export function writeCell({ filePath, sheetName, cell, value, outputPath }) {
  const resolved = path.resolve(filePath)
  if (!fs.existsSync(resolved)) {
    throw new Error(`Excel file not found at ${resolved}`)
  }
  const workbook = XLSX.readFile(resolved)
  const targetSheetName = sheetName || workbook.SheetNames[0]
  const worksheet = workbook.Sheets[targetSheetName]
  if (!worksheet) {
    throw new Error(`Sheet not found: ${targetSheetName}`)
  }
  const cellAddr = XLSX.utils.decode_cell(cell)
  const isNumeric = typeof value === 'number' || (typeof value === 'string' && value.trim() !== '' && !isNaN(Number(value)))
  worksheet[cell] = isNumeric ? { t: 'n', v: Number(value) } : { t: 's', v: String(value) }

  // Expand sheet range to include the written cell
  const currentRef = worksheet['!ref'] || cell
  const range = XLSX.utils.decode_range(currentRef)
  range.s.r = Math.min(range.s.r, cellAddr.r)
  range.s.c = Math.min(range.s.c, cellAddr.c)
  range.e.r = Math.max(range.e.r, cellAddr.r)
  range.e.c = Math.max(range.e.c, cellAddr.c)
  worksheet['!ref'] = XLSX.utils.encode_range(range)
  const outPath = outputPath ? path.resolve(outputPath) : resolved
  if (outputPath) {
    const dir = path.dirname(outPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  }
  XLSX.writeFile(workbook, outPath)
  return { ok: true, filePath: outPath, sheetName: targetSheetName, cell, value }
}

function getMergeTopLeftIfAny(worksheet, cell) {
  const merges = worksheet['!merges'] || []
  const addr = XLSX.utils.decode_cell(cell)
  for (const m of merges) {
    // m has s(start) and e(end) with r(row), c(col)
    if (addr.r >= m.s.r && addr.r <= m.e.r && addr.c >= m.s.c && addr.c <= m.e.c) {
      return XLSX.utils.encode_cell({ r: m.s.r, c: m.s.c })
    }
  }
  return cell
}

export function writeCellsBulk({ filePath, sheetName, writes, respectMerges = true, outputPath }) {
  const resolved = path.resolve(filePath)
  if (!fs.existsSync(resolved)) throw new Error(`Excel file not found at ${resolved}`)
  const workbook = XLSX.readFile(resolved)
  const targetSheetName = sheetName || workbook.SheetNames[0]
  const worksheet = workbook.Sheets[targetSheetName]
  if (!worksheet) throw new Error(`Sheet not found: ${targetSheetName}`)

  let currentRef = worksheet['!ref'] || 'A1'
  let range = XLSX.utils.decode_range(currentRef)
  const results = []

  for (const w of writes || []) {
    if (!w) continue
    let cell = String(w.cell || '').trim()
    if (!cell) continue
    if (respectMerges) {
      cell = getMergeTopLeftIfAny(worksheet, cell)
    }
    const cellAddr = XLSX.utils.decode_cell(cell)
    const isNumeric = typeof w.value === 'number' || (typeof w.value === 'string' && w.value.trim() !== '' && !isNaN(Number(w.value)))
    worksheet[cell] = isNumeric ? { t: 'n', v: Number(w.value) } : { t: 's', v: String(w.value ?? '') }

    // expand range
    range.s.r = Math.min(range.s.r, cellAddr.r)
    range.s.c = Math.min(range.s.c, cellAddr.c)
    range.e.r = Math.max(range.e.r, cellAddr.r)
    range.e.c = Math.max(range.e.c, cellAddr.c)
    results.push({ cell, value: w.value })
  }

  worksheet['!ref'] = XLSX.utils.encode_range(range)
  const outPath = outputPath ? path.resolve(outputPath) : resolved
  if (outputPath) {
    const dir = path.dirname(outPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  }
  XLSX.writeFile(workbook, outPath)
  return { ok: true, filePath: outPath, sheetName: targetSheetName, written: results }
}

export function readPreview({ filePath, sheetName, maxRows = 10, maxCols = 10 }) {
  const resolved = path.resolve(filePath)
  if (!fs.existsSync(resolved)) {
    throw new Error(`Excel file not found at ${resolved}`)
  }
  const workbook = XLSX.readFile(resolved)
  const targetSheetName = sheetName || workbook.SheetNames[0]
  const worksheet = workbook.Sheets[targetSheetName]
  if (!worksheet) {
    throw new Error(`Sheet not found: ${targetSheetName}`)
  }
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')
  const rows = Math.min(maxRows, (range.e.r - range.s.r + 1) || maxRows)
  const cols = Math.min(maxCols, (range.e.c - range.s.c + 1) || maxCols)
  const matrix = []
  for (let r = 0; r < rows; r++) {
    const row = []
    for (let c = 0; c < cols; c++) {
      const addr = XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c })
      const cell = worksheet[addr]
      row.push(cell ? cell.v : '')
    }
    matrix.push(row)
  }
  return { ok: true, sheetName: targetSheetName, rows: matrix.length, cols: cols, data: matrix }
}

export function listSheets({ filePath }) {
  const resolved = path.resolve(filePath)
  if (!fs.existsSync(resolved)) throw new Error(`Excel file not found at ${resolved}`)
  const workbook = XLSX.readFile(resolved)
  return { ok: true, sheets: workbook.SheetNames }
}

function resolveNameToRef(workbook, name) {
  const names = workbook?.Workbook?.Names || []
  const found = names.find((n) => (n?.Name || '').toLowerCase() === String(name).toLowerCase())
  return found?.Ref
}

export function listNamedRanges({ filePath }) {
  const resolved = path.resolve(filePath)
  if (!fs.existsSync(resolved)) throw new Error(`Excel file not found at ${resolved}`)
  const workbook = XLSX.readFile(resolved)
  const names = workbook?.Workbook?.Names || []
  const items = names.map(n => ({ name: n?.Name || '', ref: n?.Ref || '' })).filter(n => n.name)
  return { ok: true, names: items }
}

export function writeNamed({ filePath, name, value, outputPath }) {
  const resolved = path.resolve(filePath)
  if (!fs.existsSync(resolved)) throw new Error(`Excel file not found at ${resolved}`)
  const workbook = XLSX.readFile(resolved)
  let ref = resolveNameToRef(workbook, name)
  if (!ref) {
    // si name est déjà une ref A1 simple (ex: Sheet1!B3 ou B3)
    ref = String(name)
  }
  // ref peut être 'Sheet1'!$B$3 ou $B$3
  const m = /^(?:'?(.*?)'?!)?\$?([A-Za-z]+)\$?(\d+)$/.exec(ref)
  if (!m) throw new Error(`Invalid named ref or cell: ${ref}`)
  const sheet = m[1] || workbook.SheetNames[0]
  const cell = `${m[2]}${m[3]}`
  const ws = workbook.Sheets[sheet]
  if (!ws) throw new Error(`Sheet not found: ${sheet}`)
  const cellAddr = XLSX.utils.decode_cell(cell)
  const isNumeric = typeof value === 'number' || (typeof value === 'string' && value.trim() !== '' && !isNaN(Number(value)))
  ws[cell] = isNumeric ? { t: 'n', v: Number(value) } : { t: 's', v: String(value) }

  // Expand sheet range to include the written cell
  const currentRef = ws['!ref'] || cell
  const range = XLSX.utils.decode_range(currentRef)
  range.s.r = Math.min(range.s.r, cellAddr.r)
  range.s.c = Math.min(range.s.c, cellAddr.c)
  range.e.r = Math.max(range.e.r, cellAddr.r)
  range.e.c = Math.max(range.e.c, cellAddr.c)
  ws['!ref'] = XLSX.utils.encode_range(range)
  const outPath = outputPath ? path.resolve(outputPath) : resolved
  if (outputPath) {
    const dir = path.dirname(outPath)
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true })
  }
  XLSX.writeFile(workbook, outPath)
  return { ok: true, filePath: outPath, sheetName: sheet, cell, name }
} 