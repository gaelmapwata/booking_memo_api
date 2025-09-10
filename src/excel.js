import fs from 'fs'
import path from 'path'
import XLSX from 'xlsx'

export function writeCell({ filePath, sheetName, cell, value }) {
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
  worksheet[cell] = { t: 's', v: String(value) }
  if (!worksheet['!ref']) {
    worksheet['!ref'] = cell
  }
  XLSX.writeFile(workbook, resolved)
  return { ok: true, filePath: resolved, sheetName: targetSheetName, cell, value }
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

export function writeNamed({ filePath, name, value }) {
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
  ws[cell] = { t: 's', v: String(value) }
  if (!ws['!ref']) ws['!ref'] = cell
  XLSX.writeFile(workbook, resolved)
  return { ok: true, filePath: resolved, sheetName: sheet, cell, name }
} 