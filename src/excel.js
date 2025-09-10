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