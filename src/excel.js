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