import express from 'express'
import cors from 'cors'
import fs from 'fs'
import path from 'path'
import multer from 'multer'
import { writeCell, readPreview, listSheets, writeNamed, listNamedRanges, writeCellsBulk } from './excel.js'
import { replacePlaceholders } from './word.js'

const app = express()
app.use(cors())
app.use(express.json())

const uploadDir = path.resolve('uploads')
if (!fs.existsSync(uploadDir)) fs.mkdirSync(uploadDir)
app.use('/uploads', express.static(uploadDir))

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => {
    const ts = Date.now()
    const safe = file.originalname.replace(/[^a-zA-Z0-9_.-]/g, '_')
    cb(null, `${ts}_${safe}`)
  }
})
const upload = multer({ storage })

app.get('/health', (req, res) => {
  res.json({ status: 'ok' })
})

app.post('/excel/write', (req, res) => {
  const { filePath, sheetName, cell, value } = req.body || {}
  if (!filePath || !cell) {
    return res.status(400).json({ ok: false, error: 'filePath and cell are required' })
  }
  try {
    const result = writeCell({ filePath, sheetName, cell, value })
    res.json(result)
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/write-bulk', (req, res) => {
  const { filePath, sheetName, writes, respectMerges } = req.body || {}
  if (!filePath || !Array.isArray(writes)) {
    return res.status(400).json({ ok: false, error: 'filePath and writes[] are required' })
  }
  try {
    const result = writeCellsBulk({ filePath, sheetName, writes, respectMerges: respectMerges !== false })
    res.json(result)
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/write-bulk-upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, error: 'file is required' })
  let writes
  try {
    writes = req.body?.writes ? JSON.parse(req.body.writes) : []
  } catch (e) {
    return res.status(400).json({ ok: false, error: 'Invalid JSON in writes' })
  }
  const sheetName = req.body?.sheetName || undefined
  const respectMerges = req.body?.respectMerges !== 'false'
  try {
    const result = writeCellsBulk({ filePath: req.file.path, sheetName, writes, respectMerges })
    res.json({ ...result, uploaded: true, url: `/uploads/${path.basename(req.file.path)}` })
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/write-upload', upload.single('file'), (req, res) => {
  const { sheetName, cell, value } = req.body || {}
  if (!req.file || !cell) {
    return res.status(400).json({ ok: false, error: 'file and cell are required' })
  }
  const filePath = req.file.path
  try {
    const result = writeCell({ filePath, sheetName, cell, value })
    res.json({ ...result, uploaded: true, url: `/uploads/${path.basename(filePath)}` })
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/preview', (req, res) => {
  const { filePath, sheetName, maxRows, maxCols } = req.body || {}
  if (!filePath) return res.status(400).json({ ok: false, error: 'filePath is required' })
  try {
    const result = readPreview({ filePath, sheetName, maxRows, maxCols })
    res.json(result)
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/preview-upload', upload.single('file'), (req, res) => {
  const { sheetName, maxRows, maxCols } = req.body || {}
  if (!req.file) return res.status(400).json({ ok: false, error: 'file is required' })
  try {
    const result = readPreview({ filePath: req.file.path, sheetName, maxRows, maxCols })
    res.json({ ...result, uploaded: true })
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/sheets', (req, res) => {
  const { filePath } = req.body || {}
  if (!filePath) return res.status(400).json({ ok: false, error: 'filePath is required' })
  try {
    const result = listSheets({ filePath })
    res.json(result)
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/sheets-upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, error: 'file is required' })
  try {
    const result = listSheets({ filePath: req.file.path })
    res.json({ ...result, uploaded: true })
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/write-named', (req, res) => {
  const { filePath, name, value } = req.body || {}
  if (!filePath || !name) return res.status(400).json({ ok: false, error: 'filePath and name are required' })
  try {
    const result = writeNamed({ filePath, name, value })
    res.json(result)
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/write-named-upload', upload.single('file'), (req, res) => {
  const { name, value } = req.body || {}
  if (!req.file || !name) return res.status(400).json({ ok: false, error: 'file and name are required' })
  try {
    const result = writeNamed({ filePath: req.file.path, name, value })
    res.json({ ...result, uploaded: true })
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/names', (req, res) => {
  const { filePath } = req.body || {}
  if (!filePath) return res.status(400).json({ ok: false, error: 'filePath is required' })
  try {
    const result = listNamedRanges({ filePath })
    res.json(result)
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/excel/names-upload', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ ok: false, error: 'file is required' })
  try {
    const result = listNamedRanges({ filePath: req.file.path })
    res.json({ ...result, uploaded: true })
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/word/replace', (req, res) => {
  const { templatePath, outputPath, replacements } = req.body || {}
  if (!templatePath) {
    return res.status(400).json({ ok: false, error: 'templatePath is required' })
  }
  try {
    const result = replacePlaceholders({ templatePath, outputPath, replacements })
    res.json(result)
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

app.post('/word/replace-upload', upload.single('file'), (req, res) => {
  const { outputPath, replacements } = req.body || {}
  if (!req.file) {
    return res.status(400).json({ ok: false, error: 'file is required' })
  }
  const templatePath = req.file.path
  let parsed
  try {
    parsed = replacements ? JSON.parse(replacements) : {}
  } catch (e) {
    return res.status(400).json({ ok: false, error: 'Invalid JSON in replacements' })
  }
  try {
    const result = replacePlaceholders({ templatePath, outputPath, replacements: parsed })
    res.json({ ...result, uploaded: true, templateUrl: `/uploads/${path.basename(templatePath)}` })
  } catch (e) {
    res.status(400).json({ ok: false, error: e.message })
  }
})

const PORT = process.env.PORT || 4000
app.listen(PORT, () => {
  console.log(`API running on http://localhost:${PORT}`)
}) 