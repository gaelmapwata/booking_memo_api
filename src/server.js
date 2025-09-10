import express from 'express'
import cors from 'cors'
import { writeCell } from './excel.js'
import { replacePlaceholders } from './word.js'

const app = express()
app.use(cors())
app.use(express.json())

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

const PORT = process.env.PORT || 4000
app.listen(PORT, () => {
  console.log(`API running on http://localhost:${PORT}`)
}) 