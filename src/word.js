import fs from 'fs'
import path from 'path'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'

export function replacePlaceholders({ templatePath, outputPath, replacements }) {
  const input = path.resolve(templatePath)
  if (!fs.existsSync(input)) {
    throw new Error(`Word template not found at ${input}`)
  }
  const content = fs.readFileSync(input, 'binary')
  const zip = new PizZip(content)
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true })
  doc.setData(replacements || {})
  try {
    doc.render()
  } catch (error) {
    throw new Error(`Docx render error: ${error.message}`)
  }
  const buf = doc.getZip().generate({ type: 'nodebuffer' })
  const out = path.resolve(outputPath || input.replace(/\.docx$/i, '.out.docx'))
  fs.writeFileSync(out, buf)
  return { ok: true, outputPath: out }
} 