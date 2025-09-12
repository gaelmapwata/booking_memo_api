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
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    // Avoid printing "undefined" when a key is missing
    nullGetter: () => ''
  })
  doc.setData(replacements || {})
  try {
    doc.render()
  } catch (error) {
    // Improve Docxtemplater error visibility (e.g., "Multi error")
    try {
      const details = Array.isArray(error?.errors)
        ? error.errors.map((e) => (
            [
              e?.properties?.explanation,
              e?.properties?.id,
              e?.message
            ].filter(Boolean).join(' | ')
          ))
        : []
      const msg = details.length
        ? `Docx render error: ${error.message} :: ${details.join(' || ')}`
        : `Docx render error: ${error.message}`
      throw new Error(msg)
    } catch (_) {
      throw new Error(`Docx render error: ${error.message}`)
    }
  }
  const buf = doc.getZip().generate({ type: 'nodebuffer' })
  const out = path.resolve(outputPath || input.replace(/\.docx$/i, '.out.docx'))
  const outDir = path.dirname(out)
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir, { recursive: true })
  fs.writeFileSync(out, buf)
  return { ok: true, outputPath: out }
} 