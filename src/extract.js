// Turn a document into text Claude can read.
import path from 'node:path';

const TEXT_EXT = new Set([
  '.txt', '.md', '.csv', '.tsv', '.json', '.xml', '.html', '.htm', '.rtf', '.log', '.yaml', '.yml', '.eml',
]);

export async function extractText(name, buf) {
  const ext = path.extname(name).toLowerCase();
  if (ext === '.docx') {
    const mammoth = (await import('mammoth')).default;
    const { value } = await mammoth.extractRawText({ buffer: buf });
    return value;
  }
  if (ext === '.pdf') {
    const { extractText: pdfText, getDocumentProxy } = await import('unpdf');
    const pdf = await getDocumentProxy(new Uint8Array(buf));
    const { text } = await pdfText(pdf, { mergePages: true });
    return text;
  }
  if (TEXT_EXT.has(ext)) return buf.toString('utf8');
  // Unknown type: return text only if it decodes cleanly.
  const sample = buf.subarray(0, 4096).toString('utf8');
  if (!sample.includes('�') && !/[\x00-\x08\x0E-\x1F]/.test(sample)) return buf.toString('utf8');
  return null;
}
