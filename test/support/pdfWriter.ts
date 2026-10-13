/** Writes a PDF from numbered object bodies (object i + 1 = objects[i]; object 1 must be the catalog). */
export function writePdf(objects: string[]): Buffer {
  let out = '%PDF-1.4\n%âãÏÓ\n';
  const offsets: number[] = [];
  objects.forEach((body, i) => {
    offsets.push(Buffer.byteLength(out, 'latin1'));
    out += `${i + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = Buffer.byteLength(out, 'latin1');
  out += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  out += offsets.map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`).join('');
  out += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return Buffer.from(out, 'latin1');
}

function pdfString(text: string): string {
  return `(${text.replace(/[\\()]/g, (character) => `\\${character}`)})`;
}

/** US Letter pages, each showing its lines of text in Helvetica. */
export function textPdf(pages: string[][]): Buffer {
  const pageIds = pages.map((_, i) => 4 + i * 2);
  const objects = ['<< /Type /Catalog /Pages 2 0 R >>', `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pages.length} >>`, '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'];
  for (const [i, lines] of pages.entries()) {
    const text = lines.map((line, k) => `BT /F1 18 Tf 72 ${720 - k * 28} Td ${pdfString(line)} Tj ET`).join('\n');
    const contentsId = 4 + i * 2 + 1;
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> >> /Contents ${contentsId} 0 R >>`);
    objects.push(`<< /Length ${Buffer.byteLength(text, 'latin1')} >>\nstream\n${text}\nendstream`);
  }
  return writePdf(objects);
}
