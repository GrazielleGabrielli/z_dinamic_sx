/** Ícone Fluent por nome (extensão); alinhado ao FormAttachmentUploader. */
export function attachmentFileKindIconName(fileName: string): string {
  const n = fileName.toLowerCase();
  if (/\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(n)) return 'FileImage';
  if (n.endsWith('.pdf')) return 'PDF';
  if (n.endsWith('.doc') || n.endsWith('.docx')) return 'WordDocument';
  if ((/\.xlsx?$/.test(n) || /\.xls$/.test(n)) && !n.endsWith('.csv')) return 'ExcelDocument';
  if (/\.pptx?$/.test(n)) return 'PowerPointDocument';
  return 'Page';
}
