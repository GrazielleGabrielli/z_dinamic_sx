import type { IFieldMetadata } from '../../../../services';

/** HTML típico de coluna Nota com formatação / rich text no SharePoint. */
export function isSharePointRichNoteHtml(value: string): boolean {
  const t = value.trim();
  if (t.length < 12 || !t.startsWith('<')) return false;
  return (
    /ExternalClass[0-9A-F]{32}/i.test(t) ||
    (/<\s*div[^>]*\sclass\s*=\s*["'][^"']*ExternalClass/i.test(t) &&
      /<\/\s*(div|p|span)\s*>/i.test(t))
  );
}

export function shouldRenderMultilineNoteAsHtml(meta: IFieldMetadata, raw: string): boolean {
  if (meta.MappedType !== 'multiline') return false;
  if (meta.RichText === true) return true;
  return isSharePointRichNoteHtml(raw);
}
