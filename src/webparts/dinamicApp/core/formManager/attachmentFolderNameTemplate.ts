const INVALID_LITERAL_CHAR_CLASS = /[\\/:*?"<>|#%]/;

function sanitizeLiteralChunk(literal: string): string {
  return literal
    .replace(/[\\/:*?"<>|#%]/g, ' ')
    .split('')
    .filter((ch) => ch.charCodeAt(0) >= 32)
    .join('')
    .replace(/\s+/g, ' ');
}

function literalSegmentInvalidReason(literal: string): string | undefined {
  if (!literal.length) return undefined;
  if (INVALID_LITERAL_CHAR_CLASS.test(literal)) {
    return 'Não use estes carateres no texto fixo: \\ / : * ? " < > | # %';
  }
  for (let i = 0; i < literal.length; i++) {
    if (literal.charCodeAt(i) < 32) {
      return 'Carateres de controlo não são permitidos.';
    }
  }
  return undefined;
}

/** Valida apenas os troços literais; `{{ ... }}` fica de fora. */
export function folderTemplateLiteralInvalidReason(template: string): string | undefined {
  let i = 0;
  while (i < template.length) {
    const open = template.indexOf('{{', i);
    if (open === -1) {
      return literalSegmentInvalidReason(template.slice(i));
    }
    if (open > i) {
      const r = literalSegmentInvalidReason(template.slice(i, open));
      if (r) return r;
    }
    const close = template.indexOf('}}', open + 2);
    if (close === -1) {
      return undefined;
    }
    i = close + 2;
  }
  return undefined;
}

export function sanitizeFolderNameTemplatePreservingPlaceholders(template: string): string {
  let out = '';
  let i = 0;
  while (i < template.length) {
    const open = template.indexOf('{{', i);
    if (open === -1) {
      out += sanitizeLiteralChunk(template.slice(i));
      break;
    }
    if (open > i) {
      out += sanitizeLiteralChunk(template.slice(i, open));
    }
    const close = template.indexOf('}}', open + 2);
    if (close === -1) {
      out += template.slice(open);
      break;
    }
    out += template.slice(open, close + 2);
    i = close + 2;
  }
  return out.replace(/\s+/g, ' ').trim();
}
