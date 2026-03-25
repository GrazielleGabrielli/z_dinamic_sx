import type { TPdfPageFormat } from '../config/types';

const PORTRAIT_MM: Record<TPdfPageFormat, readonly [number, number]> = {
  A0: [841, 1189],
  A1: [594, 841],
  A2: [420, 594],
  A3: [297, 420],
  A4: [210, 297],
  A5: [148, 210],
  A6: [105, 148],
  B4: [250, 353],
  B5: [176, 250],
  Letter: [215.9, 279.4],
  Legal: [215.9, 355.6],
  Tabloid: [279.4, 431.8],
  CreditCard: [53.98, 85.73],
};

export const VALID_PDF_PAGE_FORMATS: readonly TPdfPageFormat[] = [
  'A0',
  'A1',
  'A2',
  'A3',
  'A4',
  'A5',
  'A6',
  'B4',
  'B5',
  'Letter',
  'Legal',
  'Tabloid',
  'CreditCard',
];

export const PDF_PAGE_FORMAT_DROPDOWN_OPTIONS: { key: TPdfPageFormat; text: string }[] = [
  { key: 'A4', text: 'A4 (210 × 297 mm)' },
  { key: 'A3', text: 'A3 (297 × 420 mm)' },
  { key: 'A5', text: 'A5 (148 × 210 mm)' },
  { key: 'A6', text: 'A6 (105 × 148 mm)' },
  { key: 'A2', text: 'A2 (420 × 594 mm)' },
  { key: 'A1', text: 'A1 (594 × 841 mm)' },
  { key: 'A0', text: 'A0 (841 × 1189 mm)' },
  { key: 'B4', text: 'B4 (250 × 353 mm)' },
  { key: 'B5', text: 'B5 (176 × 250 mm)' },
  { key: 'Letter', text: 'Letter (216 × 279 mm)' },
  { key: 'Legal', text: 'Legal (216 × 356 mm)' },
  { key: 'Tabloid', text: 'Tabloid / Ledger (279 × 432 mm)' },
  { key: 'CreditCard', text: 'Cartão (ISO ID-1)' },
];

export function isValidPdfPageFormat(v: unknown): v is TPdfPageFormat {
  return typeof v === 'string' && (VALID_PDF_PAGE_FORMATS as readonly string[]).indexOf(v) !== -1;
}

export function normalizePdfPageFormat(value: unknown): TPdfPageFormat {
  if (isValidPdfPageFormat(value)) return value;
  return 'A4';
}

/** Chave do formato em `new jsPDF({ format })`. */
export function toJsPdfFormat(format: TPdfPageFormat): string {
  if (format === 'CreditCard') return 'credit-card';
  return format.toLowerCase();
}

export function getPdfPageSizeMm(
  format: TPdfPageFormat,
  orientation: 'portrait' | 'landscape'
): { widthMm: number; heightMm: number } {
  const pair = PORTRAIT_MM[format] ?? PORTRAIT_MM.A4;
  const pw = pair[0];
  const ph = pair[1];
  if (orientation === 'landscape') {
    return { widthMm: Math.max(pw, ph), heightMm: Math.min(pw, ph) };
  }
  return { widthMm: pw, heightMm: ph };
}
