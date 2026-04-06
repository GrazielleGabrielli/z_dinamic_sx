import type { TFormManagerFormMode } from '../config/types/formManager';

function formatValueForUrl(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v !== null && 'Title' in (v as object)) {
    return String((v as Record<string, unknown>).Title ?? '');
  }
  return String(v);
}

function formModeToSharePointFormParam(mode: TFormManagerFormMode): string {
  if (mode === 'view') return 'Display';
  if (mode === 'edit') return 'Edit';
  return 'New';
}

/**
 * Substitui {{Campo}}, {{FormID}}, {{Form}} no URL. Valores são codificados para uso em query/caminho.
 */
export function interpolateFormButtonRedirectUrl(
  template: string,
  values: Record<string, unknown>,
  opts: { itemId?: number; formMode: TFormManagerFormMode }
): string {
  let out = template;
  const id = opts.itemId;
  out = out.replace(/\{\{\s*FormID\s*\}\}/gi, id !== undefined && id !== null ? encodeURIComponent(String(id)) : '');
  out = out.replace(/\{\{\s*Form\s*\}\}/gi, encodeURIComponent(formModeToSharePointFormParam(opts.formMode)));
  out = out.replace(/\{\{\s*([^}]+?)\s*\}\}/g, (_m, rawName: string) => {
    const name = String(rawName).trim();
    if (name === 'FormID' || name === 'Form') return '';
    const v = values[name];
    return encodeURIComponent(formatValueForUrl(v));
  });
  return out;
}
