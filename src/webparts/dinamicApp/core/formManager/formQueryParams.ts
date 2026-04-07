import type { TFormManagerFormMode } from '../config/types/formManager';

function buildQueryKeyMap(q: Record<string, string>): Map<string, string> {
  const m = new Map<string, string>();
  for (const [k, v] of Object.entries(q)) {
    m.set(k.toLowerCase(), v);
  }
  return m;
}

export function getQueryParamCI(q: Record<string, string>, ...candidates: string[]): string | undefined {
  const m = buildQueryKeyMap(q);
  for (let i = 0; i < candidates.length; i++) {
    const v = m.get(candidates[i].toLowerCase());
    if (v !== undefined && String(v).trim() !== '') return v;
  }
  return undefined;
}

function normFormToken(s: string): string {
  return s.trim().toLowerCase();
}

/**
 * Id do item na lista (SharePoint: FormID; também id, ItemID, etc.).
 */
export function parseFormItemIdFromQuery(q: Record<string, string>): number | undefined {
  const raw = getQueryParamCI(q, 'FormID', 'ItemID', 'itemId', 'ID', 'id');
  if (!raw) return undefined;
  const n = parseInt(String(raw).trim(), 10);
  if (isNaN(n) || n < 1) return undefined;
  return n;
}

/**
 * FORM=Disp|Display → visualização (campos só leitura); Edit; New|Create.
 * Sem FORM mas com item carregado → edit (compatibilidade com URLs antigas só com id).
 */
export function resolveFormModeFromQuery(
  q: Record<string, string>,
  opts: { itemLoaded: boolean }
): TFormManagerFormMode {
  const formRaw = getQueryParamCI(q, 'FORM', 'Form');
  if (formRaw) {
    const f = normFormToken(formRaw);
    if (f === 'disp' || f === 'display') return 'view';
    if (f === 'edit') return 'edit';
    if (f === 'new' || f === 'create') return 'create';
  }
  if (opts.itemLoaded) return 'edit';
  return 'create';
}

export function isFormNewModeQuery(q: Record<string, string>): boolean {
  const formRaw = getQueryParamCI(q, 'FORM', 'Form');
  if (!formRaw) return false;
  const f = normFormToken(formRaw);
  return f === 'new' || f === 'create';
}
