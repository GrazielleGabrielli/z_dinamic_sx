import { ItemsService } from '../../../../services';
import type { IFormManagerConfig } from '../config/types/formManager';

function isEmptyish(v: unknown): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === 'string' && v.trim() === '') return true;
  return false;
}

function escapeODataString(s: string): string {
  return s.replace(/'/g, "''");
}

export function replaceODataFilterTemplate(tpl: string, values: Record<string, unknown>): string {
  return tpl.replace(/\{\{([^}]+)\}\}/g, (_, key) => {
    const k = String(key).trim();
    const v = values[k];
    if (v === null || v === undefined) return "''";
    if (typeof v === 'number' && isFinite(v)) return String(v);
    if (typeof v === 'boolean') return v ? '1' : '0';
    return `'${escapeODataString(String(v))}'`;
  });
}

export async function runAsyncFormValidations(
  cfg: IFormManagerConfig,
  values: Record<string, unknown>,
  itemsService: ItemsService,
  listTitle: string,
  excludeItemId?: number,
  submitKind: 'draft' | 'submit' = 'submit'
): Promise<Record<string, string>> {
  const errors: Record<string, string> = {};
  if (submitKind === 'draft') return errors;
  const rules = cfg.rules ?? [];
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (rule.enabled === false) continue;
    if (rule.action === 'asyncUniqueness') {
      const field = rule.field;
      const v = values[field];
      if (isEmptyish(v)) continue;
      const targetList = (rule.listTitle && rule.listTitle.trim()) || listTitle;
      const lit = typeof v === 'number' && isFinite(v) ? String(v) : `'${escapeODataString(String(v))}'`;
      let filter = `${field} eq ${lit}`;
      if (excludeItemId !== undefined) filter = `(${filter}) and Id ne ${excludeItemId}`;
      try {
        const cnt = await itemsService.countItems(targetList, filter);
        if (cnt > 0) errors[field] = rule.message ?? 'Valor já existe.';
      } catch {
        errors[field] = rule.message ?? 'Não foi possível validar unicidade.';
      }
    }
    if (rule.action === 'asyncCountLimit') {
      const targetList = (rule.listTitle && rule.listTitle.trim()) || listTitle;
      const filter = replaceODataFilterTemplate(rule.filterTemplate, values);
      if (!filter.trim()) continue;
      try {
        const cnt = await itemsService.countItems(targetList, filter);
        if (cnt > rule.maxCount) errors._async = rule.message ?? 'Limite excedido.';
      } catch {
        errors._async = rule.message ?? 'Não foi possível validar limite.';
      }
    }
  }
  return errors;
}
