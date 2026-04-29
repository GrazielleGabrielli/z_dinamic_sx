import type { IFieldMetadata } from '../../../../services';
import type { IFormFieldConfig } from '../config/types/formManager';

/** Campo na lista ligada para o texto das opções (SharePoint LookupField ou Title). */
export function resolveLookupFormLabelInternalName(
  meta: IFieldMetadata,
  fc: Pick<IFormFieldConfig, 'lookupOptionLabelField'>
): string {
  const t = fc.lookupOptionLabelField?.trim();
  if (t) return t;
  const d = meta.LookupField?.trim();
  return d || 'Title';
}

/** Id + etiqueta + extras (ordenado, sem repetir). */
export function buildLookupDropdownSelectRaw(
  meta: IFieldMetadata,
  fc: Pick<IFormFieldConfig, 'lookupOptionLabelField' | 'lookupOptionExtraSelectFields'>
): string[] {
  const label = resolveLookupFormLabelInternalName(meta, fc);
  const extras = fc.lookupOptionExtraSelectFields ?? [];
  const set = new Set<string>(['Id', label]);
  for (let i = 0; i < extras.length; i++) {
    const x = extras[i]?.trim();
    if (!x || x === 'Id') continue;
    set.add(x);
  }
  return Array.from(set);
}

export function lookupRowToOptionText(
  row: Record<string, unknown>,
  labelInternal: string,
  labelMeta: IFieldMetadata | undefined
): string {
  const v = row[labelInternal];
  if (v === null || v === undefined) {
    const id = row.Id;
    return id !== undefined && id !== null ? `#${String(id)}` : '';
  }
  if (typeof v === 'string' || typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'Sim' : 'Não';
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>;
    if (typeof o.Title === 'string' || typeof o.Title === 'number') return String(o.Title ?? '');
    if (o.Title !== undefined && o.Title !== null) return String(o.Title);
    if ('Label' in o && typeof o.Label === 'string') return o.Label;
    if ('EMail' in o && typeof o.EMail === 'string') return o.EMail;
  }
  if (labelMeta?.MappedType === 'datetime' || labelMeta?.TypeAsString === 'DateTime') {
    try {
      return String(v);
    } catch {
      return `#${String(row.Id ?? '')}`;
    }
  }
  return `#${String(row.Id ?? '')}`;
}
