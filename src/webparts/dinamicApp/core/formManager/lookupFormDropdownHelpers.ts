import type { IFieldMetadata } from '../../../../services';
import type { IFormFieldConfig } from '../config/types/formManager';

function extractLookupId(v: unknown): number | undefined {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (typeof id === 'number') return id;
  }
  return undefined;
}

export function buildLookupODataFilter(
  childField: string,
  operator: string,
  parentValue: unknown,
  parentMeta: IFieldMetadata | undefined,
  childFieldMeta: IFieldMetadata | undefined
): string | undefined {
  const isLookupParent = parentMeta &&
    (parentMeta.MappedType === 'lookup' || parentMeta.MappedType === 'lookupmulti' ||
     parentMeta.MappedType === 'user' || parentMeta.MappedType === 'usermulti');
  const isLookupChild = childFieldMeta &&
    (childFieldMeta.MappedType === 'lookup' || childFieldMeta.MappedType === 'lookupmulti' ||
     childFieldMeta.MappedType === 'user' || childFieldMeta.MappedType === 'usermulti');
  const childKey = isLookupChild ? `${childField}Id` : childField;

  if (isLookupParent) {
    const id = extractLookupId(parentValue);
    if (id === undefined) return undefined;
    if (operator === 'eq') return `${childKey} eq ${id}`;
    if (operator === 'ne') return `${childKey} ne ${id}`;
    if (operator === 'lt') return `${childKey} lt ${id}`;
    if (operator === 'le') return `${childKey} le ${id}`;
    if (operator === 'gt') return `${childKey} gt ${id}`;
    if (operator === 'ge') return `${childKey} ge ${id}`;
    return undefined;
  }
  if (typeof parentValue === 'number' && isFinite(parentValue)) {
    if (operator === 'eq') return `${childKey} eq ${parentValue}`;
    if (operator === 'ne') return `${childKey} ne ${parentValue}`;
    if (operator === 'lt') return `${childKey} lt ${parentValue}`;
    if (operator === 'le') return `${childKey} le ${parentValue}`;
    if (operator === 'gt') return `${childKey} gt ${parentValue}`;
    if (operator === 'ge') return `${childKey} ge ${parentValue}`;
    return undefined;
  }
  if (typeof parentValue === 'string' && parentValue.trim()) {
    const esc = parentValue.replace(/'/g, "''");
    if (operator === 'eq') return `${childKey} eq '${esc}'`;
    if (operator === 'ne') return `${childKey} ne '${esc}'`;
    if (operator === 'lt') return `${childKey} lt '${esc}'`;
    if (operator === 'le') return `${childKey} le '${esc}'`;
    if (operator === 'gt') return `${childKey} gt '${esc}'`;
    if (operator === 'ge') return `${childKey} ge '${esc}'`;
    if (operator === 'contains') return `substringof('${esc}', ${childKey})`;
    if (operator === 'startsWith') return `startswith(${childKey}, '${esc}')`;
  }
  return undefined;
}

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

function extractSingleValue(item: unknown, subProp?: string): string {
  if (item === null || item === undefined) return '';
  if (typeof item === 'string' || typeof item === 'number') return String(item);
  if (typeof item === 'boolean') return item ? 'Sim' : 'Não';
  if (typeof item === 'object') {
    const o = item as Record<string, unknown>;
    if (subProp && o[subProp] !== undefined && o[subProp] !== null) return String(o[subProp]);
    if (typeof o.Title === 'string' || typeof o.Title === 'number') return String(o.Title ?? '');
    if (o.Title !== undefined && o.Title !== null) return String(o.Title);
    if ('LookupValue' in o && typeof o.LookupValue === 'string') return o.LookupValue;
    if ('Label' in o && typeof o.Label === 'string') return o.Label;
    if ('EMail' in o && typeof o.EMail === 'string') return o.EMail;
  }
  return String(item ?? '');
}

export function lookupRowToOptionText(
  row: Record<string, unknown>,
  labelInternal: string,
  labelMeta: IFieldMetadata | undefined,
  subProp?: string
): string {
  const v = row[labelInternal];
  if (v === null || v === undefined) {
    const id = row.Id;
    return id !== undefined && id !== null ? `#${String(id)}` : '';
  }
  if (Array.isArray(v)) {
    return v
      .map((item) => extractSingleValue(item, subProp))
      .filter(Boolean)
      .join('; ');
  }
  if (typeof v === 'string' || typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'Sim' : 'Não';
  if (typeof v === 'object') {
    return extractSingleValue(v, subProp);
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
