import type { IFieldMetadata } from '../../../../services';
import type {
  IFormFieldConfig,
  IFormManagerConfig,
  IFormStepConfig,
  TFormConditionNode,
} from '../config/types/formManager';
import { collectLookupSubfieldsFromUserVisibilityPaths } from './formButtonLookupUserVisibility';

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

export function hasConfiguredLookupFilter(lf: {
  childField?: string;
  filterOperator?: string;
  odataFilterTemplate?: string;
} | undefined): boolean {
  if (!lf) return false;
  if (lf.childField?.trim() && lf.filterOperator) return true;
  if ((lf.odataFilterTemplate ?? '').trim()) return true;
  return false;
}

export function isParentValueReadyForLookupFilter(
  parentValue: unknown,
  parentMeta: IFieldMetadata | undefined
): boolean {
  if (!parentMeta) {
    if (parentValue === null || parentValue === undefined) return false;
    if (typeof parentValue === 'string') return parentValue.trim().length > 0;
    if (typeof parentValue === 'number') return isFinite(parentValue);
    if (typeof parentValue === 'boolean') return true;
    if (typeof parentValue === 'object' && parentValue !== null && 'Id' in parentValue) {
      return extractLookupId(parentValue) !== undefined;
    }
    if (Array.isArray(parentValue)) return parentValue.length > 0;
    return false;
  }
  const mt = parentMeta.MappedType;
  if (mt === 'lookup' || mt === 'lookupmulti' || mt === 'user' || mt === 'usermulti') {
    return extractLookupId(parentValue) !== undefined;
  }
  if (mt === 'number' || mt === 'currency') {
    return typeof parentValue === 'number' && isFinite(parentValue);
  }
  if (mt === 'datetime') {
    return parentValue !== null && parentValue !== undefined && String(parentValue).trim() !== '';
  }
  if (mt === 'boolean') {
    return parentValue !== null && parentValue !== undefined;
  }
  if (mt === 'multichoice') {
    if (Array.isArray(parentValue)) return parentValue.length > 0;
    if (typeof parentValue === 'string') return parentValue.trim().length > 0;
    return false;
  }
  if (typeof parentValue === 'string') return parentValue.trim().length > 0;
  if (typeof parentValue === 'number') return isFinite(parentValue);
  return parentValue !== null && parentValue !== undefined;
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

export type TLookupReloadFilterRow = {
  parentField: string;
  childField?: string;
  filterOperator?: string;
  odataFilterTemplate?: string;
};

export function buildLookupReloadRowSignatures(params: {
  fieldConfigs: readonly IFormFieldConfig[];
  metaByName: ReadonlyMap<string, IFieldMetadata>;
  lookupFilters: Readonly<Record<string, TLookupReloadFilterRow | undefined>>;
  values: Readonly<Record<string, unknown>>;
  collectInjectFields: (lookupInternalName: string) => readonly string[];
}): { joinedKey: string; rowKeyByField: Record<string, string> } {
  const { fieldConfigs, metaByName, lookupFilters, values, collectInjectFields } = params;
  const rowKeyByField: Record<string, string> = {};
  const parts: string[] = [];
  for (let i = 0; i < fieldConfigs.length; i++) {
    const fc = fieldConfigs[i];
    const fn = fc.internalName;
    const m = metaByName.get(fn);
    if (!m || (m.MappedType !== 'lookup' && m.MappedType !== 'lookupmulti')) continue;
    const listId = String(m.LookupList ?? '');
    const labelDisp = resolveLookupFormLabelInternalName(m, fc);
    const extrasSig = JSON.stringify(fc.lookupOptionExtraSelectFields ?? []);
    const subPropSig = fc.lookupOptionLabelSubProp ?? '';
    const detailSig = JSON.stringify(fc.lookupOptionDetailBelowFields ?? []);
    const lf = lookupFilters[fn];
    const injectSig = collectInjectFields(fn).slice().sort().join('|');
    let parentSig = '';
    if (lf) {
      const parentVal = values[lf.parentField];
      const parentId = extractLookupId(parentVal);
      parentSig =
        parentId !== undefined
          ? String(parentId)
          : typeof parentVal === 'string'
            ? parentVal
            : typeof parentVal === 'number'
              ? String(parentVal)
              : '';
    }
    const row =
      lf != null
        ? `${fn}\t${listId}\t${labelDisp}\t${extrasSig}\t${subPropSig}\t${detailSig}\t${injectSig}\t${lf.parentField}\t${lf.childField ?? ''}\t${lf.filterOperator ?? ''}\t${parentSig}`
        : `${fn}\t${listId}\t${labelDisp}\t${extrasSig}\t${subPropSig}\t${detailSig}\t${injectSig}\t`;
    rowKeyByField[fn] = row;
    parts.push(row);
  }
  parts.sort();
  return { joinedKey: parts.join('\n'), rowKeyByField };
}

/** Id + etiqueta + extras + detalhe abaixo (ordenado, sem repetir). */
function walkConditionNodes(node: TFormConditionNode | undefined, visit: (n: TFormConditionNode) => void): void {
  if (!node) return;
  if (node.kind === 'all' || node.kind === 'any') {
    for (let i = 0; i < node.children.length; i++) walkConditionNodes(node.children[i], visit);
    return;
  }
  visit(node);
}

function collectFirstSubfieldAfterRoot(
  node: TFormConditionNode | undefined,
  lookupRoot: string,
  into: Set<string>
): void {
  const root = lookupRoot.trim();
  if (!root) return;
  const takePath = (raw: string | undefined): void => {
    const path = (raw ?? '').trim();
    if (!path || path.indexOf('/') === -1) return;
    const parts = path.split('/').map((p) => p.trim()).filter(Boolean);
    if (parts.length < 2 || parts[0] !== root) return;
    const sub = parts[1];
    if (sub) into.add(sub);
  };
  walkConditionNodes(node, (n) => {
    if (n.kind !== 'leaf') return;
    takePath(n.field);
    if (n.compare?.kind === 'field') takePath(n.compare.value);
  });
}

/**
 * Campos na lista ligada do lookup a incluir no $select das opções (regras, passos, ajuda dinâmica).
 */
export function collectLookupSelectInjectFields(
  cfg: Partial<Pick<IFormManagerConfig, 'rules' | 'steps' | 'dynamicHelp' | 'customButtons'>>,
  lookupInternalName: string
): string[] {
  const subs = new Set<string>();
  const rules = cfg.rules ?? [];
  for (let i = 0; i < rules.length; i++) {
    collectFirstSubfieldAfterRoot(rules[i].when, lookupInternalName, subs);
  }
  const steps = cfg.steps ?? [];
  for (let i = 0; i < steps.length; i++) {
    collectFirstSubfieldAfterRoot((steps[i] as IFormStepConfig).showStepWhen, lookupInternalName, subs);
  }
  const dh = cfg.dynamicHelp ?? [];
  for (let i = 0; i < dh.length; i++) collectFirstSubfieldAfterRoot(dh[i].when, lookupInternalName, subs);
  const buttons = cfg.customButtons ?? [];
  for (let i = 0; i < buttons.length; i++) {
    const b = buttons[i];
    collectLookupSubfieldsFromUserVisibilityPaths(b.lookupUserFieldPaths, lookupInternalName, subs);
    collectLookupSubfieldsFromUserVisibilityPaths(b.excludeLookupUserFieldPaths, lookupInternalName, subs);
  }
  return Array.from(subs);
}

export function buildLookupDropdownSelectRaw(
  meta: IFieldMetadata,
  fc: Pick<
    IFormFieldConfig,
    'lookupOptionLabelField' | 'lookupOptionExtraSelectFields' | 'lookupOptionDetailBelowFields'
  >,
  injectSelectFields?: readonly string[]
): string[] {
  const label = resolveLookupFormLabelInternalName(meta, fc);
  const extras = fc.lookupOptionExtraSelectFields ?? [];
  const details = fc.lookupOptionDetailBelowFields ?? [];
  const set = new Set<string>(['Id', label]);
  for (let i = 0; i < extras.length; i++) {
    const x = extras[i]?.trim();
    if (!x || x === 'Id') continue;
    set.add(x);
  }
  for (let i = 0; i < details.length; i++) {
    const x = details[i]?.trim();
    if (!x || x === 'Id') continue;
    set.add(x);
  }
  if (injectSelectFields) {
    for (let i = 0; i < injectSelectFields.length; i++) {
      const x = injectSelectFields[i]?.trim();
      if (!x || x === 'Id') continue;
      set.add(x);
    }
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
