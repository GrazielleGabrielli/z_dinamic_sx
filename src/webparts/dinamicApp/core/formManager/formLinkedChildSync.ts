import type { IFieldMetadata } from '../../../../services';
import type { IFormLinkedChildFormConfig, IFormManagerConfig } from '../config/types/formManager';
import type { ItemsService } from '../../../../services';
import { formValuesToSharePointPayload } from './formSharePointValues';

export interface ILinkedChildRowState {
  localKey: string;
  sharePointId?: number;
  values: Record<string, unknown>;
}

export function linkedChildFormAsManagerConfig(c: IFormLinkedChildFormConfig): IFormManagerConfig {
  return {
    sections: c.sections,
    fields: c.fields,
    rules: c.rules,
    steps: c.steps,
    stepLayout: 'segmented',
  };
}

export function itemFieldNamesToValues(
  item: Record<string, unknown>,
  names: string[]
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  for (let i = 0; i < names.length; i++) {
    const n = names[i];
    out[n] = item[n];
  }
  return out;
}

export function getLinkedChildMainStepFieldNames(cfg: IFormLinkedChildFormConfig): string[] {
  const st = cfg.steps?.find((s) => s.id === 'main');
  return st?.fieldNames?.slice() ?? [];
}

export function payloadFieldNamesForLinkedChild(
  cfg: IFormLinkedChildFormConfig
): string[] {
  const lk = cfg.parentLookupFieldInternalName.trim();
  const base = normalizeParentLookupFieldInternalName(lk);
  return getLinkedChildMainStepFieldNames(cfg).filter(
    (n) => n !== lk && n !== base && n !== `${base}Id`
  );
}

function isEmptyPayloadValue(v: unknown): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === 'string' && v.trim() === '') return true;
  if (Array.isArray(v) && v.length === 0) return true;
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (id === null || id === undefined || id === '') return true;
  }
  if (typeof v === 'object' && v !== null && 'Url' in v) {
    return String((v as Record<string, unknown>).Url ?? '').trim() === '';
  }
  return false;
}

export function isLinkedChildRowSkippableForCreate(
  values: Record<string, unknown>,
  fieldMetadata: IFieldMetadata[],
  payloadNames: string[]
): boolean {
  const byName = new Map(fieldMetadata.map((f) => [f.InternalName, f]));
  let anyRequired = false;
  let anyFilled = false;
  for (let i = 0; i < payloadNames.length; i++) {
    const n = payloadNames[i];
    const m = byName.get(n);
    if (!m) continue;
    const v = values[n];
    if (m.Required) anyRequired = true;
    if (!isEmptyPayloadValue(v)) anyFilled = true;
  }
  if (anyRequired) return false;
  return !anyFilled;
}

/** Nome interno da coluna Lookup (sem sufixo Id). Evita «CampoId» → «CampoIdId» no OData/payload. */
export function normalizeParentLookupFieldInternalName(name: string): string {
  const t = name.trim();
  if (t.length > 2 && /Id$/i.test(t)) {
    return t.slice(0, -2);
  }
  return t;
}

export function buildODataFilterForParentLookup(
  parentLookupFieldInternalName: string,
  parentItemId: number
): string {
  const base = normalizeParentLookupFieldInternalName(parentLookupFieldInternalName);
  return `${base}Id eq ${parentItemId}`;
}

export async function loadLinkedChildRows(
  itemsService: ItemsService,
  cfg: IFormLinkedChildFormConfig,
  parentItemId: number,
  fieldMetadata: IFieldMetadata[]
): Promise<ILinkedChildRowState[]> {
  const list = cfg.listTitle.trim();
  const lk = cfg.parentLookupFieldInternalName.trim();
  if (!list || !lk) return [];
  const filter = buildODataFilterForParentLookup(lk, parentItemId);
  const rows = await itemsService.getItems<Record<string, unknown>>(list, {
    filter,
    fieldMetadata,
    orderBy: { field: 'Id', ascending: true },
    top: 200,
  });
  const mainNames = getLinkedChildMainStepFieldNames(cfg);
  const out: ILinkedChildRowState[] = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const id = typeof r.Id === 'number' ? r.Id : Number(r.Id);
    if (!isFinite(id)) continue;
    out.push({
      localKey: `sp_${id}`,
      sharePointId: id,
      values: itemFieldNamesToValues(r, mainNames),
    });
  }
  return out;
}

export async function syncLinkedChildList(
  itemsService: ItemsService,
  cfg: IFormLinkedChildFormConfig,
  parentItemId: number,
  rows: ILinkedChildRowState[],
  fieldMetadata: IFieldMetadata[],
  baselineSharePointIds: number[]
): Promise<ILinkedChildRowState[]> {
  const list = cfg.listTitle.trim();
  const lk = cfg.parentLookupFieldInternalName.trim();
  if (!list || !lk) return rows.map((r) => ({ ...r, values: { ...r.values } }));
  const idKey = `${normalizeParentLookupFieldInternalName(lk)}Id`;
  const payloadNames = payloadFieldNamesForLinkedChild(cfg);
  const rowsNext = rows.map((r) => ({ ...r, values: { ...r.values } }));
  const nextIds = new Set(
    rowsNext.map((r) => r.sharePointId).filter((x): x is number => typeof x === 'number' && isFinite(x))
  );
  for (let i = 0; i < baselineSharePointIds.length; i++) {
    const id = baselineSharePointIds[i];
    if (nextIds.has(id)) continue;
    await itemsService.deleteItem(list, id);
  }
  for (let i = 0; i < rowsNext.length; i++) {
    const row = rowsNext[i];
    const payloadBase = formValuesToSharePointPayload(fieldMetadata, row.values, payloadNames);
    const payload: Record<string, unknown> = { ...payloadBase, [idKey]: parentItemId };
    if (row.sharePointId !== undefined && isFinite(row.sharePointId)) {
      await itemsService.updateItem(list, row.sharePointId, payload);
    } else {
      if (isLinkedChildRowSkippableForCreate(row.values, fieldMetadata, payloadNames)) continue;
      const { id } = await itemsService.addItem(list, payload);
      rowsNext[i] = { ...row, sharePointId: id };
    }
  }
  return rowsNext;
}

export async function syncAllLinkedChildLists(
  itemsService: ItemsService,
  configs: IFormLinkedChildFormConfig[],
  parentItemId: number,
  rowsByConfigId: Record<string, ILinkedChildRowState[]>,
  metaByConfigId: Record<string, IFieldMetadata[]>,
  baselineIdsByConfigId: Record<string, number[]>
): Promise<Record<string, ILinkedChildRowState[]>> {
  const sorted = configs
    .filter((c) => c.listTitle.trim() && c.parentLookupFieldInternalName.trim())
    .slice()
    .sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  const out: Record<string, ILinkedChildRowState[]> = {};
  for (let i = 0; i < sorted.length; i++) {
    const cfg = sorted[i];
    const meta = metaByConfigId[cfg.id] ?? [];
    const rows = rowsByConfigId[cfg.id] ?? [];
    const baseline = baselineIdsByConfigId[cfg.id] ?? [];
    out[cfg.id] = await syncLinkedChildList(itemsService, cfg, parentItemId, rows, meta, baseline);
  }
  return out;
}
