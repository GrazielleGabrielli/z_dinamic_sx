import type { IFieldMetadata } from '../../../../services';

function lookupId(v: unknown): number | undefined {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (typeof id === 'number') return id;
  }
  return undefined;
}

function isEmptyForSpPayload(v: unknown): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === 'string' && v.trim() === '') return true;
  if (Array.isArray(v) && v.length === 0) return true;
  if (typeof v === 'object' && v !== null && 'Url' in v) {
    return String((v as Record<string, unknown>).Url ?? '').trim() === '';
  }
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (id === null || id === undefined || id === '') return true;
  }
  return false;
}

function writeNullForMappedType(
  out: Record<string, unknown>,
  name: string,
  mappedType: IFieldMetadata['MappedType']
): void {
  switch (mappedType) {
    case 'lookup':
    case 'user':
      out[`${name}Id`] = null;
      break;
    case 'lookupmulti':
    case 'usermulti':
      out[`${name}Id`] = { results: [] };
      break;
    case 'multichoice':
      out[name] = { results: [] };
      break;
    default:
      out[name] = null;
  }
}

export interface IFormValuesToSharePointPayloadOptions {
  nullWhenEmptyFieldNames?: string[];
}

export function formValuesToSharePointPayload(
  metadata: IFieldMetadata[],
  values: Record<string, unknown>,
  includeFields: string[],
  opts?: IFormValuesToSharePointPayloadOptions
): Record<string, unknown> {
  const byName = new Map(metadata.map((f) => [f.InternalName, f]));
  const out: Record<string, unknown> = {};
  const nullSet =
    opts?.nullWhenEmptyFieldNames && opts.nullWhenEmptyFieldNames.length > 0
      ? new Set(opts.nullWhenEmptyFieldNames)
      : null;
  for (let i = 0; i < includeFields.length; i++) {
    const name = includeFields[i];
    const m = byName.get(name);
    if (!m || m.ReadOnlyField) continue;
    if (m.Hidden) continue;
    const v = values[name];
    if (nullSet?.has(name)) {
      if (isEmptyForSpPayload(v)) {
        writeNullForMappedType(out, name, m.MappedType);
        continue;
      }
    }
    if (v === undefined) continue;
    switch (m.MappedType) {
      case 'text':
      case 'multiline':
        out[name] = v === null ? null : String(v);
        break;
      case 'url': {
        if (v === null || v === undefined) {
          out[name] = null;
          break;
        }
        if (typeof v === 'object' && v !== null && 'Url' in v) {
          const o = v as Record<string, unknown>;
          const url = String(o.Url ?? '').trim();
          if (url === '') {
            out[name] = null;
          } else {
            out[name] = {
              Url: url,
              Description: String(o.Description ?? ''),
            };
          }
          break;
        }
        out[name] = String(v);
        break;
      }
      case 'number':
      case 'currency': {
        const n = typeof v === 'number' ? v : Number(String(v).replace(',', '.'));
        out[name] = isNaN(n) ? null : n;
        break;
      }
      case 'boolean':
        out[name] = v === true || v === 1 || v === '1' || v === 'true';
        break;
      case 'datetime':
        out[name] = v instanceof Date ? v.toISOString() : v === null ? null : String(v);
        break;
      case 'choice':
        out[name] = v === null ? null : String(v);
        break;
      case 'multichoice': {
        const arr = Array.isArray(v)
          ? (v as unknown[]).map((x) => String(x))
          : String(v)
              .split(';')
              .map((s) => s.trim())
              .filter(Boolean);
        out[name] = { results: arr };
        break;
      }
      case 'lookup': {
        const id = lookupId(v);
        if (id !== undefined) out[`${name}Id`] = id;
        break;
      }
      case 'lookupmulti': {
        const ids = Array.isArray(v)
          ? (v as unknown[]).map(lookupId).filter((x): x is number => x !== undefined)
          : lookupId(v) !== undefined
            ? [lookupId(v) as number]
            : [];
        out[`${name}Id`] = { results: ids };
        break;
      }
      case 'user': {
        const id = lookupId(v);
        if (id !== undefined) out[`${name}Id`] = id;
        break;
      }
      case 'usermulti': {
        const ids = Array.isArray(v)
          ? (v as unknown[]).map(lookupId).filter((x): x is number => x !== undefined)
          : [];
        out[`${name}Id`] = { results: ids };
        break;
      }
      default:
        break;
    }
  }
  return out;
}
