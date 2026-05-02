import type {
  IFormCustomButtonConfig,
  IFormManagerActionLogConfig,
  TFormManagerFormMode,
} from '../config/types/formManager';
import type { IFieldMetadata, ItemsService } from '../../../../services';

export const FORM_ACTION_LOG_BTN_DATA_ATTR = 'data-sx-log-btn';

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export function parseActionLogButtonIdFromStoredHtml(html: string): string | undefined {
  if (!html || typeof html !== 'string') return undefined;
  const m = html.match(/\bdata-sx-log-btn\s*=\s*"([^"]+)"/i);
  if (m?.[1]) return m[1].trim();
  const m2 = html.match(/\bdata-sx-log-btn\s*=\s*'([^']+)'/i);
  return m2?.[1]?.trim();
}

export function stripActionLogMarkerFromStoredHtml(html: string): string {
  if (!html) return html;
  return html
    .replace(/<span\b[^>]*\bdata-sx-log-btn\s*=\s*(?:"[^"]*"|'[^']*')[^>]*>\s*<\/span>/gi, '')
    .replace(/<span\b[^>]*\bdata-sx-log-btn\s*=\s*(?:"[^"]*"|'[^']*')[^>]*\/\s*>/gi, '')
    .trimStart();
}

export interface IFormActionLogRuntimeContext {
  sourceListTitle: string;
  sourceItemId: number | null | undefined;
  formMode: TFormManagerFormMode;
  /** Cor resolvida do tema para realce no corpo HTML gravado. */
  logEntryAccentHex?: string;
}

function sanitizeInlineCssColor(color: string): string {
  return color.replace(/[<>'"`;{}\\]/g, '').trim();
}

function sortKeysDeep(v: unknown): unknown {
  if (v === null || typeof v !== 'object') return v;
  if (Array.isArray(v)) return v.map(sortKeysDeep);
  const o = v as Record<string, unknown>;
  const keys = Object.keys(o).sort();
  const out: Record<string, unknown> = {};
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    out[k] = sortKeysDeep(o[k]);
  }
  return out;
}

function lookupUserComparable(v: unknown): string {
  if (typeof v === 'number' && isFinite(v)) return `id:${v}`;
  if (v && typeof v === 'object' && !Array.isArray(v)) {
    const o = v as Record<string, unknown>;
    const id = o.Id ?? o.id;
    if (typeof id === 'number' && isFinite(id)) return `id:${id}`;
    if (typeof id === 'string' && id.trim()) return `id:${id.trim()}`;
  }
  return String(v);
}

function lookupUserDisplay(v: unknown): string {
  if (v === null || v === undefined || v === '') return '—';
  if (typeof v === 'number') return `#${v}`;
  if (v && typeof v === 'object' && !Array.isArray(v)) {
    const o = v as Record<string, unknown>;
    const t = o.Title ?? o.title;
    if (t != null && String(t).trim()) return String(t);
    const id = o.Id ?? o.id;
    if (typeof id === 'number') return `#${id}`;
  }
  return String(v);
}

function auditComparable(meta: IFieldMetadata | undefined, v: unknown): string {
  if (v === undefined || v === null || v === '') return '';
  const mt = meta?.MappedType ?? 'unknown';
  if (mt === 'lookup' || mt === 'user') return lookupUserComparable(v);
  if (mt === 'lookupmulti' || mt === 'usermulti') {
    const arr = Array.isArray(v) ? v : [];
    const ids: number[] = [];
    for (let i = 0; i < arr.length; i++) {
      const x = arr[i];
      if (typeof x === 'number' && isFinite(x)) ids.push(x);
      else if (x && typeof x === 'object' && 'Id' in (x as object)) {
        const id = (x as { Id?: unknown }).Id;
        if (typeof id === 'number' && isFinite(id)) ids.push(id);
      }
    }
    ids.sort((a, b) => a - b);
    return ids.join(',');
  }
  if (mt === 'multichoice') {
    const arr = Array.isArray(v)
      ? v.map(String)
      : String(v)
          .split(';')
          .map((s) => s.trim())
          .filter(Boolean);
    const u = Array.from(new Set(arr)).sort();
    return u.join('|');
  }
  if (mt === 'boolean') {
    return v === true || v === 1 || v === '1' || String(v).toLowerCase() === 'true' ? '1' : '0';
  }
  if (mt === 'datetime') {
    if (v instanceof Date && !isNaN(v.getTime())) return v.toISOString().slice(0, 16);
    const d = new Date(String(v));
    if (!isNaN(d.getTime())) return d.toISOString().slice(0, 16);
    return String(v);
  }
  if (mt === 'number' || mt === 'currency') {
    const n = typeof v === 'number' ? v : Number(String(v).replace(',', '.'));
    return isNaN(n) ? String(v) : String(n);
  }
  if (mt === 'url') {
    if (v && typeof v === 'object' && 'Url' in (v as object)) {
      const o = v as Record<string, unknown>;
      return `url:${String(o.Url ?? '').trim()}|${String(o.Description ?? '')}`;
    }
    return String(v);
  }
  if (mt === 'taxonomy' || mt === 'taxonomymulti') {
    if (v && typeof v === 'object' && !Array.isArray(v)) {
      const l = (v as Record<string, unknown>).Label ?? (v as Record<string, unknown>).WssId;
      return String(l ?? JSON.stringify(sortKeysDeep(v)));
    }
    if (Array.isArray(v)) {
      const labels = v
        .map((x) =>
          x && typeof x === 'object' && 'Label' in (x as object)
            ? String((x as Record<string, unknown>).Label ?? '')
            : String(x)
        )
        .filter(Boolean)
        .sort();
      return labels.join('|');
    }
    return String(v);
  }
  if (typeof v === 'object') {
    try {
      return JSON.stringify(sortKeysDeep(v));
    } catch {
      return String(v);
    }
  }
  return String(v);
}

function auditDisplay(meta: IFieldMetadata | undefined, v: unknown): string {
  if (v === undefined || v === null || v === '') return '—';
  const mt = meta?.MappedType ?? 'unknown';
  if (mt === 'lookup' || mt === 'user' || mt === 'lookupmulti' || mt === 'usermulti') {
    if (mt === 'lookupmulti' || mt === 'usermulti') {
      const arr = Array.isArray(v) ? v : [];
      const parts: string[] = [];
      for (let i = 0; i < arr.length; i++) {
        parts.push(lookupUserDisplay(arr[i]));
      }
      return parts.filter((s) => s && s !== '—').join('; ') || '—';
    }
    return lookupUserDisplay(v);
  }
  if (mt === 'boolean') {
    return v === true || v === 1 || v === '1' || String(v).toLowerCase() === 'true' ? 'Sim' : 'Não';
  }
  if (mt === 'datetime') {
    if (v instanceof Date && !isNaN(v.getTime())) return v.toLocaleString('pt-PT');
    const d = new Date(String(v));
    if (!isNaN(d.getTime())) return d.toLocaleString('pt-PT');
    return String(v);
  }
  if (mt === 'url' && v && typeof v === 'object' && 'Url' in (v as object)) {
    const o = v as Record<string, unknown>;
    const u = String(o.Url ?? '').trim();
    const desc = String(o.Description ?? '').trim();
    if (!u) return '—';
    return desc ? `${u} (${desc})` : u;
  }
  if (mt === 'multichoice') {
    const arr = Array.isArray(v)
      ? v.map(String)
      : String(v)
          .split(';')
          .map((s) => s.trim())
          .filter(Boolean);
    return arr.length ? arr.join('; ') : '—';
  }
  if (mt === 'taxonomy' || mt === 'taxonomymulti') {
    if (Array.isArray(v)) {
      const parts = v.map((x) =>
        x && typeof x === 'object' && 'Label' in (x as object)
          ? String((x as Record<string, unknown>).Label ?? '')
          : String(x)
      );
      return parts.filter(Boolean).join('; ') || '—';
    }
    if (v && typeof v === 'object' && 'Label' in (v as object)) {
      return String((v as Record<string, unknown>).Label ?? '—');
    }
  }
  if (typeof v === 'object') {
    try {
      return JSON.stringify(v);
    } catch {
      return String(v);
    }
  }
  return String(v);
}

export interface IFormActionLogAutomaticChangesInput {
  baseline: Record<string, unknown>;
  final: Record<string, unknown>;
  payloadFieldInternalNames: string[];
  metaByName: ReadonlyMap<string, IFieldMetadata>;
}

export function buildAutomaticFieldChangesHtml(input: IFormActionLogAutomaticChangesInput): string {
  const { baseline, final, payloadFieldInternalNames, metaByName } = input;
  const lines: string[] = [];
  for (let i = 0; i < payloadFieldInternalNames.length; i++) {
    const name = payloadFieldInternalNames[i];
    const meta = metaByName.get(name);
    if (meta?.Hidden || meta?.ReadOnlyField) continue;
    const b = baseline[name];
    const f = final[name];
    if (auditComparable(meta, b) === auditComparable(meta, f)) continue;
    const label = (meta?.Title ?? name).trim() || name;
    lines.push(
      `<p><b>${escapeHtml(label)}:</b> de ${escapeHtml(auditDisplay(meta, b))} para ${escapeHtml(
        auditDisplay(meta, f)
      )}</p>`
    );
  }
  return lines.join('\n');
}

export interface IAppendFormActionLogEntryOpts {
  automaticChanges?: IFormActionLogAutomaticChangesInput;
}

export async function appendFormActionLogEntry(
  itemsService: ItemsService,
  actionLog: IFormManagerActionLogConfig | undefined,
  btn: IFormCustomButtonConfig,
  ctx: IFormActionLogRuntimeContext,
  opts?: IAppendFormActionLogEntryOpts
): Promise<void> {
  if (!actionLog?.captureEnabled) return;
  const logList = actionLog.listTitle?.trim();
  const fieldName = actionLog.actionFieldInternalName?.trim();
  const linkField = actionLog.sourceListLookupFieldInternalName?.trim();
  if (!logList || !fieldName) return;

  if (linkField) {
    const sid = ctx.sourceItemId;
    if (sid === undefined || sid === null || typeof sid !== 'number' || !isFinite(sid)) {
      return;
    }
  }

  const customHtml = (actionLog.descriptionsHtmlByButtonId?.[btn.id] ?? '').trim();
  let autoHtml = '';
  if (
    actionLog.automaticChangesOnUpdate === true &&
    btn.operation === 'update' &&
    opts?.automaticChanges
  ) {
    autoHtml = buildAutomaticFieldChangesHtml(opts.automaticChanges).trim();
  }
  const meta = `<p style="color:#605e5c;font-size:12px"><em>Lista de origem</em>: ${escapeHtml(
    ctx.sourceListTitle
  )} · <em>Item</em>: ${ctx.sourceItemId ?? '—'} · <em>Modo</em>: ${escapeHtml(String(ctx.formMode))}</p>`;
  const marker = `<span ${FORM_ACTION_LOG_BTN_DATA_ATTR}="${escapeHtml(
    btn.id
  )}" style="display:none!important" aria-hidden="true"></span>`;
  const intro =
    customHtml.length > 0 ? customHtml : `<p>${escapeHtml(btn.label || btn.id)}</p>`;
  let body = marker + (autoHtml.length > 0 ? `${intro}\n${autoHtml}\n${meta}` : `${intro}\n${meta}`);
  const accentRaw = (ctx.logEntryAccentHex ?? '').trim();
  if (accentRaw) {
    const accent = sanitizeInlineCssColor(accentRaw);
    if (accent) {
      body = `<div style="border-left:4px solid ${accent};padding-left:12px;margin:0 0 8px 0">${body}</div>`;
    }
  }

  const titleBase = (btn.label || btn.id).slice(0, 200);
  const title = `${titleBase} · ${new Date().toLocaleString('pt-PT')}`.slice(0, 255);

  const payload: Record<string, unknown> = {
    Title: title,
    [fieldName]: body,
  };
  if (linkField) {
    payload[`${linkField}Id`] = ctx.sourceItemId as number;
  }

  await itemsService.addItem(logList, payload);
}
