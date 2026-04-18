import type {
  IFormCustomButtonConfig,
  IFormManagerActionLogConfig,
  TFormManagerFormMode,
} from '../config/types/formManager';
import type { ItemsService } from '../../../../services';

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

export async function appendFormActionLogEntry(
  itemsService: ItemsService,
  actionLog: IFormManagerActionLogConfig | undefined,
  btn: IFormCustomButtonConfig,
  ctx: IFormActionLogRuntimeContext
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
  const meta = `<p style="color:#605e5c;font-size:12px"><em>Lista de origem</em>: ${escapeHtml(
    ctx.sourceListTitle
  )} · <em>Item</em>: ${ctx.sourceItemId ?? '—'} · <em>Modo</em>: ${escapeHtml(String(ctx.formMode))}</p>`;
  const marker = `<span ${FORM_ACTION_LOG_BTN_DATA_ATTR}="${escapeHtml(
    btn.id
  )}" style="display:none!important" aria-hidden="true"></span>`;
  let body =
    marker +
    (customHtml.length > 0 ? `${customHtml}\n${meta}` : `<p>${escapeHtml(btn.label || btn.id)}</p>\n${meta}`);
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
