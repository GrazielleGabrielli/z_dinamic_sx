import type {
  IFormCustomButtonConfig,
  IFormManagerActionLogConfig,
  TFormManagerFormMode,
} from '../config/types/formManager';
import type { ItemsService } from '../../../../services';

function escapeHtml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

export interface IFormActionLogRuntimeContext {
  sourceListTitle: string;
  sourceItemId: number | null | undefined;
  formMode: TFormManagerFormMode;
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
  const body =
    customHtml.length > 0 ? `${customHtml}\n${meta}` : `<p>${escapeHtml(btn.label || btn.id)}</p>\n${meta}`;

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
