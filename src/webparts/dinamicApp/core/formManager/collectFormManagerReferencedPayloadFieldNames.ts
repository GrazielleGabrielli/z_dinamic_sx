import type { IFormManagerConfig, TFormButtonAction, TFormRule } from '../config/types/formManager';
import { FORM_ATTACHMENTS_FIELD_INTERNAL, FORM_BANNER_INTERNAL_PREFIX } from '../config/types/formManager';

const JOIN_PH_RE = /\{\{([^}]+)\}\}/g;

function isSyntheticFieldName(name: string): boolean {
  return name === FORM_ATTACHMENTS_FIELD_INTERNAL || name.startsWith(FORM_BANNER_INTERNAL_PREFIX);
}

function addFieldName(out: Set<string>, raw: string | undefined): void {
  const t = typeof raw === 'string' ? raw.trim() : '';
  if (!t || isSyntheticFieldName(t)) return;
  out.add(t);
}

function collectFromJoinTemplate(tpl: string | undefined, out: Set<string>): void {
  if (!tpl) return;
  JOIN_PH_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = JOIN_PH_RE.exec(tpl)) !== null) {
    const full = String(m[1] ?? '').trim();
    if (!full) continue;
    const base = full.split('/')[0]?.trim();
    addFieldName(out, base);
  }
}

function collectFromValueTemplate(tpl: string | undefined, out: Set<string>): void {
  if (!tpl) return;
  JOIN_PH_RE.lastIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = JOIN_PH_RE.exec(tpl)) !== null) {
    const full = String(m[1] ?? '').trim();
    if (!full) continue;
    const base = full.split('/')[0]?.trim();
    addFieldName(out, base);
  }
}

function collectFromButtonActions(actions: TFormButtonAction[] | undefined, out: Set<string>): void {
  if (!actions?.length) return;
  for (let i = 0; i < actions.length; i++) {
    const a = actions[i];
    switch (a.kind) {
      case 'setFieldValue':
        addFieldName(out, a.field);
        break;
      case 'joinFields':
        addFieldName(out, a.targetField);
        for (let j = 0; j < (a.sourceFields ?? []).length; j++) {
          addFieldName(out, a.sourceFields[j]);
        }
        collectFromJoinTemplate(a.valueTemplate, out);
        break;
      case 'showFields':
        for (let j = 0; j < (a.fields ?? []).length; j++) {
          addFieldName(out, a.fields[j]);
        }
        break;
      case 'hideFields':
        for (let j = 0; j < (a.fields ?? []).length; j++) {
          addFieldName(out, a.fields[j]);
        }
        break;
      default:
        break;
    }
  }
}

function collectFromRules(rules: TFormRule[] | undefined, out: Set<string>): void {
  if (!rules?.length) return;
  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    if (r.enabled === false) continue;
    switch (r.action) {
      case 'setVisibility':
        if (r.targetKind === 'field') addFieldName(out, r.targetId);
        break;
      case 'setRequired':
      case 'setDisabled':
      case 'setReadOnly':
      case 'setDefault':
      case 'validateValue':
      case 'multiMinMax':
      case 'setEffectiveSection':
      case 'profileVisibility':
      case 'profileEditable':
      case 'profileRequired':
      case 'authorFieldAccess':
      case 'asyncUniqueness':
        addFieldName(out, r.field);
        if (r.action === 'setDefault' && typeof r.value === 'string') {
          collectFromValueTemplate(r.value, out);
        }
        break;
      case 'clearFields':
        addFieldName(out, r.triggerField);
        for (let j = 0; j < (r.fields ?? []).length; j++) {
          addFieldName(out, r.fields[j]);
        }
        break;
      case 'validateDate':
        addFieldName(out, r.field);
        addFieldName(out, r.gteField);
        addFieldName(out, r.lteField);
        addFieldName(out, r.gtField);
        addFieldName(out, r.ltField);
        break;
      case 'atLeastOne':
        for (let j = 0; j < (r.fields ?? []).length; j++) {
          addFieldName(out, r.fields[j]);
        }
        break;
      case 'filterLookupOptions':
        addFieldName(out, r.field);
        addFieldName(out, r.parentField);
        addFieldName(out, r.childField);
        break;
      case 'setComputed':
        addFieldName(out, r.field);
        collectFromJoinTemplate(r.expression, out);
        break;
      case 'showMessage':
      case 'attachmentRules':
      case 'asyncCountLimit':
        break;
      default:
        break;
    }
  }
}

export function collectFormManagerReferencedPayloadFieldNames(cfg: IFormManagerConfig): string[] {
  const out = new Set<string>();
  const buttons = cfg.customButtons ?? [];
  for (let i = 0; i < buttons.length; i++) {
    const b = buttons[i];
    collectFromButtonActions(b.actions, out);
    addFieldName(out, b.confirmBeforeRun?.promptFieldInternalName);
  }
  collectFromButtonActions(cfg.historyButtonActions, out);
  collectFromRules(cfg.rules, out);
  const dh = cfg.dynamicHelp ?? [];
  for (let i = 0; i < dh.length; i++) {
    addFieldName(out, dh[i].field);
  }
  const pb = cfg.permissionBreak?.assignments ?? [];
  for (let i = 0; i < pb.length; i++) {
    const a = pb[i];
    if (a.kind === 'field') addFieldName(out, a.fieldInternalName);
  }
  return Array.from(out);
}
