import type {
  TFormConditionNode,
  TFormConditionOp,
  TFormManagerFormMode,
  TFormRule,
  IFormCompareRef,
  IFormRuleAttachment,
} from '../config/types/formManager';

export const CONDITION_OP_OPTIONS: { key: TFormConditionOp; text: string }[] = [
  { key: 'eq', text: 'é igual a' },
  { key: 'ne', text: 'é diferente de' },
  { key: 'contains', text: 'contém' },
  { key: 'startsWith', text: 'começa com' },
  { key: 'endsWith', text: 'termina com' },
  { key: 'gt', text: 'maior que' },
  { key: 'ge', text: 'maior ou igual a' },
  { key: 'lt', text: 'menor que' },
  { key: 'le', text: 'menor ou igual a' },
  { key: 'isEmpty', text: 'está vazio' },
  { key: 'isFilled', text: 'não está vazio' },
  { key: 'isTrue', text: 'é verdadeiro' },
  { key: 'isFalse', text: 'é falso' },
];

export function conditionOpLabel(op: TFormConditionOp): string {
  for (let i = 0; i < CONDITION_OP_OPTIONS.length; i++) {
    if (CONDITION_OP_OPTIONS[i].key === op) return CONDITION_OP_OPTIONS[i].text;
  }
  return op;
}

export function safeIdSegment(internalName: string): string {
  return internalName.replace(/[^a-zA-Z0-9]/g, '_');
}

const RE_FIELD_PREFIX = /^ui_f_([a-zA-Z0-9_]+)_/;
const RE_CARD_RULE = /^ui_card_([a-zA-Z0-9]+)_(\d+)_(.+)$/;
const UI_ATTACHMENT_ID = 'ui_form_attachment';

export function isFieldUiRuleId(id: string): boolean {
  return id.indexOf('ui_f_') === 0;
}

export function isCardUiRuleId(id: string): boolean {
  return id.indexOf('ui_card_') === 0;
}

export function uiSegmentFromFieldRuleId(id: string): string | undefined {
  const m = id.match(RE_FIELD_PREFIX);
  return m ? m[1] : undefined;
}

export function stripFieldUiRules(rules: TFormRule[], internalName: string): TFormRule[] {
  const seg = safeIdSegment(internalName);
  const prefix = `ui_f_${seg}_`;
  return rules.filter((r) => r.id.indexOf(prefix) !== 0);
}

export function stripAllFieldUiRules(rules: TFormRule[]): TFormRule[] {
  return rules.filter((r) => !isFieldUiRuleId(r.id));
}

export function stripCardUiRules(rules: TFormRule[]): TFormRule[] {
  return rules.filter((r) => !isCardUiRuleId(r.id));
}

export function isManagedUiRuleId(id: string): boolean {
  return isFieldUiRuleId(id) || isCardUiRuleId(id) || id === UI_ATTACHMENT_ID;
}

export function customRulesOnly(rules: TFormRule[]): TFormRule[] {
  return rules.filter((r) => !isManagedUiRuleId(r.id));
}

function leafWhen(
  field: string,
  op: TFormConditionOp,
  compare?: IFormCompareRef
): TFormConditionNode {
  return compare ? { kind: 'leaf', field, op, compare } : { kind: 'leaf', field, op };
}

export type TCompareUiKind = 'literal' | 'field' | 'token';

export interface IWhenUi {
  field: string;
  op: TFormConditionOp;
  compareKind: TCompareUiKind;
  compareValue: string;
}

export function whenUiToNode(w: IWhenUi): TFormConditionNode {
  const needsCompare =
    w.op !== 'isEmpty' &&
    w.op !== 'isFilled' &&
    w.op !== 'isTrue' &&
    w.op !== 'isFalse';
  if (!needsCompare) return leafWhen(w.field, w.op);
  const compare: IFormCompareRef = {
    kind: w.compareKind,
    value: w.compareValue,
  };
  return leafWhen(w.field, w.op, compare);
}

export function whenNodeToUi(node: TFormConditionNode | undefined): IWhenUi | undefined {
  if (!node || node.kind !== 'leaf') return undefined;
  const c = node.compare;
  return {
    field: node.field,
    op: node.op,
    compareKind: (c?.kind as TCompareUiKind) ?? 'literal',
    compareValue: c?.value ?? '',
  };
}

export type TConditionalEffectKind =
  | 'showField'
  | 'hideField'
  | 'requireField'
  | 'optionalField'
  | 'disableField'
  | 'enableField'
  | 'readonlyField'
  | 'editableField'
  | 'message';

export const CONDITIONAL_EFFECT_OPTIONS: { key: TConditionalEffectKind; text: string }[] = [
  { key: 'showField', text: 'Mostrar campo' },
  { key: 'hideField', text: 'Ocultar campo' },
  { key: 'requireField', text: 'Tornar obrigatório' },
  { key: 'optionalField', text: 'Tornar opcional' },
  { key: 'disableField', text: 'Desativar campo' },
  { key: 'enableField', text: 'Ativar campo' },
  { key: 'readonlyField', text: 'Somente leitura' },
  { key: 'editableField', text: 'Permitir edição' },
  { key: 'message', text: 'Exibir mensagem' },
];

export interface IConditionalEffectUi {
  kind: TConditionalEffectKind;
  targetField?: string;
  messageVariant?: 'info' | 'warning' | 'error';
  messageText?: string;
}

export interface IConditionalRuleCard {
  id: string;
  enabled?: boolean;
  when: IWhenUi;
  modes?: TFormManagerFormMode[];
  /** Títulos de grupos SharePoint; vazio = todos os utilizadores. */
  groupTitles?: string[];
  effects: IConditionalEffectUi[];
}

export function newCardId(): string {
  return `c${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
}

export function compileConditionalCard(card: IConditionalRuleCard): TFormRule[] {
  if (card.enabled === false) return [];
  const when = whenUiToNode(card.when);
  const out: TFormRule[] = [];
  let idx = 0;
  for (let i = 0; i < card.effects.length; i++) {
    const e = card.effects[i];
    const base = {
      when,
      ...(card.modes?.length ? { modes: card.modes } : {}),
      ...(card.groupTitles?.length ? { groupTitles: card.groupTitles } : {}),
    };
    const id = (suffix: string): string => `ui_card_${card.id}_${idx++}_${suffix}`;
    const tf = (e.targetField ?? '').trim();
    switch (e.kind) {
      case 'showField':
        if (tf)
          out.push({
            id: id('vis'),
            action: 'setVisibility',
            targetKind: 'field',
            targetId: tf,
            visibility: 'show',
            ...base,
          });
        break;
      case 'hideField':
        if (tf)
          out.push({
            id: id('hid'),
            action: 'setVisibility',
            targetKind: 'field',
            targetId: tf,
            visibility: 'hide',
            ...base,
          });
        break;
      case 'requireField':
        if (tf) out.push({ id: id('req'), action: 'setRequired', field: tf, required: true, ...base });
        break;
      case 'optionalField':
        if (tf) out.push({ id: id('opt'), action: 'setRequired', field: tf, required: false, ...base });
        break;
      case 'disableField':
        if (tf) out.push({ id: id('dis'), action: 'setDisabled', field: tf, disabled: true, ...base });
        break;
      case 'enableField':
        if (tf) out.push({ id: id('ena'), action: 'setDisabled', field: tf, disabled: false, ...base });
        break;
      case 'readonlyField':
        if (tf) out.push({ id: id('ro'), action: 'setReadOnly', field: tf, readOnly: true, ...base });
        break;
      case 'editableField':
        if (tf) out.push({ id: id('rw'), action: 'setReadOnly', field: tf, readOnly: false, ...base });
        break;
      case 'message':
        if (e.messageText && e.messageText.trim())
          out.push({
            id: id('msg'),
            action: 'showMessage',
            variant: e.messageVariant ?? 'info',
            text: e.messageText.trim(),
            ...base,
          });
        break;
      default:
        break;
    }
  }
  return out;
}

function effectFromRule(r: TFormRule): IConditionalEffectUi | undefined {
  switch (r.action) {
    case 'setVisibility':
      if (r.targetKind !== 'field' || !r.targetId) return undefined;
      return r.visibility === 'show'
        ? { kind: 'showField', targetField: r.targetId }
        : { kind: 'hideField', targetField: r.targetId };
    case 'setRequired':
      return r.required
        ? { kind: 'requireField', targetField: r.field }
        : { kind: 'optionalField', targetField: r.field };
    case 'setDisabled':
      return r.disabled
        ? { kind: 'disableField', targetField: r.field }
        : { kind: 'enableField', targetField: r.field };
    case 'setReadOnly':
      return r.readOnly
        ? { kind: 'readonlyField', targetField: r.field }
        : { kind: 'editableField', targetField: r.field };
    case 'showMessage':
      return {
        kind: 'message',
        messageVariant: r.variant,
        messageText: r.text,
      };
    default:
      return undefined;
  }
}

export function parseConditionalCardsFromRules(rules: TFormRule[]): {
  cards: IConditionalRuleCard[];
  cardRuleIds: Set<string>;
} {
  const byCard = new Map<string, TFormRule[]>();
  const cardRuleIds = new Set<string>();
  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    const m = r.id.match(RE_CARD_RULE);
    if (!m) continue;
    const cardId = m[1];
    cardRuleIds.add(r.id);
    const arr = byCard.get(cardId) ?? [];
    arr.push(r);
    byCard.set(cardId, arr);
  }
  const cards: IConditionalRuleCard[] = [];
  byCard.forEach((list, cardId) => {
    list.sort((a, b) => {
      const ma = a.id.match(RE_CARD_RULE);
      const mb = b.id.match(RE_CARD_RULE);
      const ia = ma ? parseInt(ma[2], 10) : 0;
      const ib = mb ? parseInt(mb[2], 10) : 0;
      return ia - ib;
    });
    const first = list[0];
    const w = whenNodeToUi(first.when);
    if (!w) return;
    const modes = first.modes;
    const groupTitles = first.groupTitles;
    const effects: IConditionalEffectUi[] = [];
    for (let j = 0; j < list.length; j++) {
      const eff = effectFromRule(list[j]);
      if (eff) effects.push(eff);
    }
    cards.push({
      id: cardId,
      when: w,
      ...(modes?.length ? { modes } : {}),
      ...(groupTitles?.length ? { groupTitles } : {}),
      effects,
    });
  });
  return { cards, cardRuleIds };
}

export function mergeCardRulesIntoAll(
  baseRules: TFormRule[],
  cards: IConditionalRuleCard[]
): TFormRule[] {
  const without = stripCardUiRules(baseRules);
  const compiled: TFormRule[] = [];
  for (let i = 0; i < cards.length; i++) {
    compiled.push(...compileConditionalCard(cards[i]));
  }
  return without.concat(compiled);
}

export interface IFieldRuleEditorState {
  modes: TFormManagerFormMode[];
  defaultValue: string;
  validateValue: {
    minLength: string;
    maxLength: string;
    minNumber: string;
    maxNumber: string;
    pattern: string;
    patternMessage: string;
  };
  validateDate: {
    minDaysFromToday: string;
    maxDaysFromToday: string;
    blockWeekends: boolean;
    gteField: string;
    lteField: string;
    message: string;
  };
  filterLookup: {
    parentField: string;
    odataFilterTemplate: string;
  };
  computedExpression: string;
  /** Id do nó na árvore de pastas (Anexos); gera expressão `attfolder:id`. */
  computedAttachmentFolderNodeId: string;
  disableWhenActive: boolean;
  disableWhenUi: IWhenUi;
  enableWhenActive: boolean;
  enableWhenUi: IWhenUi;
}

export function mergeFieldRuleEditorState(
  base: IFieldRuleEditorState,
  patch: Partial<IFieldRuleEditorState>
): IFieldRuleEditorState {
  return {
    ...base,
    ...patch,
    validateValue: { ...base.validateValue, ...(patch.validateValue ?? {}) },
    validateDate: { ...base.validateDate, ...(patch.validateDate ?? {}) },
    filterLookup: { ...base.filterLookup, ...(patch.filterLookup ?? {}) },
    disableWhenUi: { ...base.disableWhenUi, ...(patch.disableWhenUi ?? {}) },
    enableWhenUi: { ...base.enableWhenUi, ...(patch.enableWhenUi ?? {}) },
  };
}

export function emptyFieldRuleEditorState(): IFieldRuleEditorState {
  return {
    modes: [],
    defaultValue: '',
    validateValue: {
      minLength: '',
      maxLength: '',
      minNumber: '',
      maxNumber: '',
      pattern: '',
      patternMessage: '',
    },
    validateDate: {
      minDaysFromToday: '',
      maxDaysFromToday: '',
      blockWeekends: false,
      gteField: '',
      lteField: '',
      message: '',
    },
    filterLookup: { parentField: '', odataFilterTemplate: '' },
    computedExpression: '',
    computedAttachmentFolderNodeId: '',
    disableWhenActive: false,
    disableWhenUi: { field: 'Title', op: 'eq', compareKind: 'literal', compareValue: '' },
    enableWhenActive: false,
    enableWhenUi: { field: 'Title', op: 'eq', compareKind: 'literal', compareValue: '' },
  };
}

export function fieldRuleStateFromRules(
  internalName: string,
  rules: TFormRule[]
): IFieldRuleEditorState {
  const st = emptyFieldRuleEditorState();
  const seg = safeIdSegment(internalName);
  const mine = rules.filter((r) => r.id.indexOf(`ui_f_${seg}_`) === 0);
  for (let i = 0; i < mine.length; i++) {
    const r = mine[i];
    if (r.action === 'setDefault' && r.field === internalName) st.defaultValue = r.value;
    if (r.action === 'validateValue' && r.field === internalName) {
      if (r.minLength !== undefined) st.validateValue.minLength = String(r.minLength);
      if (r.maxLength !== undefined) st.validateValue.maxLength = String(r.maxLength);
      if (r.minNumber !== undefined) st.validateValue.minNumber = String(r.minNumber);
      if (r.maxNumber !== undefined) st.validateValue.maxNumber = String(r.maxNumber);
      if (r.pattern) st.validateValue.pattern = r.pattern;
      if (r.patternMessage) st.validateValue.patternMessage = r.patternMessage;
    }
    if (r.action === 'validateDate' && r.field === internalName) {
      if (r.minDaysFromToday !== undefined) st.validateDate.minDaysFromToday = String(r.minDaysFromToday);
      if (r.maxDaysFromToday !== undefined) st.validateDate.maxDaysFromToday = String(r.maxDaysFromToday);
      if (r.blockWeekends) st.validateDate.blockWeekends = true;
      if (r.gteField) st.validateDate.gteField = r.gteField;
      if (r.lteField) st.validateDate.lteField = r.lteField;
      if (r.message) st.validateDate.message = r.message;
    }
    if (r.action === 'filterLookupOptions' && r.field === internalName) {
      st.filterLookup.parentField = r.parentField;
      st.filterLookup.odataFilterTemplate = r.odataFilterTemplate;
    }
    if (r.action === 'setComputed' && r.field === internalName) {
      const ex = String(r.expression ?? '').trim();
      if (ex.indexOf('attfolder:') === 0) {
        st.computedAttachmentFolderNodeId = ex.slice('attfolder:'.length).trim();
        st.computedExpression = '';
      } else {
        st.computedExpression = ex;
        st.computedAttachmentFolderNodeId = '';
      }
    }
    if (r.action === 'setDisabled' && r.field === internalName && r.when) {
      const w = whenNodeToUi(r.when);
      if (w) {
        if (r.id === `ui_f_${seg}_discond`) {
          st.disableWhenActive = true;
          st.disableWhenUi = w;
        } else if (r.id === `ui_f_${seg}_enacond`) {
          st.enableWhenActive = true;
          st.enableWhenUi = w;
        }
      }
    }
    if (r.modes && r.modes.length && st.modes.length === 0 && r.action !== 'setComputed') {
      st.modes = r.modes.slice();
    }
  }
  return st;
}

function numOrUndef(s: string): number | undefined {
  const t = s.trim();
  if (!t) return undefined;
  const n = Number(t);
  return isNaN(n) ? undefined : n;
}

export function buildFieldUiRules(internalName: string, st: IFieldRuleEditorState): TFormRule[] {
  const seg = safeIdSegment(internalName);
  const id = (s: string): string => `ui_f_${seg}_${s}`;
  const baseModes = st.modes.length ? { modes: st.modes } : {};
  const out: TFormRule[] = [];

  if (st.defaultValue.trim())
    out.push({
      id: id('def'),
      action: 'setDefault',
      field: internalName,
      value: st.defaultValue.trim(),
      when: { kind: 'leaf', field: internalName, op: 'isEmpty' },
      ...baseModes,
    });

  const vv = st.validateValue;
  const hasVal =
    vv.minLength ||
    vv.maxLength ||
    vv.minNumber ||
    vv.maxNumber ||
    vv.pattern;
  if (hasVal) {
    out.push({
      id: id('val'),
      action: 'validateValue',
      field: internalName,
      ...(numOrUndef(vv.minLength) !== undefined ? { minLength: numOrUndef(vv.minLength) } : {}),
      ...(numOrUndef(vv.maxLength) !== undefined ? { maxLength: numOrUndef(vv.maxLength) } : {}),
      ...(numOrUndef(vv.minNumber) !== undefined ? { minNumber: numOrUndef(vv.minNumber) } : {}),
      ...(numOrUndef(vv.maxNumber) !== undefined ? { maxNumber: numOrUndef(vv.maxNumber) } : {}),
      ...(vv.pattern.trim() ? { pattern: vv.pattern.trim() } : {}),
      ...(vv.patternMessage.trim() ? { patternMessage: vv.patternMessage.trim() } : {}),
      ...baseModes,
    });
  }

  const vd = st.validateDate;
  const hasDate =
    vd.minDaysFromToday ||
    vd.maxDaysFromToday ||
    vd.blockWeekends ||
    vd.gteField.trim() ||
    vd.lteField.trim();
  if (hasDate) {
    out.push({
      id: id('date'),
      action: 'validateDate',
      field: internalName,
      ...(numOrUndef(vd.minDaysFromToday) !== undefined ? { minDaysFromToday: numOrUndef(vd.minDaysFromToday) } : {}),
      ...(numOrUndef(vd.maxDaysFromToday) !== undefined ? { maxDaysFromToday: numOrUndef(vd.maxDaysFromToday) } : {}),
      ...(vd.blockWeekends ? { blockWeekends: true } : {}),
      ...(vd.gteField.trim() ? { gteField: vd.gteField.trim() } : {}),
      ...(vd.lteField.trim() ? { lteField: vd.lteField.trim() } : {}),
      ...(vd.message.trim() ? { message: vd.message.trim() } : {}),
      ...baseModes,
    });
  }

  const fl = st.filterLookup;
  if (fl.parentField.trim() && fl.odataFilterTemplate.trim()) {
    out.push({
      id: id('lk'),
      action: 'filterLookupOptions',
      field: internalName,
      parentField: fl.parentField.trim(),
      odataFilterTemplate: fl.odataFilterTemplate.trim(),
      ...baseModes,
    });
  }

  const attFolderId = st.computedAttachmentFolderNodeId.trim();
  const cmpExpr = attFolderId ? `attfolder:${attFolderId}` : st.computedExpression.trim();
  if (cmpExpr) {
    out.push({
      id: id('cmp'),
      action: 'setComputed',
      field: internalName,
      expression: cmpExpr,
    });
  }

  if (st.disableWhenActive && st.disableWhenUi.field.trim()) {
    out.push({
      id: id('discond'),
      action: 'setDisabled',
      field: internalName,
      disabled: true,
      when: whenUiToNode(st.disableWhenUi),
      ...baseModes,
    });
  }
  if (st.enableWhenActive && st.enableWhenUi.field.trim()) {
    out.push({
      id: id('enacond'),
      action: 'setDisabled',
      field: internalName,
      disabled: false,
      when: whenUiToNode(st.enableWhenUi),
      ...baseModes,
    });
  }

  return out;
}

export function mergeFieldRules(
  allRules: TFormRule[],
  internalName: string,
  newFieldRules: TFormRule[]
): TFormRule[] {
  return stripFieldUiRules(allRules, internalName).concat(newFieldRules);
}

export function stripAttachmentUiRule(rules: TFormRule[]): TFormRule[] {
  return rules.filter((r) => r.id !== UI_ATTACHMENT_ID);
}

function normAttachmentExtensions(arr: string[] | undefined): string[] {
  if (!arr || !arr.length) return [];
  const out: string[] = [];
  const seen: Record<string, boolean> = {};
  for (let i = 0; i < arr.length; i++) {
    const e = String(arr[i]).trim().replace(/^\./, '').toLowerCase();
    if (!e || seen[e]) continue;
    seen[e] = true;
    out.push(e);
  }
  return out;
}

export function buildAttachmentUiRule(opts: {
  minCount?: number;
  maxCount?: number;
  message?: string;
  allowedFileExtensions?: string[];
}): TFormRule | undefined {
  const { minCount, maxCount, message, allowedFileExtensions } = opts;
  const ext = normAttachmentExtensions(allowedFileExtensions);
  if (
    minCount === undefined &&
    maxCount === undefined &&
    ext.length === 0 &&
    !(message && message.trim())
  ) {
    return undefined;
  }
  return {
    id: UI_ATTACHMENT_ID,
    action: 'attachmentRules',
    ...(minCount !== undefined ? { minCount } : {}),
    ...(maxCount !== undefined ? { maxCount } : {}),
    ...(message?.trim() ? { message: message.trim() } : {}),
    ...(ext.length ? { allowedFileExtensions: ext } : {}),
  };
}

export function parseAttachmentUiRule(rules: TFormRule[]): {
  minCount: string;
  maxCount: string;
  message: string;
  allowedFileExtensions: string[];
} {
  let r: TFormRule | undefined;
  for (let i = 0; i < rules.length; i++) {
    const x = rules[i];
    if (x.id === UI_ATTACHMENT_ID && x.action === 'attachmentRules') {
      r = x;
      break;
    }
  }
  if (!r || r.action !== 'attachmentRules')
    return { minCount: '', maxCount: '', message: '', allowedFileExtensions: [] };
  const att = r as IFormRuleAttachment;
  return {
    minCount: r.minCount !== undefined ? String(r.minCount) : '',
    maxCount: r.maxCount !== undefined ? String(r.maxCount) : '',
    message: r.message ?? '',
    allowedFileExtensions: normAttachmentExtensions(att.allowedFileExtensions),
  };
}

export function mergeAttachmentUiRule(
  rules: TFormRule[],
  opts: {
    minCount?: number;
    maxCount?: number;
    message?: string;
    allowedFileExtensions?: string[];
  }
): TFormRule[] {
  const next = stripAttachmentUiRule(rules);
  const a = buildAttachmentUiRule(opts);
  return a ? next.concat([a]) : next;
}

export function templateFieldRulesDateNotPast(): Partial<IFieldRuleEditorState> {
  return {
    validateDate: {
      minDaysFromToday: '0',
      maxDaysFromToday: '',
      blockWeekends: false,
      gteField: '',
      lteField: '',
      message: 'Não é permitida data no passado.',
    },
  };
}

export function templateFieldRulesEmail(): Partial<IFieldRuleEditorState> {
  return {
    validateValue: {
      minLength: '',
      maxLength: '',
      minNumber: '',
      maxNumber: '',
      pattern: '^[^\\s@]+@[^\\s@]+\\.[^\\s@]+$',
      patternMessage: 'Informe um e-mail válido.',
    },
  };
}

export function templateFieldRulesChoiceRequiresOther(whenField: string, whenValue: string, targetField: string): IConditionalRuleCard {
  return {
    id: newCardId(),
    when: {
      field: whenField,
      op: 'eq',
      compareKind: 'literal',
      compareValue: whenValue,
    },
    effects: [{ kind: 'requireField', targetField }],
  };
}

export function templateConditionalShowWhenEquals(
  whenField: string,
  whenValue: string,
  targetField: string
): IConditionalRuleCard {
  return {
    id: newCardId(),
    when: {
      field: whenField,
      op: 'eq',
      compareKind: 'literal',
      compareValue: whenValue,
    },
    effects: [{ kind: 'showField', targetField }],
  };
}

/** Cartão «mostrar campo alvo» com operador livre (contém, diferente, maior, …). */
export function templateConditionalShowWhenCompare(
  whenField: string,
  op: TFormConditionOp,
  compareValue: string,
  targetField: string
): IConditionalRuleCard {
  const noCompare =
    op === 'isEmpty' || op === 'isFilled' || op === 'isTrue' || op === 'isFalse';
  return {
    id: newCardId(),
    when: {
      field: whenField,
      op,
      compareKind: 'literal',
      compareValue: noCompare ? '' : compareValue,
    },
    effects: [{ kind: 'showField', targetField }],
  };
}

export function countFieldUiRules(internalName: string, rules: TFormRule[]): number {
  const seg = safeIdSegment(internalName);
  const prefix = `ui_f_${seg}_`;
  let n = 0;
  for (let i = 0; i < rules.length; i++) {
    if (rules[i].id.indexOf(prefix) === 0) n++;
  }
  return n;
}

export function describeRuleShort(rule: TFormRule): string {
  switch (rule.action) {
    case 'setVisibility':
      return `${rule.visibility === 'show' ? 'Mostrar' : 'Ocultar'} ${rule.targetKind} ${rule.targetId}`;
    case 'setRequired':
      return rule.required ? `Obrigar ${rule.field}` : `Opcional ${rule.field}`;
    case 'validateValue':
      return `Validar ${rule.field}`;
    case 'validateDate':
      return `Validar data ${rule.field}`;
    case 'showMessage':
      return rule.text.slice(0, 40) + (rule.text.length > 40 ? '…' : '');
    case 'filterLookupOptions':
      return `Lookup ${rule.field} depende de ${rule.parentField}`;
    case 'setComputed':
      return `Calculado ${rule.field}`;
    default:
      return rule.action;
  }
}

function describeWhenPt(when: TFormConditionNode | undefined): string {
  const leaf = whenNodeToUi(when);
  if (!leaf) return 'sempre';
  const op = conditionOpLabel(leaf.op);
  const needsVal =
    leaf.op !== 'isEmpty' &&
    leaf.op !== 'isFilled' &&
    leaf.op !== 'isTrue' &&
    leaf.op !== 'isFalse';
  const val = needsVal ? ` "${leaf.compareValue}"` : '';
  return `${leaf.field} ${op}${val}`;
}

export function describeRule(rule: TFormRule): string {
  const whenPt = describeWhenPt(rule.when);
  switch (rule.action) {
    case 'setVisibility':
      return `Se ${whenPt}: ${rule.visibility === 'show' ? 'mostrar' : 'ocultar'} ${rule.targetKind} "${rule.targetId}".`;
    case 'setRequired':
      return `Se ${whenPt}: ${rule.required ? 'tornar obrigatório' : 'tornar opcional'} o campo "${rule.field}".`;
    case 'setDisabled':
      return `Se ${whenPt}: ${rule.disabled ? 'desativar' : 'ativar'} o campo "${rule.field}".`;
    case 'setReadOnly':
      return `Se ${whenPt}: ${rule.readOnly ? 'somente leitura' : 'editável'} o campo "${rule.field}".`;
    case 'setDefault':
      return `Se ${whenPt}: valor padrão em "${rule.field}" = "${rule.value}".`;
    case 'validateValue':
      return `Se ${whenPt}: validar valor do campo "${rule.field}".`;
    case 'validateDate':
      return `Se ${whenPt}: validar data do campo "${rule.field}".`;
    case 'showMessage':
      return `Se ${whenPt}: mensagem (${rule.variant}) — ${rule.text}`;
    case 'filterLookupOptions':
      return `Se ${whenPt}: filtrar opções de lookup "${rule.field}" pelo campo "${rule.parentField}".`;
    case 'setComputed':
      return `Se ${whenPt}: calcular "${rule.field}" com expressão.`;
    case 'clearFields':
      return `Ao mudar "${rule.triggerField ?? '?'}": limpar ${rule.fields.join(', ')}.`;
    case 'attachmentRules':
      return 'Regras de anexo no formulário.';
    default:
      return describeRuleShort(rule);
  }
}

export function describeConditionalCardPT(card: IConditionalRuleCard): string {
  const w = card.when;
  const op = conditionOpLabel(w.op);
  const val =
    w.op === 'isEmpty' || w.op === 'isFilled' || w.op === 'isTrue' || w.op === 'isFalse'
      ? ''
      : ` ${w.compareValue}`;
  const g =
    card.groupTitles && card.groupTitles.length
      ? ` · grupos: ${card.groupTitles.join(', ')}`
      : '';
  return `Quando ${w.field} ${op}${val} → ${card.effects.length} efeito(s)${g}`;
}
