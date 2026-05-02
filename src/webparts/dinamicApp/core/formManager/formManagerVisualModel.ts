import type { FieldMappedType } from '../../../../services/shared/types';
import {
  FORM_VISIBILITY_PREFER_HIDE_TAG,
  type TFormConditionNode,
  type TFormConditionOp,
  type TFormManagerFormMode,
  type TFormRule,
  type IFormCompareRef,
  type IFormRuleAttachment,
  type IFormFieldConfig,
  type ITextFieldConditionalCondition,
  type ITextFieldConditionalGroup,
  type ITextFieldConditionalVisibility,
  type TLookupFilterOperator,
  type TTextFieldConditionalAction,
} from '../config/types/formManager';

const ALL_TEXT_COND_MODES: TFormManagerFormMode[] = ['create', 'edit', 'view'];

function activeModesForTextConditionalGroup(g: ITextFieldConditionalGroup): TFormManagerFormMode[] {
  return g.modes?.length ? g.modes : ALL_TEXT_COND_MODES;
}

function effectiveTextConditionalAction(g: ITextFieldConditionalGroup, mode: TFormManagerFormMode): TTextFieldConditionalAction {
  const per = g.actionByMode?.[mode];
  return per ?? g.action;
}

function clusterTextConditionalModesByAction(
  g: ITextFieldConditionalGroup
): { modes: TFormManagerFormMode[]; action: TTextFieldConditionalAction }[] {
  const active = activeModesForTextConditionalGroup(g);
  const map = new Map<TTextFieldConditionalAction, TFormManagerFormMode[]>();
  for (let i = 0; i < active.length; i++) {
    const m = active[i];
    const a = effectiveTextConditionalAction(g, m);
    if (!map.has(a)) map.set(a, []);
    map.get(a)!.push(m);
  }
  return Array.from(map.entries()).map(([action, modes]) => ({ action, modes }));
}

export function isSetComputedAllowedForMappedType(mt: FieldMappedType | 'unknown' | undefined): boolean {
  if (!mt || mt === 'unknown' || mt === 'calculated') return false;
  return true;
}

export const CONDITION_OP_OPTIONS: { key: TFormConditionOp; text: string }[] = [
  { key: 'eq', text: 'é igual a' },
  { key: 'ne', text: 'é diferente de' },
  { key: 'contains', text: 'contém' },
  { key: 'notContains', text: 'não contém' },
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

export type TCompareUiKind =
  | 'literal'
  | 'field'
  | 'token'
  | 'spGroupMember'
  | 'spGroupNotMember';

export interface IWhenUi {
  field: string;
  op: TFormConditionOp;
  compareKind: TCompareUiKind;
  compareValue: string;
}

function whenUiCompleteForSetDisabledWhen(ui: IWhenUi): boolean {
  if (ui.compareKind === 'spGroupMember' || ui.compareKind === 'spGroupNotMember') {
    return ui.compareValue.trim().length > 0;
  }
  return ui.field.trim().length > 0;
}

export function whenUiToNode(w: IWhenUi): TFormConditionNode {
  if (w.compareKind === 'spGroupMember' || w.compareKind === 'spGroupNotMember') {
    const t = w.compareValue.trim().slice(0, 256);
    return {
      kind: 'userGroup',
      invert: w.compareKind === 'spGroupNotMember',
      groupTitle: t,
    };
  }
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
  if (!node) return undefined;
  if (node.kind === 'userGroup') {
    return {
      field: '',
      op: 'eq',
      compareKind: node.invert ? 'spGroupNotMember' : 'spGroupMember',
      compareValue: node.groupTitle,
    };
  }
  if (node.kind !== 'leaf') return undefined;
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
  excludeGroupTitles?: string[];
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
      ...(card.excludeGroupTitles?.length ? { excludeGroupTitles: card.excludeGroupTitles } : {}),
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
    const excludeGroupTitles = first.excludeGroupTitles;
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
      ...(excludeGroupTitles?.length ? { excludeGroupTitles } : {}),
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
    blockedWeekdays: number[];
    gteField: string;
    lteField: string;
    message: string;
  };
  filterLookup: {
    parentField: string;
    childField: string;
    filterOperator: TLookupFilterOperator | '';
  };
  computedExpression: string;
  /** Id do nó na árvore de pastas (Anexos); gera expressão `attfolder:id`. */
  computedAttachmentFolderNodeId: string;
  /** Sempre substituir por expressão em edição (valor gravado ignorado). */
  computedLiveInEditView: boolean;
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
      blockedWeekdays: [],
      gteField: '',
      lteField: '',
      message: '',
    },
    filterLookup: { parentField: '', childField: '', filterOperator: '' },
    computedExpression: '',
    computedAttachmentFolderNodeId: '',
    computedLiveInEditView: false,
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
      if (typeof r.minDaysFromTodayExpr === 'string' && r.minDaysFromTodayExpr.trim()) {
        st.validateDate.minDaysFromToday = r.minDaysFromTodayExpr.trim();
      } else if (r.minDaysFromToday !== undefined) {
        st.validateDate.minDaysFromToday = String(r.minDaysFromToday);
      }
      if (typeof r.maxDaysFromTodayExpr === 'string' && r.maxDaysFromTodayExpr.trim()) {
        st.validateDate.maxDaysFromToday = r.maxDaysFromTodayExpr.trim();
      } else if (r.maxDaysFromToday !== undefined) {
        st.validateDate.maxDaysFromToday = String(r.maxDaysFromToday);
      }
      if (r.blockWeekends) st.validateDate.blockWeekends = true;
      if (r.blockedWeekdays?.length) {
        const set = new Set(st.validateDate.blockedWeekdays);
        for (let bi = 0; bi < r.blockedWeekdays.length; bi++) {
          const x = r.blockedWeekdays[bi];
          if (typeof x === 'number' && x >= 0 && x <= 6 && x === Math.floor(x)) set.add(x);
        }
        st.validateDate.blockedWeekdays = Array.from(set).sort((a, b) => a - b);
      }
      if (r.gteField) st.validateDate.gteField = r.gteField;
      if (r.lteField) st.validateDate.lteField = r.lteField;
      if (r.message) st.validateDate.message = r.message;
    }
    if (r.action === 'filterLookupOptions' && r.field === internalName) {
      st.filterLookup.parentField = r.parentField;
      st.filterLookup.childField = r.childField ?? '';
      st.filterLookup.filterOperator = r.filterOperator ?? '';
    }
    if (r.action === 'setComputed' && r.field === internalName) {
      const ex = String(r.expression ?? '').trim();
      st.computedLiveInEditView = r.alwaysLiveComputed === true;
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

export type TDaysOffsetEditorParse =
  | { kind: 'empty' }
  | { kind: 'literal'; n: number }
  | { kind: 'expr'; expr: string };

export function parseDaysOffsetEditorString(raw: string): TDaysOffsetEditorParse {
  const t = raw.trim();
  if (!t) return { kind: 'empty' };
  if (/^-?\d+$/.test(t)) return { kind: 'literal', n: Number(t) };
  return { kind: 'expr', expr: t };
}

function textConditionalConditionToWhenUi(c: ITextFieldConditionalCondition): IWhenUi {
  return {
    field: c.refField.trim(),
    op: c.op as TFormConditionOp,
    compareKind: c.compareKind as TCompareUiKind,
    compareValue: c.compareValue,
  };
}

function buildWhenFromTextConditionalGroup(group: ITextFieldConditionalGroup): TFormConditionNode | undefined {
  const leaves: TFormConditionNode[] = [];
  for (let i = 0; i < group.conditions.length; i++) {
    const c = group.conditions[i];
    if (!c.refField.trim()) continue;
    leaves.push(whenUiToNode(textConditionalConditionToWhenUi(c)));
  }
  if (!leaves.length) return undefined;
  if (leaves.length === 1) return leaves[0];
  return group.groupOp === 'any' ? { kind: 'any', children: leaves } : { kind: 'all', children: leaves };
}

export function compileTextFieldConditionalVisibilityRules(
  internalName: string,
  vis: ITextFieldConditionalVisibility | undefined
): TFormRule[] {
  if (!vis?.groups?.length) return [];
  const seg = safeIdSegment(internalName);
  const out: TFormRule[] = [];
  for (let i = 0; i < vis.groups.length; i++) {
    const g = vis.groups[i];
    const when = buildWhenFromTextConditionalGroup(g);
    if (!when) continue;
    const gid = safeIdSegment(g.id || `g${i}`);
    const groupPayload = g.groupTitles?.length ? { groupTitles: g.groupTitles } : {};
    const excludePayload = g.excludeGroupTitles?.length ? { excludeGroupTitles: g.excludeGroupTitles } : {};
    const clusters = clusterTextConditionalModesByAction(g);
    for (let c = 0; c < clusters.length; c++) {
      const cl = clusters[c];
      const { modes, action } = cl;
      if (!modes.length) continue;
      const modePayload = modes.length === ALL_TEXT_COND_MODES.length ? {} : { modes };
      const suf = clusters.length > 1 ? `_${c + 1}` : '';
      if (action === 'disable') {
        out.push({
          ...modePayload,
          ...groupPayload,
          ...excludePayload,
          id: `ui_f_${seg}_txdis_${gid}${suf}`,
          action: 'setDisabled',
          field: internalName,
          disabled: true,
          when,
        });
      } else {
        out.push({
          ...modePayload,
          ...groupPayload,
          ...excludePayload,
          id: `ui_f_${seg}_txvis_${gid}${suf}`,
          action: 'setVisibility',
          targetKind: 'field',
          targetId: internalName,
          visibility: action === 'hide' ? 'hide' : 'show',
          when,
          tags: [FORM_VISIBILITY_PREFER_HIDE_TAG],
        });
      }
    }
  }
  return out;
}

export interface IBuildFieldUiRulesOptions {
  mappedType?: FieldMappedType | 'unknown';
}

export function buildFieldUiRules(
  internalName: string,
  st: IFieldRuleEditorState,
  fieldConfig?: Pick<IFormFieldConfig, 'textConditionalVisibility'>,
  opts?: IBuildFieldUiRulesOptions
): TFormRule[] {
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
    (vd.blockedWeekdays?.length ?? 0) > 0 ||
    vd.gteField.trim() ||
    vd.lteField.trim();
  if (hasDate) {
    const bw =
      vd.blockedWeekdays?.length && vd.blockedWeekdays.every((x) => x >= 0 && x <= 6)
        ? [...vd.blockedWeekdays].sort((a, b) => a - b)
        : [];
    const minOff = parseDaysOffsetEditorString(vd.minDaysFromToday);
    const maxOff = parseDaysOffsetEditorString(vd.maxDaysFromToday);
    out.push({
      id: id('date'),
      action: 'validateDate',
      field: internalName,
      ...(minOff.kind === 'literal' ? { minDaysFromToday: minOff.n } : {}),
      ...(minOff.kind === 'expr' ? { minDaysFromTodayExpr: minOff.expr } : {}),
      ...(maxOff.kind === 'literal' ? { maxDaysFromToday: maxOff.n } : {}),
      ...(maxOff.kind === 'expr' ? { maxDaysFromTodayExpr: maxOff.expr } : {}),
      ...(vd.blockWeekends ? { blockWeekends: true } : {}),
      ...(bw.length ? { blockedWeekdays: bw } : {}),
      ...(vd.gteField.trim() ? { gteField: vd.gteField.trim() } : {}),
      ...(vd.lteField.trim() ? { lteField: vd.lteField.trim() } : {}),
      ...(vd.message.trim() ? { message: vd.message.trim() } : {}),
      ...baseModes,
    });
  }

  const fl = st.filterLookup;
  if (fl.parentField.trim() && fl.childField.trim() && fl.filterOperator) {
    out.push({
      id: id('lk'),
      action: 'filterLookupOptions',
      field: internalName,
      parentField: fl.parentField.trim(),
      childField: fl.childField.trim(),
      filterOperator: fl.filterOperator as TLookupFilterOperator,
      ...baseModes,
    });
  }

  const attFolderId = st.computedAttachmentFolderNodeId.trim();
  const cmpExpr = attFolderId ? `attfolder:${attFolderId}` : st.computedExpression.trim();
  if (cmpExpr && isSetComputedAllowedForMappedType(opts?.mappedType)) {
    out.push({
      id: id('cmp'),
      action: 'setComputed',
      field: internalName,
      expression: cmpExpr,
      ...(st.computedLiveInEditView ? { alwaysLiveComputed: true as const } : {}),
    });
  }

  if (st.disableWhenActive && whenUiCompleteForSetDisabledWhen(st.disableWhenUi)) {
    out.push({
      id: id('discond'),
      action: 'setDisabled',
      field: internalName,
      disabled: true,
      when: whenUiToNode(st.disableWhenUi),
      ...baseModes,
    });
  }
  if (st.enableWhenActive && whenUiCompleteForSetDisabledWhen(st.enableWhenUi)) {
    out.push({
      id: id('enacond'),
      action: 'setDisabled',
      field: internalName,
      disabled: false,
      when: whenUiToNode(st.enableWhenUi),
      ...baseModes,
    });
  }

  out.push(...compileTextFieldConditionalVisibilityRules(internalName, fieldConfig?.textConditionalVisibility));

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
      blockedWeekdays: [],
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
  if (leaf.compareKind === 'spGroupMember') {
    return `utilizador no grupo "${leaf.compareValue}"`;
  }
  if (leaf.compareKind === 'spGroupNotMember') {
    return `utilizador fora do grupo "${leaf.compareValue}"`;
  }
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
  const gx =
    card.excludeGroupTitles && card.excludeGroupTitles.length
      ? ` · excluir: ${card.excludeGroupTitles.join(', ')}`
      : '';
  return `Quando ${w.field} ${op}${val} → ${card.effects.length} efeito(s)${g}${gx}`;
}
