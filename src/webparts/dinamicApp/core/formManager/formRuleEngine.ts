import type { IDynamicContext } from '../dynamicTokens/types';
import { DynamicTokenResolver } from '../dynamicTokens/services/DynamicTokenResolver';
import { isDynamicToken, resolveStringToken, toIsoDateString } from '../dynamicTokens';
import type {
  IAttachmentLibraryFolderTreeNode,
  IFormManagerConfig,
  IFormFieldConfig,
  IFormCustomButtonConfig,
  IFormLinkedChildFormConfig,
  TFormConditionNode,
  TFormRule,
  TFormManagerFormMode,
  TFormSubmitKind,
  IFormCompareRef,
  TFormConditionOp,
  TFormCustomButtonOperation,
} from '../config/types/formManager';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_BANNER_INTERNAL_PREFIX,
  FORM_VISIBILITY_PREFER_HIDE_TAG,
  isFormBannerFieldConfig,
} from '../config/types/formManager';
import type { IFieldMetadata } from '../../../../services';
import { buildAttachmentFolderAbsoluteUrl } from './formAttachmentLibrary';
import { applyFormFieldTextTransform } from './formTextValueTransform';
import { ensureAbsoluteSharePointUrl } from './formUrlUtils';

const FULL_SUBMIT_TAG = 'fullSubmitOnly';

const ATT_FOLDER_EXPR_PREFIX = 'attfolder:';

export interface IFormAttachmentFolderUrlContext {
  libraryRootServerRelativeUrl?: string;
  itemId?: number;
  folderTree?: IAttachmentLibraryFolderTreeNode[];
}

export interface IFormRuleRuntimeContext {
  formMode: TFormManagerFormMode;
  values: Record<string, unknown>;
  submitKind?: TFormSubmitKind;
  userGroupTitles: string[];
  currentUserId: number;
  authorId?: number;
  dynamicContext: IDynamicContext;
  /** Resolução de expressões `attfolder:id` (pasta da árvore em Anexos → biblioteca). */
  attachmentFolderUrl?: IFormAttachmentFolderUrlContext;
}

export function withRuleRuntimeDynamicContext(ctx: IDynamicContext, currentUserId: number): IDynamicContext {
  if (!currentUserId || ctx.currentUser?.id !== undefined) return ctx;
  return { ...ctx, currentUser: { ...(ctx.currentUser ?? {}), id: currentUserId } };
}

export interface IFormDerivedUiState {
  fieldVisible: Record<string, boolean>;
  sectionVisible: Record<string, boolean>;
  fieldRequired: Record<string, boolean>;
  fieldDisabled: Record<string, boolean>;
  fieldReadOnly: Record<string, boolean>;
  effectiveSectionByField: Record<string, string>;
  messages: { variant: 'info' | 'warning' | 'error'; text: string; ruleId: string }[];
  lookupFilters: Record<string, { parentField: string; childField?: string; filterOperator?: string; odataFilterTemplate?: string }>;
  computedDisplay: Record<string, unknown>;
  dynamicHelpByField: Record<string, string>;
}

const tokenResolver = new DynamicTokenResolver();

const DATE_DEFAULT_COMPOUND_RE =
  /^((?:\[[^\]]+\]|\{\{[^}]+\}\}))\s*([+-])\s*(\d+)(?:\s+(dias?|days?|semanas?|weeks?))?$/i;

const DATE_DEFAULT_SINGLE_PLACEHOLDER_RE = /^\{\{([^}]+)\}\}$/;
const DEFAULT_PLACEHOLDER_RE = /\{\{([^}]+)\}\}/g;
const DEFAULT_INLINE_TOKEN_RE = /\[(.+?)\]/g;

function normalizeDateDefaultExpression(expr: string): string {
  let t = expr.trim().replace(/\s+/g, ' ');
  if (!t) return '';
  t = t.replace(/\bhoje\b/gi, '[today]');
  t = t.replace(/\bontem\b/gi, '[yesterday]');
  t = t.replace(/\bamanhã\b/gi, '[tomorrow]');
  t = t.replace(/\bamanha\b/gi, '[tomorrow]');
  t = t.replace(/\btoday\b/gi, '[today]');
  t = t.replace(/\btomorrow\b/gi, '[tomorrow]');
  t = t.replace(/\byesterday\b/gi, '[yesterday]');
  t = t.replace(/\bagora\b/gi, '[now]');
  return t;
}

export type TResolveDateFieldDefaultResult =
  | { kind: 'resolved'; value: unknown }
  | { kind: 'generic' }
  | { kind: 'skip' };

export function resolveDateFieldDefaultValue(
  expr: string,
  dynamicContext: IDynamicContext,
  values?: Record<string, unknown>
): TResolveDateFieldDefaultResult {
  const normalized = normalizeDateDefaultExpression(expr);
  if (!normalized) return { kind: 'generic' };

  if (isDynamicToken(normalized)) {
    return { kind: 'resolved', value: tokenResolver.resolveStringToken(normalized, dynamicContext) };
  }

  const normStr = String(normalized).trim();
  const singlePh = DATE_DEFAULT_SINGLE_PLACEHOLDER_RE.exec(normStr);
  if (singlePh && values) {
    const name = singlePh[1].trim();
    const raw = values[name];
    if (raw === undefined || raw === null) return { kind: 'skip' };
    const baseStr = typeof raw === 'string' ? raw : String(raw);
    const d = parseFormCalendarDateString(baseStr) ?? parseIsoDate(baseStr);
    if (!d) return { kind: 'skip' };
    return { kind: 'resolved', value: toIsoDateString(startOfDay(d)) };
  }

  const m = normStr.match(DATE_DEFAULT_COMPOUND_RE);
  if (!m) return { kind: 'generic' };

  const baseTok = m[1];
  const sign = m[2] === '-' ? -1 : 1;
  let n = parseInt(m[3], 10);
  const unit = (m[4] ?? '').toLowerCase();
  if (unit.startsWith('sem') || unit.startsWith('week')) n *= 7;

  let baseResolved: unknown;
  if (baseTok.indexOf('{{') === 0 && baseTok.indexOf('}}') === baseTok.length - 2) {
    const inner = baseTok.slice(2, -2).trim();
    const raw = values?.[inner];
    if (raw === undefined || raw === null) return { kind: 'skip' };
    baseResolved = typeof raw === 'string' ? raw : String(raw);
  } else {
    baseResolved = tokenResolver.resolveStringToken(baseTok, dynamicContext);
  }
  if (baseResolved === undefined || baseResolved === null) return { kind: 'skip' };

  const baseStr = typeof baseResolved === 'string' ? baseResolved : String(baseResolved);
  const d = parseFormCalendarDateString(baseStr) ?? parseIsoDate(baseStr);
  if (!d) return { kind: 'skip' };

  const s = startOfDay(d);
  s.setDate(s.getDate() + sign * n);
  return { kind: 'resolved', value: toIsoDateString(s) };
}

export interface IGetDefaultValuesFromRulesOptions {
  isDateTimeField?: (internalName: string) => boolean;
}

function normGroupTitle(s: string): string {
  return s.trim().toLowerCase();
}

export function userInAnyGroup(userTitles: string[], ruleGroups: string[] | undefined): boolean {
  if (!ruleGroups || ruleGroups.length === 0) return true;
  const set = new Set(userTitles.map(normGroupTitle));
  for (let i = 0; i < ruleGroups.length; i++) {
    if (set.has(normGroupTitle(ruleGroups[i]))) return true;
  }
  return false;
}

export function isAttachmentFolderUploaderVisible(
  node: IAttachmentLibraryFolderTreeNode,
  ctx: IFormRuleRuntimeContext
): boolean {
  const modes = node.showUploaderModes;
  if (modes !== undefined) {
    if (!modes.length || modes.indexOf(ctx.formMode) === -1) return false;
  }
  if (!userInAnyGroup(ctx.userGroupTitles, node.showUploaderGroupTitles)) return false;
  return evaluateCondition(node.showUploaderWhen, ctx.values, ctx.dynamicContext);
}

function isEmptyish(v: unknown): boolean {
  if (v === null || v === undefined) return true;
  if (typeof v === 'string' && v.trim() === '') return true;
  if (Array.isArray(v) && v.length === 0) return true;
  if (typeof v === 'object' && v !== null && 'Id' in (v as object)) {
    const id = (v as Record<string, unknown>).Id;
    if (id === null || id === undefined || id === '') return true;
  }
  return false;
}

function tryGetObjectProp(obj: Record<string, unknown>, key: string): unknown {
  if (key in obj) return obj[key];
  const low = key.toLowerCase();
  const keys = Object.keys(obj);
  for (let i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === low) return obj[keys[i]];
  }
  return undefined;
}

function readPathValue(root: unknown, path: string[]): unknown {
  if (!path.length) return root;
  if (root === null || root === undefined) return undefined;
  if (Array.isArray(root)) {
    const mapped = root
      .map((item) => readPathValue(item, path))
      .filter((x) => x !== undefined && x !== null && String(x).trim() !== '');
    return mapped.length ? mapped.join('; ') : undefined;
  }
  if (typeof root === 'object') {
    const obj = root as Record<string, unknown>;
    const head = path[0];
    const next = tryGetObjectProp(obj, head);
    return readPathValue(next, path.slice(1));
  }
  return undefined;
}

function toScalarString(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') {
    const s = v;
    if (s.startsWith('/') && !s.startsWith('//')) return ensureAbsoluteSharePointUrl(s);
    return s;
  }
  if (typeof v === 'number') return isFinite(v) ? String(v) : '';
  if (typeof v === 'boolean') return v ? 'true' : 'false';
  if (Array.isArray(v)) return v.map((x) => toScalarString(x)).filter(Boolean).join('; ');
  if (typeof v === 'object') {
    const o = v as Record<string, unknown>;
    const urlRaw = tryGetObjectProp(o, 'Url');
    if (typeof urlRaw === 'string' && urlRaw.trim()) return ensureAbsoluteSharePointUrl(urlRaw);
    const title = tryGetObjectProp(o, 'Title');
    if (title !== undefined && title !== null) return String(title);
    const id = tryGetObjectProp(o, 'Id');
    if (id !== undefined && id !== null) return String(id);
  }
  return String(v);
}

function resolveDefaultTemplateValue(
  raw: string,
  values: Record<string, unknown>,
  dynamicContext: IDynamicContext
): { usedTemplate: boolean; value: unknown } {
  const trimmed = raw.trim();
  const hasPlaceholder = DEFAULT_PLACEHOLDER_RE.test(trimmed);
  DEFAULT_PLACEHOLDER_RE.lastIndex = 0;
  const hasInlineToken = DEFAULT_INLINE_TOKEN_RE.test(trimmed);
  DEFAULT_INLINE_TOKEN_RE.lastIndex = 0;
  if (!hasPlaceholder && !hasInlineToken) return { usedTemplate: false, value: raw };

  let replaced = trimmed.replace(DEFAULT_PLACEHOLDER_RE, (_full, innerRaw: string) => {
    const inner = String(innerRaw).trim();
    if (!inner) return '';
    const path = inner.split('/').map((p) => p.trim()).filter(Boolean);
    if (!path.length) return '';
    const rootName = path[0];
    const rootVal = values[rootName];
    const resolved = path.length > 1 ? readPathValue(rootVal, path.slice(1)) : rootVal;
    return toScalarString(resolved);
  });

  replaced = replaced.replace(DEFAULT_INLINE_TOKEN_RE, (full, innerRaw: string) => {
    const tok = `[${String(innerRaw).trim()}]`;
    if (!isDynamicToken(tok)) return full;
    const resolved = tokenResolver.resolveStringToken(tok, dynamicContext);
    return resolved === undefined || resolved === null ? '' : toScalarString(resolved);
  });

  const single = DATE_DEFAULT_SINGLE_PLACEHOLDER_RE.exec(trimmed);
  if (single) {
    const inner = single[1].trim();
    const path = inner.split('/').map((p) => p.trim()).filter(Boolean);
    if (path.length) {
      const rootName = path[0];
      const rootVal = values[rootName];
      const resolved = path.length > 1 ? readPathValue(rootVal, path.slice(1)) : rootVal;
      if (resolved !== undefined && resolved !== null) return { usedTemplate: true, value: resolved };
    }
  }

  return { usedTemplate: true, value: replaced };
}

function coerceBool(v: unknown): boolean | undefined {
  if (v === true || v === false) return v;
  if (v === 1 || v === '1' || v === 'true' || v === 'True') return true;
  if (v === 0 || v === '0' || v === 'false' || v === 'False') return false;
  return undefined;
}

function coerceNumber(v: unknown): number {
  if (typeof v === 'number') return isFinite(v) ? v : NaN;
  if (typeof v === 'bigint') return Number(v);
  if (v === null || v === undefined) return NaN;
  if (typeof v === 'boolean') return NaN;
  if (typeof v === 'object' && v !== null) {
    const ro = v as Record<string, unknown>;
    const results = ro.results;
    if (Array.isArray(results) && results.length === 1) return coerceNumber(results[0]);
    const wrapped = ro.Value ?? ro.value;
    if (wrapped !== undefined && wrapped !== v) return coerceNumber(wrapped);
  }
  if (typeof v === 'string') {
    let s = v.replace(/[\s\u00a0\u202f]/g, '').trim();
    if (!s) return NaN;
    const hasComma = s.indexOf(',') !== -1;
    const hasDot = s.indexOf('.') !== -1;
    if (hasComma && hasDot) {
      const lastComma = s.lastIndexOf(',');
      const lastDot = s.lastIndexOf('.');
      if (lastComma > lastDot) {
        s = s.replace(/\./g, '').replace(',', '.');
      } else {
        s = s.replace(/,/g, '');
      }
    } else if (hasComma) {
      s = s.replace(',', '.');
    }
    const n = Number(s);
    return isNaN(n) ? NaN : n;
  }
  return NaN;
}

function normalizeForEqNe(v: unknown): unknown {
  if (v !== null && typeof v === 'object' && 'Id' in (v as object)) {
    const id = (v as Record<string, unknown>).Id;
    if (typeof id === 'number' && isFinite(id)) return id;
  }
  if (typeof v === 'string') return v.trim();
  return v;
}

function compareResolved(left: unknown, op: string, right: unknown): boolean {
  switch (op) {
    case 'isEmpty':
      return isEmptyish(left);
    case 'isFilled':
      return !isEmptyish(left);
    case 'isTrue':
      return coerceBool(left) === true;
    case 'isFalse':
      return coerceBool(left) === false;
    case 'eq': {
      const L = normalizeForEqNe(left);
      const R = normalizeForEqNe(right);
      if (typeof L === 'number' && typeof R === 'number') return L === R;
      return String(L ?? '') === String(R ?? '');
    }
    case 'ne': {
      const L = normalizeForEqNe(left);
      const R = normalizeForEqNe(right);
      if (typeof L === 'number' && typeof R === 'number') return L !== R;
      return String(L ?? '') !== String(R ?? '');
    }
    case 'contains':
      return String(left ?? '').toLowerCase().indexOf(String(right ?? '').toLowerCase()) !== -1;
    case 'notContains':
      return String(left ?? '').toLowerCase().indexOf(String(right ?? '').toLowerCase()) === -1;
    case 'startsWith':
      return String(left ?? '').toLowerCase().indexOf(String(right ?? '').toLowerCase()) === 0;
    case 'endsWith': {
      const a = String(left ?? '').toLowerCase();
      const b = String(right ?? '').toLowerCase();
      return b.length === 0 || (a.length >= b.length && a.slice(a.length - b.length) === b);
    }
    case 'gt':
      return coerceNumber(left) > coerceNumber(right);
    case 'ge':
      return coerceNumber(left) >= coerceNumber(right);
    case 'lt':
      return coerceNumber(left) < coerceNumber(right);
    case 'le':
      return coerceNumber(left) <= coerceNumber(right);
    default:
      return false;
  }
}

function resolveCompare(
  ref: IFormCompareRef | undefined,
  values: Record<string, unknown>,
  ctx: IDynamicContext
): unknown {
  if (!ref) return undefined;
  if (ref.kind === 'literal') return ref.value;
  if (ref.kind === 'field') return values[ref.value];
  const tok = ref.value.indexOf('[') === 0 ? ref.value : `[${ref.value}]`;
  return tokenResolver.resolveStringToken(tok, ctx);
}

export function evaluateCondition(
  node: TFormConditionNode | undefined,
  values: Record<string, unknown>,
  dynamicContext: IDynamicContext
): boolean {
  if (!node) return true;
  if (node.kind === 'all') {
    for (let i = 0; i < node.children.length; i++) {
      if (!evaluateCondition(node.children[i], values, dynamicContext)) return false;
    }
    return true;
  }
  if (node.kind === 'any') {
    for (let i = 0; i < node.children.length; i++) {
      if (evaluateCondition(node.children[i], values, dynamicContext)) return true;
    }
    return false;
  }
  if (node.kind === 'leaf') {
    const left = values[node.field];
    const right = resolveCompare(node.compare, values, dynamicContext);
    return compareResolved(left, node.op, right);
  }
  const legacy = node as { field?: string; op?: TFormConditionOp; compare?: IFormCompareRef };
  if (typeof legacy.field === 'string' && legacy.field.trim() && legacy.op) {
    const left = values[legacy.field];
    const right = resolveCompare(legacy.compare, values, dynamicContext);
    return compareResolved(left, legacy.op, right);
  }
  return false;
}

export function findEnabledSetComputedRule(
  rules: TFormRule[] | undefined,
  fieldName: string
): Extract<TFormRule, { action: 'setComputed' }> | undefined {
  if (!rules?.length) return undefined;
  for (let i = 0; i < rules.length; i++) {
    const r = rules[i];
    if (r.enabled === false) continue;
    if (r.action === 'setComputed' && r.field === fieldName) return r as Extract<TFormRule, { action: 'setComputed' }>;
  }
  return undefined;
}

export function resolveSetComputedDisplayValue(args: {
  derivedComputed: unknown;
  formMode: TFormManagerFormMode;
  itemId: number | undefined;
  fieldName: string;
  expressionSnapAtItemOpenByField: Readonly<Record<string, string>>;
  setComputedRule: Extract<TFormRule, { action: 'setComputed' }> | undefined;
}): unknown {
  const { derivedComputed, formMode, itemId, fieldName, expressionSnapAtItemOpenByField, setComputedRule } = args;
  if (derivedComputed === undefined) return undefined;
  if (setComputedRule?.alwaysLiveComputed === true) return derivedComputed;
  if (formMode === 'create') return derivedComputed;
  if (itemId === undefined || itemId === null || typeof itemId !== 'number' || !isFinite(itemId)) {
    return derivedComputed;
  }
  const snap = expressionSnapAtItemOpenByField[fieldName];
  const cur = (setComputedRule?.expression ?? '').trim();
  if (snap === undefined) return derivedComputed;
  if (cur !== snap) return derivedComputed;
  return undefined;
}

/**
 * Junta limites de caracteres de regras `validateValue` do campo que passam em modo, when e grupos
 * (alinhado à validação em submissão). Vários limites ativos: mínimo efetivo = maior dos mínimos, máximo = menor dos máximos.
 */
export function getMergedValidateValueLengthBounds(
  rules: TFormRule[] | undefined,
  fieldName: string,
  ctx: Pick<IFormRuleRuntimeContext, 'formMode' | 'values' | 'userGroupTitles' | 'dynamicContext'>,
  fieldVisibleMap: Record<string, boolean> | undefined
): { minLength?: number; maxLength?: number } | undefined {
  if (!rules?.length) return undefined;
  const { formMode, values, userGroupTitles, dynamicContext } = ctx;
  let minL: number | undefined;
  let maxL: number | undefined;
  for (let r = 0; r < rules.length; r++) {
    const rule = rules[r];
    if (rule.action !== 'validateValue') continue;
    if (rule.field !== fieldName) continue;
    if (rule.enabled === false) continue;
    if (!ruleAppliesMode(rule, formMode)) continue;
    if (!userInAnyGroup(userGroupTitles, rule.groupTitles)) continue;
    if (!evaluateCondition(rule.when, values, dynamicContext)) continue;
    if (fieldVisibleMap && fieldVisibleMap[fieldName] === false) continue;
    if (rule.minLength !== undefined) {
      minL = minL === undefined ? rule.minLength : Math.max(minL, rule.minLength);
    }
    if (rule.maxLength !== undefined) {
      maxL = maxL === undefined ? rule.maxLength : Math.min(maxL, rule.maxLength);
    }
  }
  if (minL === undefined && maxL === undefined) return undefined;
  return { minLength: minL, maxLength: maxL };
}

/**
 * Junta min/max numéricos de regras validateValue ativas para o campo (mesmo critério que em submissão).
 * Vários mínimos → maior; vários máximos → menor.
 */
export function getMergedValidateValueNumberBounds(
  rules: TFormRule[] | undefined,
  fieldName: string,
  ctx: Pick<IFormRuleRuntimeContext, 'formMode' | 'values' | 'userGroupTitles' | 'dynamicContext'>,
  fieldVisibleMap: Record<string, boolean> | undefined
): { minNumber?: number; maxNumber?: number } | undefined {
  if (!rules?.length) return undefined;
  const { formMode, values, userGroupTitles, dynamicContext } = ctx;
  let minN: number | undefined;
  let maxN: number | undefined;
  for (let r = 0; r < rules.length; r++) {
    const rule = rules[r];
    if (rule.action !== 'validateValue') continue;
    if (rule.field !== fieldName) continue;
    if (rule.enabled === false) continue;
    if (!ruleAppliesMode(rule, formMode)) continue;
    if (!userInAnyGroup(userGroupTitles, rule.groupTitles)) continue;
    if (!evaluateCondition(rule.when, values, dynamicContext)) continue;
    if (fieldVisibleMap && fieldVisibleMap[fieldName] === false) continue;
    if (rule.minNumber !== undefined) {
      minN = minN === undefined ? rule.minNumber : Math.max(minN, rule.minNumber);
    }
    if (rule.maxNumber !== undefined) {
      maxN = maxN === undefined ? rule.maxNumber : Math.min(maxN, rule.maxNumber);
    }
  }
  if (minN === undefined && maxN === undefined) return undefined;
  return { minNumber: minN, maxNumber: maxN };
}

function isValueEmptyForRequiredField(v: unknown, mappedType: string): boolean {
  if (mappedType === 'boolean') {
    return v === undefined || v === null;
  }
  if (mappedType === 'url') {
    if (v === null || v === undefined) return true;
    if (typeof v === 'object' && v !== null && 'Url' in (v as object)) {
      return String((v as Record<string, unknown>).Url ?? '').trim() === '';
    }
    if (typeof v === 'string') return v.trim() === '';
    return false;
  }
  if (v === null || v === undefined) return true;
  if (typeof v === 'string' && v.trim() === '') return true;
  if (Array.isArray(v) && v.length === 0) return true;
  if (typeof v === 'object' && v !== null && 'Id' in (v as object)) {
    const id = (v as Record<string, unknown>).Id;
    if (id === null || id === undefined || id === '') return true;
  }
  return false;
}

export function areAllRequiredFieldsFilled(
  cfg: IFormManagerConfig,
  fieldConfigs: IFormFieldConfig[],
  ctx: IFormRuleRuntimeContext,
  metaByName: Map<string, IFieldMetadata>,
  buttonOverlay?: IFormButtonVisibilityOverlay,
  attachmentCtx?: IFormValidationAttachmentContext
): boolean {
  const ctxSubmit: IFormRuleRuntimeContext = { ...ctx, submitKind: 'submit' };
  const derived = buildFormDerivedState(cfg, fieldConfigs, ctxSubmit, buttonOverlay, metaByName);
  const fv = (n: string): boolean => derived.fieldVisible[n] !== false;
  const { values, formMode } = ctx;

  for (let i = 0; i < fieldConfigs.length; i++) {
    const fc = fieldConfigs[i];
    const name = fc.internalName;
    if (!fv(name)) continue;

    if (isFormBannerFieldConfig(fc)) continue;

    if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) {
      const attReq = derived.fieldRequired[name] === true;
      if (!attReq) continue;
      const readOnly = formMode === 'view' || derived.fieldDisabled[name] === true;
      const pending = attachmentCtx?.pendingFiles?.length ?? 0;
      const existing = attachmentCtx?.attachmentCount ?? 0;
      const attSatisfied = pending > 0 || (formMode !== 'create' && existing > 0);
      const attReqEmpty = attReq && !readOnly && !attSatisfied;
      if (attReqEmpty) return false;
      continue;
    }

    const m = metaByName.get(name);
    const mappedType = m?.MappedType ?? 'text';
    const isRequired = derived.fieldRequired[name] === true || m?.Required === true;
    if (!isRequired) continue;
    const readOnly =
      formMode === 'view' || derived.fieldReadOnly[name] === true || derived.fieldDisabled[name] === true;
    if (readOnly) continue;
    if (isValueEmptyForRequiredField(values[name], mappedType)) return false;
  }
  return true;
}

/** Botão de histórico integrado (config na aba Componentes): só com item gravado. */
export function shouldShowBuiltinHistoryButton(visibilityOpts?: IFormCustomButtonVisibilityOpts): boolean {
  if (visibilityOpts?.historyEnabledInConfig !== true) return false;
  const hid = visibilityOpts?.historyItemId;
  if (hid === undefined || hid === null || typeof hid !== 'number' || !isFinite(hid)) return false;
  if (!userInAnyGroup(visibilityOpts.userGroupTitles ?? [], visibilityOpts.historyGroupTitles)) return false;
  return true;
}

export interface IFormCustomButtonVisibilityOpts {
  allRequiredFilled?: boolean;
  historyEnabledInConfig?: boolean;
  /** Item gravado na lista; ausente em modo novo sem id. */
  historyItemId?: number;
  /** Grupos SharePoint que podem ver o botão de histórico integrado. Vazio = todos. */
  historyGroupTitles?: string[];
  /** Títulos dos grupos do utilizador atual (runtime). */
  userGroupTitles?: string[];
}

export function shouldShowCustomButton(
  b: IFormCustomButtonConfig,
  ctx: IFormRuleRuntimeContext,
  visibilityOpts?: IFormCustomButtonVisibilityOpts
): boolean {
  if (b.enabled === false) return false;
  if (b.modes !== undefined && b.modes.length === 0) return false;
  if (b.modes?.length && b.modes.indexOf(ctx.formMode) === -1) return false;
  const op: TFormCustomButtonOperation = b.operation ?? 'legacy';
  if (op === 'history') {
    if (visibilityOpts?.historyEnabledInConfig !== true) return false;
    const hid = visibilityOpts?.historyItemId;
    if (hid === undefined || hid === null || typeof hid !== 'number' || !isFinite(hid)) return false;
  }
  if (op === 'delete') {
    if (ctx.formMode === 'create') return false;
    const sv = b.deleteShowInView !== false;
    const se = b.deleteShowInEdit !== false;
    if (ctx.formMode === 'view' && !sv) return false;
    if (ctx.formMode === 'edit' && !se) return false;
  }
  if (op === 'update' && ctx.formMode === 'create') return false;
  if (!userInAnyGroup(ctx.userGroupTitles, b.groupTitles)) return false;
  if (b.when && !evaluateCondition(b.when, ctx.values, ctx.dynamicContext)) return false;
  if (b.showOnlyWhenAllRequiredFilled === true && visibilityOpts?.allRequiredFilled !== true) return false;
  return true;
}

function ruleAppliesMode(rule: TFormRule, mode: TFormManagerFormMode): boolean {
  if (!rule.modes || rule.modes.length === 0) return true;
  return rule.modes.indexOf(mode) !== -1;
}

function ruleAppliesSubmit(rule: TFormRule, submitKind: TFormSubmitKind | undefined): boolean {
  if (submitKind !== 'draft') return true;
  const tags = rule.tags ?? [];
  return tags.indexOf(FULL_SUBMIT_TAG) === -1;
}

function parseIsoDate(s: string): Date | undefined {
  const d = new Date(s);
  return isNaN(d.getTime()) ? undefined : d;
}

/** dd/mm/aaaa ou dd-mm-aaaa (formato do formulário); depois ISO / Date nativo. */
function parseFormCalendarDateString(s: string): Date | undefined {
  const t = String(s).trim();
  if (!t) return undefined;
  const br = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(t);
  if (br) {
    const day = parseInt(br[1], 10);
    const month = parseInt(br[2], 10) - 1;
    const year = parseInt(br[3], 10);
    const dt = new Date(year, month, day);
    if (dt.getFullYear() !== year || dt.getMonth() !== month || dt.getDate() !== day) return undefined;
    return dt;
  }
  const d = new Date(t);
  return isNaN(d.getTime()) ? undefined : d;
}

function startOfDay(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function tryResolveEvaluatedDatePlusDaysString(s: string): string | undefined {
  const t = s.trim();
  const br = /^(\d{1,2}[-/]\d{1,2}[-/]\d{4})\s*([+-])\s*(\d+)$/.exec(t);
  if (br) {
    const base = parseFormCalendarDateString(br[1]);
    if (!base) return undefined;
    const sign = br[2] === '-' ? -1 : 1;
    const days = parseInt(br[3], 10);
    const out = startOfDay(base);
    out.setDate(out.getDate() + sign * days);
    return toIsoDateString(out);
  }
  const iso = /^(\d{4}-\d{2}-\d{2})\s*([+-])\s*(\d+)$/.exec(t);
  if (iso) {
    const base = parseIsoDate(iso[1]);
    if (!base) return undefined;
    const sign = iso[2] === '-' ? -1 : 1;
    const days = parseInt(iso[3], 10);
    const out = startOfDay(base);
    out.setDate(out.getDate() + sign * days);
    return toIsoDateString(out);
  }
  return undefined;
}

function resolveDatetimeComputedDisplayValue(
  expression: string,
  evaluated: unknown,
  values: Record<string, unknown>,
  dynamicContext: IDynamicContext
): string | undefined {
  const ex = expression.trim();
  if (ex) {
    const dr = resolveDateFieldDefaultValue(ex, dynamicContext, values);
    if (dr.kind === 'resolved' && dr.value !== undefined) {
      const raw = dr.value;
      const s = typeof raw === 'string' ? raw : String(raw);
      const d = parseFormCalendarDateString(s) ?? parseIsoDate(s);
      if (d) return toIsoDateString(startOfDay(d));
    }
  }
  if (typeof evaluated === 'number' && isFinite(evaluated)) {
    const base = startOfDay(new Date());
    base.setDate(base.getDate() + Math.trunc(evaluated));
    return toIsoDateString(base);
  }
  if (typeof evaluated === 'string' && evaluated.trim()) {
    const t = evaluated.trim();
    const plusDays = tryResolveEvaluatedDatePlusDaysString(t);
    if (plusDays) return plusDays;
    const d0 = parseFormCalendarDateString(t) ?? parseIsoDate(t);
    if (d0) return toIsoDateString(startOfDay(d0));
    const ms = Date.parse(t);
    if (!isNaN(ms)) return toIsoDateString(startOfDay(new Date(ms)));
  }
  return undefined;
}

function evalArithmetic(expr: string): number | undefined {
  let i = 0;
  const peek = (): string => expr[i] ?? '';
  const eat = (c: string): boolean => {
    if (peek() === c) {
      i++;
      return true;
    }
    return false;
  };
  const parseNumber = (): number | undefined => {
    let start = i;
    while (/[\d.]/.test(peek())) i++;
    if (start === i) return undefined;
    const n = Number(expr.slice(start, i));
    return isNaN(n) ? undefined : n;
  };
  let parseExpr: () => number | undefined;
  const parseFactor = (): number | undefined => {
    if (eat('(')) {
      const v = parseExpr();
      if (!eat(')')) return undefined;
      return v;
    }
    if (peek() === '-') {
      i++;
      const f = parseFactor();
      return f === undefined ? undefined : -f;
    }
    return parseNumber();
  };
  const parseTerm = (): number | undefined => {
    let left = parseFactor();
    if (left === undefined) return undefined;
    while (peek() === '*' || peek() === '/') {
      const op = peek();
      i++;
      const right = parseFactor();
      if (right === undefined) return undefined;
      left = op === '*' ? left * right : left / right;
    }
    return left;
  };
  parseExpr = (): number | undefined => {
    let left = parseTerm();
    if (left === undefined) return undefined;
    while (peek() === '+' || peek() === '-') {
      const op = peek();
      i++;
      const right = parseTerm();
      if (right === undefined) return undefined;
      left = op === '+' ? left + right : left - right;
    }
    return left;
  };
  const v = parseExpr();
  if (i !== expr.length) return undefined;
  return v;
}

function resolvePlaceholderScalarFromValues(innerRaw: string, values: Record<string, unknown>): string {
  const inner = String(innerRaw).trim();
  if (!inner || inner.indexOf('DAYS:') === 0) return '';
  const path = inner.split('/').map((p) => p.trim()).filter(Boolean);
  if (!path.length) return '';
  const rootVal = values[path[0]];
  const resolved = path.length > 1 ? readPathValue(rootVal, path.slice(1)) : rootVal;
  return toScalarString(resolved);
}

function resolvePlaceholderNumberFromValues(innerRaw: string, values: Record<string, unknown>): number | undefined {
  const inner = String(innerRaw).trim();
  if (!inner || inner.indexOf('DAYS:') === 0) return undefined;
  const path = inner.split('/').map((p) => p.trim()).filter(Boolean);
  if (!path.length) return undefined;
  const rootVal = values[path[0]];
  const resolved = path.length > 1 ? readPathValue(rootVal, path.slice(1)) : rootVal;
  const n = coerceNumber(resolved);
  return typeof n === 'number' && isFinite(n) ? n : undefined;
}

export function evaluateFormValueExpression(
  expr: string,
  values: Record<string, unknown>,
  dynamicContext?: IDynamicContext,
  attachmentFolderUrl?: IFormAttachmentFolderUrlContext
): unknown {
  const t = expr.trim();
  if (dynamicContext && isDynamicToken(t)) {
    const v = resolveStringToken(t, dynamicContext);
    return v !== undefined ? v : '';
  }
  if (t.indexOf(ATT_FOLDER_EXPR_PREFIX) === 0) {
    const nodeId = t.slice(ATT_FOLDER_EXPR_PREFIX.length).trim();
    if (!nodeId || !attachmentFolderUrl?.libraryRootServerRelativeUrl?.trim()) return '';
    const url = buildAttachmentFolderAbsoluteUrl({
      libraryRootServerRelativeUrl: attachmentFolderUrl.libraryRootServerRelativeUrl,
      itemId: attachmentFolderUrl.itemId,
      folderTree: attachmentFolderUrl.folderTree,
      folderNodeId: nodeId,
      itemFieldValues: values,
    });
    return url ?? '';
  }
  if (t.indexOf('str:') === 0) {
    let s = t.slice(4).replace(/\{\{([^}]+)\}\}/g, (_, name) => {
      const key = String(name).trim();
      if (key.indexOf('DAYS:') === 0) return '';
      return resolvePlaceholderScalarFromValues(key, values);
    });
    if (dynamicContext) {
      s = s.replace(/\[(.+?)\]/g, (full, inner: string) => {
        const tok = `[${String(inner).trim()}]`;
        if (!isDynamicToken(tok)) return full;
        const r = resolveStringToken(tok, dynamicContext);
        return r !== null && r !== undefined ? String(r) : '';
      });
    }
    return s;
  }
  if (dynamicContext && t.indexOf('[') !== -1) {
    const daysFirst = t.replace(/\{\{DAYS:([^:}]+):([^}]+)\}\}/g, (_, a, b) => {
      const sa = String(values[String(a).trim()] ?? '');
      const sb = String(values[String(b).trim()] ?? '');
      const da = parseFormCalendarDateString(sa) ?? parseIsoDate(sa);
      const db = parseFormCalendarDateString(sb) ?? parseIsoDate(sb);
      if (!da || !db) return '0';
      const ms = startOfDay(da).getTime() - startOfDay(db).getTime();
      return String(Math.round(ms / 86400000));
    });
    const withFields = daysFirst.replace(/\{\{([^}]+)\}\}/g, (_, name) => {
      const key = String(name).trim();
      if (key.indexOf('DAYS:') === 0) return '';
      return resolvePlaceholderScalarFromValues(key, values);
    });
    const withTok = withFields.replace(/\[(.+?)\]/g, (full, inner: string) => {
      const tok = `[${String(inner).trim()}]`;
      if (!isDynamicToken(tok)) return full;
      const r = resolveStringToken(tok, dynamicContext);
      return r !== null && r !== undefined ? String(r) : '';
    });
    if (/\[/.test(withTok)) return undefined;
    return withTok;
  }
  const withDays = t.replace(/\{\{DAYS:([^:}]+):([^}]+)\}\}/g, (_, a, b) => {
    const sa = String(values[String(a).trim()] ?? '');
    const sb = String(values[String(b).trim()] ?? '');
    const da = parseFormCalendarDateString(sa) ?? parseIsoDate(sa);
    const db = parseFormCalendarDateString(sb) ?? parseIsoDate(sb);
    if (!da || !db) return '0';
    const ms = startOfDay(da).getTime() - startOfDay(db).getTime();
    return String(Math.round(ms / 86400000));
  });

  const placeholderRe = /\{\{([^}]+)\}\}/g;
  let ph: RegExpExecArray | null;
  let anyPlaceholderNeedsText = false;
  placeholderRe.lastIndex = 0;
  while ((ph = placeholderRe.exec(withDays)) !== null) {
    const key = String(ph[1]).trim();
    if (key.indexOf('DAYS:') === 0) continue;
    const n = resolvePlaceholderNumberFromValues(key, values);
    if (typeof n === 'number' && isFinite(n)) continue;
    if (resolvePlaceholderScalarFromValues(key, values) !== '') {
      anyPlaceholderNeedsText = true;
      break;
    }
  }

  if (anyPlaceholderNeedsText) {
    return withDays.replace(/\{\{([^}]+)\}\}/g, (_, name) => {
      const key = String(name).trim();
      if (key.indexOf('DAYS:') === 0) return '';
      return resolvePlaceholderScalarFromValues(key, values);
    });
  }

  const replaced = withDays.replace(/\{\{([^}]+)\}\}/g, (_, name) => {
    const key = String(name).trim();
    if (key.indexOf('DAYS:') === 0) return '0';
    const n = resolvePlaceholderNumberFromValues(key, values);
    return typeof n === 'number' && isFinite(n) ? String(n) : '0';
  });
  const compactExpr = replaced.replace(/\s+/g, '');
  if (!/^[-+*/().0-9]+$/.test(compactExpr)) return undefined;
  return evalArithmetic(compactExpr);
}

function evaluateDaysOffsetExpr(
  exprRaw: string | undefined,
  values: Record<string, unknown>,
  dynamicContext: IDynamicContext
): number | undefined {
  const t = typeof exprRaw === 'string' ? exprRaw.trim() : '';
  if (!t) return undefined;
  const ev = evaluateFormValueExpression(t, values, dynamicContext);
  if (typeof ev === 'number' && isFinite(ev)) return Math.trunc(ev);
  if (typeof ev === 'string') {
    const n = Number(String(ev).trim());
    if (!isNaN(n) && isFinite(n)) return Math.trunc(n);
  }
  return undefined;
}

function resolveValidateDateMinMaxDaysFromToday(
  rule: import('../config/types/formManager').IFormRuleValidateDate,
  values: Record<string, unknown>,
  dynamicContext: IDynamicContext
): { minDays?: number; maxDays?: number } {
  let minDays: number | undefined;
  let maxDays: number | undefined;
  const exMin = rule.minDaysFromTodayExpr?.trim();
  if (exMin) {
    minDays = evaluateDaysOffsetExpr(exMin, values, dynamicContext);
  } else if (rule.minDaysFromToday !== undefined && typeof rule.minDaysFromToday === 'number') {
    minDays = Math.trunc(rule.minDaysFromToday);
  }
  const exMax = rule.maxDaysFromTodayExpr?.trim();
  if (exMax) {
    maxDays = evaluateDaysOffsetExpr(exMax, values, dynamicContext);
  } else if (rule.maxDaysFromToday !== undefined && typeof rule.maxDaysFromToday === 'number') {
    maxDays = Math.trunc(rule.maxDaysFromToday);
  }
  return { minDays, maxDays };
}

function validateDateRule(
  field: string,
  rule: import('../config/types/formManager').IFormRuleValidateDate,
  values: Record<string, unknown>,
  now: Date,
  dynamicContext: IDynamicContext
): string | undefined {
  const raw = values[field];
  if (isEmptyish(raw)) return undefined;
  const iso = typeof raw === 'string' ? raw : (raw instanceof Date ? raw.toISOString() : String(raw));
  const d = parseFormCalendarDateString(iso) ?? parseIsoDate(iso);
  if (!d) return rule.message ?? 'Data inválida.';
  const day = startOfDay(d);
  if (rule.blockWeekends === true) {
    const wd = day.getDay();
    if (wd === 0 || wd === 6) return rule.message ?? 'Fim de semana não permitido.';
  }
  if (rule.blockedWeekdays?.length) {
    const wd = day.getDay();
    if (rule.blockedWeekdays.indexOf(wd) !== -1) {
      return rule.message ?? 'Este dia da semana não é permitido.';
    }
  }
  if (rule.blockedIsoDates?.length) {
    const key = day.toISOString().slice(0, 10);
    for (let i = 0; i < rule.blockedIsoDates.length; i++) {
      const b = rule.blockedIsoDates[i].slice(0, 10);
      if (b === key) return rule.message ?? 'Data indisponível.';
    }
  }
  const today = startOfDay(now);
  const { minDays, maxDays } = resolveValidateDateMinMaxDaysFromToday(rule, values, dynamicContext);
  if (minDays !== undefined) {
    const minD = new Date(today.getTime());
    minD.setDate(minD.getDate() + minDays);
    if (day < startOfDay(minD)) return rule.message ?? 'Data abaixo do permitido.';
  }
  if (maxDays !== undefined) {
    const maxD = new Date(today.getTime());
    maxD.setDate(maxD.getDate() + maxDays);
    if (day > startOfDay(maxD)) return rule.message ?? 'Data acima do permitido.';
  }
  if (rule.minIso) {
    const min = parseIsoDate(rule.minIso);
    if (min && day < startOfDay(min)) return rule.message ?? 'Data abaixo do mínimo.';
  }
  if (rule.maxIso) {
    const max = parseIsoDate(rule.maxIso);
    if (max && day > startOfDay(max)) return rule.message ?? 'Data acima do máximo.';
  }
  if (rule.gteField) {
    const o = values[rule.gteField];
    const os = typeof o === 'string' ? o : String(o ?? '');
    const od = parseFormCalendarDateString(os) ?? parseIsoDate(os);
    if (od && day < startOfDay(od)) return rule.message ?? 'Data anterior à data de referência.';
  }
  if (rule.lteField) {
    const o = values[rule.lteField];
    const os = typeof o === 'string' ? o : String(o ?? '');
    const od = parseFormCalendarDateString(os) ?? parseIsoDate(os);
    if (od && day > startOfDay(od)) return rule.message ?? 'Data posterior à data de referência.';
  }
  if (rule.gtField) {
    const o = values[rule.gtField];
    const os = typeof o === 'string' ? o : String(o ?? '');
    const od = parseFormCalendarDateString(os) ?? parseIsoDate(os);
    if (od && day <= startOfDay(od)) return rule.message ?? 'Data deve ser posterior à referência.';
  }
  if (rule.ltField) {
    const o = values[rule.ltField];
    const os = typeof o === 'string' ? o : String(o ?? '');
    const od = parseFormCalendarDateString(os) ?? parseIsoDate(os);
    if (od && day >= startOfDay(od)) return rule.message ?? 'Data deve ser anterior à referência.';
  }
  return undefined;
}

export function evaluateValidateDateRulesForField(
  rules: readonly TFormRule[],
  field: string,
  values: Record<string, unknown>,
  params: {
    formMode: TFormManagerFormMode;
    submitKind: TFormSubmitKind | undefined;
    userGroupTitles: string[];
    dynamicContext: IDynamicContext;
    fieldVisible: (name: string) => boolean;
    now?: Date;
  }
): string | undefined {
  const { formMode, submitKind, userGroupTitles, dynamicContext, fieldVisible, now } = params;
  if (formMode === 'view') return undefined;
  const isDraft = submitKind === 'draft';
  const ts = now ?? new Date();
  for (let r = 0; r < rules.length; r++) {
    const rule = rules[r];
    if (rule.action !== 'validateDate') continue;
    if (rule.field !== field) continue;
    if (rule.enabled === false) continue;
    if (!ruleAppliesMode(rule, formMode)) continue;
    if (!ruleAppliesSubmit(rule, submitKind)) continue;
    if (!userInAnyGroup(userGroupTitles, rule.groupTitles)) continue;
    const whenOk = evaluateCondition(rule.when, values, dynamicContext);
    if (!whenOk) continue;
    if (!fieldVisible(field)) continue;
    if (isDraft) continue;
    const msg = validateDateRule(field, rule, values, ts, dynamicContext);
    if (msg) return msg;
  }
  return undefined;
}

function validateValueRule(
  field: string,
  rule: import('../config/types/formManager').IFormRuleValidateValue,
  values: Record<string, unknown>
): string | undefined {
  const raw = values[field];
  if (isEmptyish(raw)) return undefined;
  const s = String(raw);
  if (rule.minLength !== undefined && s.length < rule.minLength) return rule.message ?? `Mínimo ${rule.minLength} caracteres.`;
  if (rule.maxLength !== undefined && s.length > rule.maxLength) return rule.message ?? `Máximo ${rule.maxLength} caracteres.`;
  const n = coerceNumber(raw);
  if (rule.minNumber !== undefined && typeof n === 'number' && isFinite(n) && n < rule.minNumber) return rule.message ?? `Valor mínimo ${rule.minNumber}.`;
  if (rule.maxNumber !== undefined && typeof n === 'number' && isFinite(n) && n > rule.maxNumber) return rule.message ?? `Valor máximo ${rule.maxNumber}.`;
  if (rule.pattern) {
    try {
      // eslint-disable-next-line @rushstack/security/no-unsafe-regexp
      const re = new RegExp(rule.pattern);
      if (!re.test(s)) return rule.patternMessage ?? rule.message ?? 'Formato inválido.';
    } catch {
      return rule.message ?? 'Padrão inválido.';
    }
  }
  if (rule.allowList?.length) {
    if (rule.allowList.indexOf(s) === -1) return rule.message ?? 'Valor não permitido.';
  }
  if (rule.denyList?.length) {
    if (rule.denyList.indexOf(s) !== -1) return rule.message ?? 'Valor proibido.';
  }
  return undefined;
}

function multiCount(raw: unknown): number {
  if (Array.isArray(raw)) return raw.length;
  if (typeof raw === 'string' && raw.indexOf(';#') !== -1) return raw.split(';#').filter((x) => x.trim()).length;
  if (isEmptyish(raw)) return 0;
  return 1;
}

export interface IFormButtonVisibilityOverlay {
  show?: Set<string>;
  hide?: Set<string>;
}

export function buildFormDerivedState(
  cfg: IFormManagerConfig,
  fieldConfigs: IFormFieldConfig[],
  ctx: IFormRuleRuntimeContext,
  buttonOverlay?: IFormButtonVisibilityOverlay,
  fieldMetaByName?: ReadonlyMap<string, IFieldMetadata>
): IFormDerivedUiState {
  const { values, formMode, dynamicContext, attachmentFolderUrl } = ctx;
  const fieldVisible: Record<string, boolean> = {};
  const sectionVisible: Record<string, boolean> = {};
  const fieldRequired: Record<string, boolean> = {};
  const fieldDisabled: Record<string, boolean> = {};
  const fieldReadOnly: Record<string, boolean> = {};
  const effectiveSectionByField: Record<string, string> = {};
  const messages: IFormDerivedUiState['messages'] = [];
  const lookupFilters: IFormDerivedUiState['lookupFilters'] = {};
  const computedDisplay: Record<string, unknown> = {};
  const dynamicHelpByField: Record<string, string> = {};

  for (let i = 0; i < cfg.sections.length; i++) {
    const s = cfg.sections[i];
    sectionVisible[s.id] = s.visible !== false;
  }

  for (let i = 0; i < fieldConfigs.length; i++) {
    const f = fieldConfigs[i];
    fieldVisible[f.internalName] = f.visible !== false;
    fieldRequired[f.internalName] = f.required === true;
    fieldDisabled[f.internalName] = f.disabled === true;
    fieldReadOnly[f.internalName] = f.readOnly === true;
    if (f.sectionId) effectiveSectionByField[f.internalName] = f.sectionId;
  }

  if (formMode === 'view') {
    for (const k of Object.keys(fieldVisible)) fieldReadOnly[k] = true;
    for (const k of Object.keys(fieldDisabled)) fieldDisabled[k] = true;
  }

  const rules = cfg.rules ?? [];
  const visibilityPreferHide: Record<string, { anyHide: boolean; anyShow: boolean }> = {};
  for (let r = 0; r < rules.length; r++) {
    const rule = rules[r];
    if (rule.enabled === false) continue;
    if (!ruleAppliesMode(rule, formMode)) continue;
    if (!ruleAppliesSubmit(rule, ctx.submitKind)) continue;
    if (!userInAnyGroup(ctx.userGroupTitles, rule.groupTitles)) continue;
    const whenOk = evaluateCondition(rule.when, values, dynamicContext);
    if (!whenOk) continue;

    switch (rule.action) {
      case 'setVisibility': {
        if (rule.targetKind === 'section') {
          sectionVisible[rule.targetId] = rule.visibility === 'show';
          break;
        }
        const visTags = rule.tags;
        const preferHide =
          Array.isArray(visTags) && visTags.indexOf(FORM_VISIBILITY_PREFER_HIDE_TAG) !== -1;
        if (preferHide) {
          const tid = rule.targetId;
          if (!visibilityPreferHide[tid]) visibilityPreferHide[tid] = { anyHide: false, anyShow: false };
          if (rule.visibility === 'hide') visibilityPreferHide[tid].anyHide = true;
          else visibilityPreferHide[tid].anyShow = true;
        } else {
          fieldVisible[rule.targetId] = rule.visibility === 'show';
        }
        break;
      }
      case 'setRequired':
        fieldRequired[rule.field] = rule.required;
        break;
      case 'setDisabled':
        if (formMode !== 'view') fieldDisabled[rule.field] = rule.disabled;
        break;
      case 'setReadOnly':
        if (formMode !== 'view') fieldReadOnly[rule.field] = rule.readOnly;
        break;
      case 'setDefault':
        break;
      case 'clearFields':
        break;
      case 'filterLookupOptions':
        lookupFilters[rule.field] = {
          parentField: rule.parentField,
          ...(rule.childField ? { childField: rule.childField } : {}),
          ...(rule.filterOperator ? { filterOperator: rule.filterOperator } : {}),
          ...(rule.odataFilterTemplate ? { odataFilterTemplate: rule.odataFilterTemplate } : {}),
        };
        break;
      case 'setComputed': {
        if (fieldMetaByName) {
          const mtc = fieldMetaByName.get(rule.field)?.MappedType;
          if (mtc !== 'text' && mtc !== 'multiline' && mtc !== 'datetime') break;
        }
        const mtcField = fieldMetaByName?.get(rule.field)?.MappedType;
        let v = evaluateFormValueExpression(rule.expression, values, dynamicContext, attachmentFolderUrl);
        if (mtcField === 'datetime') {
          const disp = resolveDatetimeComputedDisplayValue(
            rule.expression ?? '',
            v,
            values,
            dynamicContext
          );
          if (disp !== undefined) computedDisplay[rule.field] = disp;
          break;
        }
        if (v !== undefined) {
          const fc = fieldConfigs.find((f) => f.internalName === rule.field);
          const tft = fc?.textValueTransform;
          if (tft && typeof v === 'string') v = applyFormFieldTextTransform(v, tft);
          computedDisplay[rule.field] = v;
        }
        break;
      }
      case 'setEffectiveSection':
        effectiveSectionByField[rule.field] = rule.sectionId;
        break;
      case 'showMessage':
        messages.push({ variant: rule.variant, text: rule.text, ruleId: rule.id });
        break;
      case 'profileVisibility': {
        const inG = userInAnyGroup(ctx.userGroupTitles, rule.groupTitles);
        if (rule.allow) fieldVisible[rule.field] = inG;
        else fieldVisible[rule.field] = !inG;
        break;
      }
      case 'profileEditable': {
        if (formMode === 'view') break;
        const inG = userInAnyGroup(ctx.userGroupTitles, rule.groupTitles);
        if (rule.allow) fieldReadOnly[rule.field] = !inG;
        else fieldReadOnly[rule.field] = inG;
        break;
      }
      case 'profileRequired': {
        const inG = userInAnyGroup(ctx.userGroupTitles, rule.groupTitles);
        if (rule.allow) fieldRequired[rule.field] = inG;
        else fieldRequired[rule.field] = !inG;
        break;
      }
      case 'authorFieldAccess': {
        const isAuthor = ctx.authorId !== undefined && ctx.authorId === ctx.currentUserId;
        if (formMode === 'view') break;
        if (!isAuthor) fieldReadOnly[rule.field] = true;
        break;
      }
      default:
        break;
    }
  }

  for (const vid of Object.keys(visibilityPreferHide)) {
    const m = visibilityPreferHide[vid];
    if (m.anyHide) fieldVisible[vid] = false;
    else if (m.anyShow) fieldVisible[vid] = true;
  }

  const dh = cfg.dynamicHelp ?? [];
  for (let i = 0; i < dh.length; i++) {
    const h = dh[i];
    if (evaluateCondition(h.when, values, dynamicContext)) dynamicHelpByField[h.field] = h.helpText;
  }

  if (buttonOverlay?.show) {
    buttonOverlay.show.forEach((k) => {
      if (k) fieldVisible[k] = true;
    });
  }
  if (buttonOverlay?.hide) {
    buttonOverlay.hide.forEach((k) => {
      if (k) fieldVisible[k] = false;
    });
  }

  return {
    fieldVisible,
    sectionVisible,
    fieldRequired,
    fieldDisabled,
    fieldReadOnly,
    effectiveSectionByField,
    messages,
    lookupFilters,
    computedDisplay,
    dynamicHelpByField,
  };
}

export interface IFormValidationAttachmentContext {
  attachmentCount?: number;
  pendingFiles?: { size: number; type: string; name: string }[];
}

export function collectFormValidationErrors(
  cfg: IFormManagerConfig,
  fieldConfigs: IFormFieldConfig[],
  ctx: IFormRuleRuntimeContext,
  attachmentCtx?: IFormValidationAttachmentContext,
  buttonOverlay?: IFormButtonVisibilityOverlay,
  metaByName?: ReadonlyMap<string, IFieldMetadata>
): Record<string, string> {
  const errors: Record<string, string> = {};
  const { values, formMode, submitKind, dynamicContext } = ctx;
  if (formMode === 'view') return errors;

  const derived = buildFormDerivedState(cfg, fieldConfigs, ctx, buttonOverlay, metaByName);
  const rules = cfg.rules ?? [];
  const fieldVisible = (name: string): boolean => derived.fieldVisible[name] !== false;
  const isDraft = submitKind === 'draft';

  if (!isDraft) {
    if (metaByName) {
      for (let i = 0; i < fieldConfigs.length; i++) {
        const fc = fieldConfigs[i];
        const name = fc.internalName;
        if (!fieldVisible(name)) continue;
        if (isFormBannerFieldConfig(fc)) continue;
        if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) {
          const attReq = derived.fieldRequired[name] === true;
          if (!attReq) continue;
          const readOnly = derived.fieldDisabled[name] === true;
          const pending = attachmentCtx?.pendingFiles?.length ?? 0;
          const existing = attachmentCtx?.attachmentCount ?? 0;
          const attSatisfied = pending > 0 || (formMode !== 'create' && existing > 0);
          if (!readOnly && !attSatisfied) {
            if (!errors._attachments) errors._attachments = 'Obrigatório.';
          }
          continue;
        }
        const m = metaByName.get(name);
        const mappedType = m?.MappedType ?? 'text';
        const isRequired = derived.fieldRequired[name] === true || m?.Required === true;
        if (!isRequired) continue;
        const readOnly = derived.fieldReadOnly[name] === true || derived.fieldDisabled[name] === true;
        if (readOnly) continue;
        if (isValueEmptyForRequiredField(values[name], mappedType)) errors[name] = 'Obrigatório.';
      }
    } else {
      for (let i = 0; i < fieldConfigs.length; i++) {
        const fc = fieldConfigs[i];
        const name = fc.internalName;
        if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) continue;
        if (isFormBannerFieldConfig(fc)) continue;
        if (!fieldVisible(name)) continue;
        const req = derived.fieldRequired[name] === true;
        if (req && isEmptyish(values[name])) errors[name] = 'Obrigatório.';
      }
    }
  }

  for (let r = 0; r < rules.length; r++) {
    const rule = rules[r];
    if (rule.enabled === false) continue;
    if (!ruleAppliesMode(rule, formMode)) continue;
    if (!ruleAppliesSubmit(rule, submitKind)) continue;
    if (!userInAnyGroup(ctx.userGroupTitles, rule.groupTitles)) continue;
    const whenOk = evaluateCondition(rule.when, values, dynamicContext);

    switch (rule.action) {
      case 'validateValue': {
        if (isDraft) break;
        if (!whenOk) break;
        if (!fieldVisible(rule.field)) break;
        const msg = validateValueRule(rule.field, rule, values);
        if (msg) errors[rule.field] = msg;
        break;
      }
      case 'validateDate': {
        if (isDraft) break;
        if (!whenOk) break;
        if (!fieldVisible(rule.field)) break;
        const msg = validateDateRule(rule.field, rule, values, ctx.dynamicContext.now ?? new Date(), dynamicContext);
        if (msg) errors[rule.field] = msg;
        break;
      }
      case 'atLeastOne': {
        if (isDraft) break;
        if (!whenOk) break;
        let any = false;
        for (let k = 0; k < rule.fields.length; k++) {
          if (!isEmptyish(values[rule.fields[k]])) {
            any = true;
            break;
          }
        }
        if (!any) {
          const key = rule.fields[0] ?? '_form';
          errors[key] = rule.message ?? 'Preencha ao menos um dos campos.';
        }
        break;
      }
      case 'multiMinMax': {
        if (isDraft) break;
        if (!whenOk) break;
        const c = multiCount(values[rule.field]);
        if (rule.min !== undefined && c < rule.min) errors[rule.field] = rule.message ?? `Selecione pelo menos ${rule.min}.`;
        if (rule.max !== undefined && c > rule.max) errors[rule.field] = rule.message ?? `No máximo ${rule.max}.`;
        break;
      }
      case 'setRequired': {
        if (isDraft) break;
        if (!whenOk) break;
        if (rule.field === FORM_ATTACHMENTS_FIELD_INTERNAL) break;
        if (rule.field.indexOf(FORM_BANNER_INTERNAL_PREFIX) === 0) break;
        if (!fieldVisible(rule.field)) break;
        if (rule.required && isEmptyish(values[rule.field])) errors[rule.field] = 'Obrigatório.';
        break;
      }
      case 'attachmentRules': {
        if (isDraft) break;
        const attWhen = rule.requiredWhen ? evaluateCondition(rule.requiredWhen, values, dynamicContext) : true;
        const count =
          (attachmentCtx?.attachmentCount ?? 0) +
          (attachmentCtx?.pendingFiles?.length ?? 0);
        if (rule.requiredWhen && attWhen && count < 1) {
          errors._attachments = rule.message ?? 'Anexo obrigatório.';
        }
        if (rule.minCount !== undefined && count < rule.minCount) {
          errors._attachments = rule.message ?? `Mínimo ${rule.minCount} anexo(s).`;
        }
        if (rule.maxCount !== undefined && count > rule.maxCount) {
          errors._attachments = rule.message ?? `Máximo ${rule.maxCount} anexo(s).`;
        }
        const files = attachmentCtx?.pendingFiles ?? [];
        if (rule.maxBytesPerFile !== undefined) {
          for (let fi = 0; fi < files.length; fi++) {
            if (files[fi].size > rule.maxBytesPerFile) {
              errors._attachments = rule.message ?? 'Arquivo excede o tamanho máximo.';
              break;
            }
          }
        }
        if (rule.allowedMimeTypes?.length && files.length > 0) {
          const allow = new Set(rule.allowedMimeTypes.map((x) => x.toLowerCase()));
          for (let fi = 0; fi < files.length; fi++) {
            const t = files[fi].type.toLowerCase();
            if (!allow.has(t) && !allow.has('*/*')) {
              errors._attachments = rule.message ?? 'Tipo de arquivo não permitido.';
              break;
            }
          }
        }
        const extList = rule.allowedFileExtensions;
        if (extList && extList.length > 0 && files.length > 0) {
          const allowExt = new Set(extList.map((x) => String(x).trim().replace(/^\./, '').toLowerCase()).filter(Boolean));
          for (let fi = 0; fi < files.length; fi++) {
            const nm = files[fi].name ?? '';
            const dot = nm.lastIndexOf('.');
            const ext = dot >= 0 && dot < nm.length - 1 ? nm.slice(dot + 1).toLowerCase() : '';
            if (!ext || !allowExt.has(ext)) {
              errors._attachments = rule.message ?? 'Extensão de ficheiro não permitida.';
              break;
            }
          }
        }
        break;
      }
      default:
        break;
    }
  }

  return errors;
}

export function buildFormFieldLabelMap(
  fieldConfigs: IFormFieldConfig[],
  metaByName: ReadonlyMap<string, IFieldMetadata>
): Map<string, string> {
  const m = new Map<string, string>();
  for (let i = 0; i < fieldConfigs.length; i++) {
    const fc = fieldConfigs[i];
    const name = fc.internalName;
    const meta = metaByName.get(name);
    const label = isFormBannerFieldConfig(fc)
      ? (fc.label?.trim() || 'Banner')
      : (fc.label?.trim() || meta?.Title?.trim() || name).trim() || name;
    m.set(name, label);
  }
  metaByName.forEach((meta, name) => {
    if (!m.has(name)) {
      m.set(name, (meta.Title?.trim() || name).trim() || name);
    }
  });
  return m;
}

function isSpecialValidationKey(key: string): boolean {
  if (key === '_attachments' || key === '_async' || key === '_linked') return true;
  return key.startsWith('_attf_');
}

function isGenericRequiredMessage(msg: string): boolean {
  const t = msg.trim().toLowerCase();
  return t === 'obrigatório.' || t === 'obrigatório' || t === 'obrigatório!';
}

export interface IFormValidationModalSection {
  heading: string;
  lines: string[];
}

function resolveSectionHeading(cfg: IFormManagerConfig, sectionId: string | undefined): string {
  if (!sectionId) return 'Formulário principal';
  const s = cfg.sections?.find((x) => x.id === sectionId);
  return (s?.title?.trim() || sectionId).trim() || 'Formulário principal';
}

function linkedChildAsManagerShell(cfg: IFormLinkedChildFormConfig): IFormManagerConfig {
  return {
    sections: cfg.sections,
    fields: cfg.fields,
    rules: cfg.rules,
    steps: cfg.steps,
    stepLayout: 'segmented',
  };
}

export function buildValidationModalSections(args: {
  mainErrors: Record<string, string>;
  formManager: IFormManagerConfig;
  fieldConfigs: IFormFieldConfig[];
  ctx: IFormRuleRuntimeContext;
  buttonOverlay?: IFormButtonVisibilityOverlay;
  fieldLabelByName: ReadonlyMap<string, string>;
  mainFieldMetaByName?: ReadonlyMap<string, IFieldMetadata>;
  linkedConfigs?: readonly IFormLinkedChildFormConfig[];
  linkedRowErrorsById?: Record<string, Record<string, string>[]>;
  linkedRowsById?: Record<string, { values: Record<string, unknown> }[]>;
  linkedMetaById?: Record<string, IFieldMetadata[]>;
  mainListLabel?: string;
}): IFormValidationModalSection[] {
  const {
    mainErrors,
    formManager,
    fieldConfigs,
    ctx,
    buttonOverlay,
    fieldLabelByName,
    mainFieldMetaByName,
    linkedConfigs = [],
    linkedRowErrorsById = {},
    linkedRowsById = {},
    linkedMetaById = {},
    mainListLabel,
  } = args;
  const derivedMain = buildFormDerivedState(
    formManager,
    fieldConfigs,
    ctx,
    buttonOverlay,
    mainFieldMetaByName
  );
  const map = new Map<string, string[]>();
  const push = (heading: string, line: string): void => {
    const arr = map.get(heading) ?? [];
    arr.push(line);
    map.set(heading, arr);
  };
  const mainRoot = (mainListLabel ?? '').trim() || 'Lista principal';

  for (const [key, raw] of Object.entries(mainErrors)) {
    const msg = String(raw ?? '').trim();
    if (!msg) continue;
    if (key === '_linked') continue;
    if (key === '_async') {
      push(`${mainRoot} · Validação`, msg);
      continue;
    }
    let sectionId = derivedMain.effectiveSectionByField[key];
    if (key === '_attachments' || key.startsWith('_attf_')) {
      sectionId = derivedMain.effectiveSectionByField[FORM_ATTACHMENTS_FIELD_INTERNAL] ?? sectionId;
    }
    const secTitle = resolveSectionHeading(formManager, sectionId);
    const heading = `${mainRoot} · ${secTitle}`;
    if (key === '_attachments' || key.startsWith('_attf_')) {
      const attLabel = 'Anexos';
      if (isGenericRequiredMessage(msg)) push(heading, `${attLabel}: obrigatório.`);
      else push(heading, `${attLabel}: ${msg}`);
      continue;
    }
    const label = fieldLabelByName.get(key) ?? key;
    if (isGenericRequiredMessage(msg)) push(heading, `${label}: obrigatório.`);
    else push(heading, `${label}: ${msg}`);
  }

  for (let ci = 0; ci < linkedConfigs.length; ci++) {
    const cfg = linkedConfigs[ci];
    const rowsErr = linkedRowErrorsById[cfg.id];
    if (!rowsErr?.length) continue;
    const blockRoot = (cfg.title?.trim() || cfg.listTitle.trim() || 'Lista vinculada').trim() || 'Lista vinculada';
    const meta = linkedMetaById[cfg.id] ?? [];
    const labelMap = buildFormFieldLabelMap(cfg.fields, new Map(meta.map((m) => [m.InternalName, m])));
    const rows = linkedRowsById[cfg.id] ?? [];
    const multi = rowsErr.filter((c) => c && Object.keys(c).length > 0).length > 1;
    for (let ri = 0; ri < rowsErr.length; ri++) {
      const cell = rowsErr[ri];
      if (!cell || Object.keys(cell).length === 0) continue;
      const rowVals = rows[ri]?.values ?? {};
      const ctxL: IFormRuleRuntimeContext = { ...ctx, values: rowVals };
      const shell = linkedChildAsManagerShell(cfg);
      const linkedMetaMap = new Map(meta.map((m) => [m.InternalName, m]));
      const derivedL = buildFormDerivedState(shell, cfg.fields, ctxL, undefined, linkedMetaMap);
      const rowSuffix = multi ? ` — Linha ${ri + 1}` : '';
      for (const [fk, fraw] of Object.entries(cell)) {
        const v = String(fraw ?? '').trim();
        if (!v) continue;
        if (fk === '_block') {
          push(`${blockRoot}${rowSuffix} · Geral`, v);
          continue;
        }
        const sid = fk === '_attachments' ? derivedL.effectiveSectionByField[FORM_ATTACHMENTS_FIELD_INTERNAL] : derivedL.effectiveSectionByField[fk];
        const secTitle = resolveSectionHeading(shell, sid);
        const heading = `${blockRoot}${rowSuffix} · ${secTitle}`;
        const lbl = labelMap.get(fk) ?? fk;
        if (isGenericRequiredMessage(v)) push(heading, `${lbl}: obrigatório.`);
        else push(heading, `${lbl}: ${v}`);
      }
    }
  }

  const out: IFormValidationModalSection[] = [];
  map.forEach((lines, heading) => {
    const u = Array.from(new Set(lines.filter((x) => x.trim())));
    if (u.length) out.push({ heading, lines: u });
  });
  if (!out.length) {
    const fb = formatValidationSummaryForForm(mainErrors, fieldLabelByName);
    return [{ heading: 'Validação', lines: [fb] }];
  }
  return out;
}

/** Texto único para MessageBar quando a validação falha (gravar ou mudar de etapa). */
export function formatValidationSummaryForForm(
  errors: Record<string, string>,
  labelByField?: ReadonlyMap<string, string>
): string {
  const entries = Object.entries(errors).filter(([, v]) => String(v).trim());
  if (!entries.length) return 'Corrija os problemas indicados antes de continuar.';

  const lines: string[] = [];
  for (let i = 0; i < entries.length; i++) {
    const key = entries[i][0];
    const raw = entries[i][1];
    const v = String(raw).trim();
    if (!v) continue;

    if (isSpecialValidationKey(key)) {
      lines.push(v);
      continue;
    }

    const label = labelByField?.get(key) ?? key;
    if (isGenericRequiredMessage(v)) {
      lines.push(`${label}: obrigatório.`);
      continue;
    }
    const lblLower = label.toLowerCase();
    const vLower = v.toLowerCase();
    if (vLower.startsWith(lblLower + ':') || vLower.startsWith(lblLower + ' :')) {
      lines.push(v);
    } else {
      lines.push(`${label}: ${v}`);
    }
  }

  const unique = Array.from(new Set(lines));
  if (unique.length === 1) return unique[0];
  return unique.map((l) => `• ${l}`).join('\n');
}

export function filterValidationErrorsToStepFields(
  errors: Record<string, string>,
  stepFieldNames: Set<string>
): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(errors)) {
    if (k === '_attachments') {
      if (stepFieldNames.has(FORM_ATTACHMENTS_FIELD_INTERNAL)) out[k] = v;
      continue;
    }
    if (k === '_async') continue;
    if (stepFieldNames.has(k)) out[k] = v;
  }
  return out;
}

export function pickRequiredStyleStepErrors(filtered: Record<string, string>): Record<string, string> {
  const out: Record<string, string> = {};
  for (const [k, v] of Object.entries(filtered)) {
    if (k === '_attachments') {
      const al = v.toLowerCase();
      if (
        al.includes('obrigatório') ||
        al.includes('obrigatorio') ||
        al.includes('mínimo') ||
        al.includes('minimo') ||
        al.includes('mín.') ||
        al.includes('extensão') ||
        al.includes('extensao')
      ) {
        out[k] = v;
      }
      continue;
    }
    const t = v.trim().toLowerCase();
    if (t === 'obrigatório.' || t === 'obrigatório' || t.startsWith('obrigatório')) {
      out[k] = v;
    }
  }
  return out;
}

export function getDefaultValuesFromRules(
  cfg: IFormManagerConfig,
  values: Record<string, unknown>,
  dynamicContext: IDynamicContext,
  opts?: IGetDefaultValuesFromRulesOptions
): Record<string, unknown> {
  const next = { ...values };
  const rules = cfg.rules ?? [];
  const isDt = opts?.isDateTimeField;
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (rule.action !== 'setDefault' || rule.enabled === false) continue;
    if (!evaluateCondition(rule.when, next, dynamicContext)) continue;

    let resolved: unknown;
    if (isDt?.(rule.field) === true && typeof rule.value === 'string') {
      const dr = resolveDateFieldDefaultValue(rule.value, dynamicContext, next);
      if (dr.kind === 'skip') continue;
      if (dr.kind === 'resolved') resolved = dr.value;
      else if (dr.kind === 'generic') {
        const ev = evaluateFormValueExpression(rule.value, next, dynamicContext);
        let iso: string | undefined;
        if (typeof ev === 'string') {
          const tt = ev.trim();
          iso =
            tryResolveEvaluatedDatePlusDaysString(tt) ??
            ((): string | undefined => {
              const d = parseFormCalendarDateString(tt) ?? parseIsoDate(tt);
              return d ? toIsoDateString(startOfDay(d)) : undefined;
            })();
        } else if (typeof ev === 'number' && isFinite(ev)) {
          const base = startOfDay(new Date());
          base.setDate(base.getDate() + Math.trunc(ev));
          iso = toIsoDateString(base);
        }
        resolved = iso !== undefined ? iso : tokenResolver.resolveValue(rule.value, dynamicContext);
      } else {
        resolved = tokenResolver.resolveValue(rule.value, dynamicContext);
      }
    } else {
      if (typeof rule.value === 'string') {
        const templateResolved = resolveDefaultTemplateValue(rule.value, next, dynamicContext);
        resolved = templateResolved.usedTemplate
          ? templateResolved.value
          : tokenResolver.resolveValue(rule.value, dynamicContext);
      } else {
        resolved = tokenResolver.resolveValue(rule.value, dynamicContext);
      }
    }

    if (resolved !== undefined && isEmptyish(next[rule.field])) next[rule.field] = resolved;
  }
  return next;
}

export function expressionReferencesSharePointItemId(expression: string): boolean {
  const re = /\{\{([^}]+)\}\}/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(expression)) !== null) {
    const base = String(m[1] ?? '')
      .trim()
      .split('/')[0]
      ?.trim();
    if (base && base.toLowerCase() === 'id') return true;
  }
  return false;
}

export function buildPostCreateItemIdComputedPatch(params: {
  cfg: IFormManagerConfig;
  fieldConfigs: IFormFieldConfig[];
  values: Record<string, unknown>;
  dynamicContext: IDynamicContext;
  attachmentFolderUrl?: IFormAttachmentFolderUrlContext;
  userGroupTitles: string[];
  submitKind: TFormSubmitKind | undefined;
  newItemId: number;
  fieldMetaByName: ReadonlyMap<string, IFieldMetadata>;
}): Record<string, unknown> {
  const {
    cfg,
    fieldConfigs,
    values,
    dynamicContext,
    attachmentFolderUrl,
    userGroupTitles,
    submitKind,
    newItemId,
    fieldMetaByName,
  } = params;

  const valuesWithId: Record<string, unknown> = {
    ...values,
    Id: newItemId,
    ID: newItemId,
  };

  const out: Record<string, unknown> = {};
  const rules = cfg.rules ?? [];

  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (rule.action !== 'setComputed') continue;
    if (rule.enabled === false) continue;
    if (!ruleAppliesMode(rule, 'create')) continue;
    if (!ruleAppliesSubmit(rule, submitKind)) continue;
    if (!userInAnyGroup(userGroupTitles, rule.groupTitles)) continue;
    if (rule.when && !evaluateCondition(rule.when, valuesWithId, dynamicContext)) continue;

    const expr = rule.expression ?? '';
    if (!expressionReferencesSharePointItemId(expr)) continue;

    const mtc = fieldMetaByName.get(rule.field)?.MappedType;
    if (mtc !== 'text' && mtc !== 'multiline' && mtc !== 'datetime') continue;

    let v = evaluateFormValueExpression(expr, valuesWithId, dynamicContext, attachmentFolderUrl);

    if (mtc === 'datetime') {
      const disp = resolveDatetimeComputedDisplayValue(expr, v, valuesWithId, dynamicContext);
      if (disp !== undefined) out[rule.field] = disp;
      continue;
    }

    if (v === undefined) continue;

    const fc = fieldConfigs.find((f) => f.internalName === rule.field);
    const tft = fc?.textValueTransform;
    if (tft && typeof v === 'string') v = applyFormFieldTextTransform(v, tft);

    out[rule.field] = v;
  }

  return out;
}
