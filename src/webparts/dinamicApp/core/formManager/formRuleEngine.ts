import type { IDynamicContext } from '../dynamicTokens/types';
import { DynamicTokenResolver } from '../dynamicTokens/services/DynamicTokenResolver';
import type {
  IFormManagerConfig,
  IFormFieldConfig,
  IFormCustomButtonConfig,
  TFormConditionNode,
  TFormRule,
  TFormManagerFormMode,
  TFormSubmitKind,
  IFormCompareRef,
  TFormConditionOp,
  TFormCustomButtonOperation,
} from '../config/types/formManager';
import { FORM_ATTACHMENTS_FIELD_INTERNAL } from '../config/types/formManager';

const FULL_SUBMIT_TAG = 'fullSubmitOnly';

export interface IFormRuleRuntimeContext {
  formMode: TFormManagerFormMode;
  values: Record<string, unknown>;
  submitKind?: TFormSubmitKind;
  userGroupTitles: string[];
  currentUserId: number;
  authorId?: number;
  dynamicContext: IDynamicContext;
}

export interface IFormDerivedUiState {
  fieldVisible: Record<string, boolean>;
  sectionVisible: Record<string, boolean>;
  fieldRequired: Record<string, boolean>;
  fieldDisabled: Record<string, boolean>;
  fieldReadOnly: Record<string, boolean>;
  effectiveSectionByField: Record<string, string>;
  messages: { variant: 'info' | 'warning' | 'error'; text: string; ruleId: string }[];
  lookupFilters: Record<string, { parentField: string; odataFilterTemplate: string }>;
  computedDisplay: Record<string, unknown>;
  dynamicHelpByField: Record<string, string>;
}

const tokenResolver = new DynamicTokenResolver();

function normGroupTitle(s: string): string {
  return s.trim().toLowerCase();
}

function userInAnyGroup(userTitles: string[], ruleGroups: string[] | undefined): boolean {
  if (!ruleGroups || ruleGroups.length === 0) return true;
  const set = new Set(userTitles.map(normGroupTitle));
  for (let i = 0; i < ruleGroups.length; i++) {
    if (set.has(normGroupTitle(ruleGroups[i]))) return true;
  }
  return false;
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

function coerceBool(v: unknown): boolean | undefined {
  if (v === true || v === false) return v;
  if (v === 1 || v === '1' || v === 'true' || v === 'True') return true;
  if (v === 0 || v === '0' || v === 'false' || v === 'False') return false;
  return undefined;
}

function coerceNumber(v: unknown): number {
  if (typeof v === 'number' && !isNaN(v)) return v;
  if (typeof v === 'string' && v.trim() !== '') {
    const n = Number(v.replace(',', '.'));
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

export function shouldShowCustomButton(b: IFormCustomButtonConfig, ctx: IFormRuleRuntimeContext): boolean {
  if (b.enabled === false) return false;
  if (b.modes !== undefined && b.modes.length === 0) return false;
  if (b.modes?.length && b.modes.indexOf(ctx.formMode) === -1) return false;
  const op: TFormCustomButtonOperation = b.operation ?? 'legacy';
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

function startOfDay(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
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

export function evaluateFormValueExpression(expr: string, values: Record<string, unknown>): unknown {
  const t = expr.trim();
  if (t.indexOf('str:') === 0) {
    return t.slice(4).replace(/\{\{([^}]+)\}\}/g, (_, name) => {
      const key = String(name).trim();
      const v = values[key];
      if (v === null || v === undefined) return '';
      if (typeof v === 'object' && v !== null && 'Title' in (v as object)) return String((v as Record<string, unknown>).Title ?? '');
      return String(v);
    });
  }
  const withDays = t.replace(/\{\{DAYS:([^:}]+):([^}]+)\}\}/g, (_, a, b) => {
    const da = parseIsoDate(String(values[String(a).trim()] ?? ''));
    const db = parseIsoDate(String(values[String(b).trim()] ?? ''));
    if (!da || !db) return '0';
    const ms = startOfDay(da).getTime() - startOfDay(db).getTime();
    return String(Math.round(ms / 86400000));
  });
  const replaced = withDays.replace(/\{\{([^}]+)\}\}/g, (_, name) => {
    const n = coerceNumber(values[String(name).trim()]);
    return typeof n === 'number' && isFinite(n) ? String(n) : '0';
  });
  if (!/^[-+*/().0-9]+$/.test(replaced)) return undefined;
  return evalArithmetic(replaced);
}

function validateDateRule(
  field: string,
  rule: import('../config/types/formManager').IFormRuleValidateDate,
  values: Record<string, unknown>,
  now: Date
): string | undefined {
  const raw = values[field];
  if (isEmptyish(raw)) return undefined;
  const iso = typeof raw === 'string' ? raw : (raw instanceof Date ? raw.toISOString() : String(raw));
  const d = parseIsoDate(iso);
  if (!d) return rule.message ?? 'Data inválida.';
  const day = startOfDay(d);
  if (rule.blockWeekends === true) {
    const wd = day.getDay();
    if (wd === 0 || wd === 6) return rule.message ?? 'Fim de semana não permitido.';
  }
  if (rule.blockedIsoDates?.length) {
    const key = day.toISOString().slice(0, 10);
    for (let i = 0; i < rule.blockedIsoDates.length; i++) {
      const b = rule.blockedIsoDates[i].slice(0, 10);
      if (b === key) return rule.message ?? 'Data indisponível.';
    }
  }
  const today = startOfDay(now);
  if (rule.minDaysFromToday !== undefined) {
    const minD = new Date(today.getTime());
    minD.setDate(minD.getDate() + rule.minDaysFromToday);
    if (day < startOfDay(minD)) return rule.message ?? 'Data abaixo do permitido.';
  }
  if (rule.maxDaysFromToday !== undefined) {
    const maxD = new Date(today.getTime());
    maxD.setDate(maxD.getDate() + rule.maxDaysFromToday);
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
    const od = parseIsoDate(typeof o === 'string' ? o : String(o ?? ''));
    if (od && day < startOfDay(od)) return rule.message ?? 'Data anterior à data de referência.';
  }
  if (rule.lteField) {
    const o = values[rule.lteField];
    const od = parseIsoDate(typeof o === 'string' ? o : String(o ?? ''));
    if (od && day > startOfDay(od)) return rule.message ?? 'Data posterior à data de referência.';
  }
  if (rule.gtField) {
    const o = values[rule.gtField];
    const od = parseIsoDate(typeof o === 'string' ? o : String(o ?? ''));
    if (od && day <= startOfDay(od)) return rule.message ?? 'Data deve ser posterior à referência.';
  }
  if (rule.ltField) {
    const o = values[rule.ltField];
    const od = parseIsoDate(typeof o === 'string' ? o : String(o ?? ''));
    if (od && day >= startOfDay(od)) return rule.message ?? 'Data deve ser anterior à referência.';
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
  buttonOverlay?: IFormButtonVisibilityOverlay
): IFormDerivedUiState {
  const { values, formMode, dynamicContext } = ctx;
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
        if (rule.targetKind === 'section') sectionVisible[rule.targetId] = rule.visibility === 'show';
        else fieldVisible[rule.targetId] = rule.visibility === 'show';
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
        lookupFilters[rule.field] = { parentField: rule.parentField, odataFilterTemplate: rule.odataFilterTemplate };
        break;
      case 'setComputed': {
        const v = evaluateFormValueExpression(rule.expression, values);
        if (v !== undefined) computedDisplay[rule.field] = v;
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
  pendingFiles?: { size: number; type: string }[];
}

export function collectFormValidationErrors(
  cfg: IFormManagerConfig,
  fieldConfigs: IFormFieldConfig[],
  ctx: IFormRuleRuntimeContext,
  attachmentCtx?: IFormValidationAttachmentContext,
  buttonOverlay?: IFormButtonVisibilityOverlay
): Record<string, string> {
  const errors: Record<string, string> = {};
  const { values, formMode, submitKind, dynamicContext } = ctx;
  if (formMode === 'view') return errors;

  const derived = buildFormDerivedState(cfg, fieldConfigs, ctx, buttonOverlay);
  const rules = cfg.rules ?? [];
  const fieldVisible = (name: string): boolean => derived.fieldVisible[name] !== false;
  const isDraft = submitKind === 'draft';

  if (!isDraft) {
    for (let i = 0; i < fieldConfigs.length; i++) {
      const fc = fieldConfigs[i];
      const name = fc.internalName;
      if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) continue;
      if (!fieldVisible(name)) continue;
      const req = derived.fieldRequired[name] === true;
      if (req && isEmptyish(values[name])) errors[name] = 'Obrigatório.';
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
        const msg = validateDateRule(rule.field, rule, values, ctx.dynamicContext.now ?? new Date());
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
        break;
      }
      default:
        break;
    }
  }

  return errors;
}

export function getDefaultValuesFromRules(
  cfg: IFormManagerConfig,
  values: Record<string, unknown>,
  dynamicContext: IDynamicContext
): Record<string, unknown> {
  const next = { ...values };
  const rules = cfg.rules ?? [];
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    if (rule.action !== 'setDefault' || rule.enabled === false) continue;
    if (!evaluateCondition(rule.when, next, dynamicContext)) continue;
    const resolved = tokenResolver.resolveValue(rule.value, dynamicContext);
    if (resolved !== undefined && isEmptyish(next[rule.field])) next[rule.field] = resolved;
  }
  return next;
}
