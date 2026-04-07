import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  Checkbox,
  Spinner,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  Link,
  Icon,
  IconButton,
  Toggle,
} from '@fluentui/react';
import { FieldsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import type {
  IFormManagerConfig,
  IFormStepNavigationConfig,
  IFormFieldConfig,
  IFormSectionConfig,
  IFormStepConfig,
  IFormCustomButtonConfig,
  TFormButtonAction,
  TFormConditionOp,
  TFormCustomButtonBehavior,
  TFormCustomButtonOperation,
  TFormManagerFormMode,
  TFormRule,
  TFormConditionNode,
  TFormStepLayoutKind,
  TFormStepNavButtonsKind,
  TFormDataLoadingUiKind,
  TFormSubmitLoadingUiKind,
} from '../../core/config/types/formManager';
import { FORM_ATTACHMENTS_FIELD_INTERNAL, FORM_OCULTOS_STEP_ID } from '../../core/config/types/formManager';
import { getDefaultFormManagerConfig } from '../../core/config/utils';
import { sanitizeFormManagerConfig } from '../../core/formManager/sanitizeFormManagerConfig';
import {
  buildFieldUiRules,
  compileConditionalCard,
  customRulesOnly,
  describeConditionalCardPT,
  describeRule,
  mergeAttachmentUiRule,
  mergeCardRulesIntoAll,
  mergeFieldRules,
  newCardId,
  parseAttachmentUiRule,
  parseConditionalCardsFromRules,
  countFieldUiRules,
  CONDITIONAL_EFFECT_OPTIONS,
  CONDITION_OP_OPTIONS,
  type IConditionalEffectUi,
  type IConditionalRuleCard,
  type IWhenUi,
  type TConditionalEffectKind,
  fieldRuleStateFromRules,
  templateConditionalShowWhenEquals,
  templateFieldRulesChoiceRequiresOther,
  whenUiToNode,
  whenNodeToUi,
} from '../../core/formManager/formManagerVisualModel';
import { FormFieldRulesPanel } from './FormFieldRulesPanel';
import { FormManagerComponentsTabContent } from './FormManagerComponentsTab';
import {
  FORM_SUBMIT_LOADING_DROPDOWN_OPTIONS,
  FORM_SUBMIT_LOADING_INHERIT_KEY,
} from './FormLoadingUi';

function newId(prefix: string): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
}

function reorderByIndex<T>(arr: T[], from: number, to: number): T[] {
  if (from === to || from < 0 || to < 0 || from >= arr.length || to >= arr.length) return arr.slice();
  const next = arr.slice();
  const moved = next.splice(from, 1);
  const item = moved[0] as T;
  next.splice(to, 0, item);
  return next;
}

const REQ_FIELD_BG_OK = '#dff6dd';
const REQ_FIELD_BG_BAD = '#fde7e9';
const REQ_FIELD_BORDER_OK = '#92c353';
const REQ_FIELD_BORDER_BAD = '#d13438';

function hasAnyFieldInAnyStep(steps: IFormStepConfig[]): boolean {
  for (let s = 0; s < steps.length; s++) {
    if (steps[s].fieldNames.length > 0) return true;
  }
  return false;
}

function requiredListFieldIsSatisfied(
  m: IFieldMetadata,
  steps: IFormStepConfig[],
  fields: IFormFieldConfig[]
): boolean {
  if (!m.Required) return true;
  const hasAny = hasAnyFieldInAnyStep(steps);
  const inSteps = steps.some((st) => st.fieldNames.indexOf(m.InternalName) !== -1);
  const inFields = fields.some((f) => f.internalName === m.InternalName);
  if (!hasAny) return inFields;
  return inSteps;
}

function requiredFieldRowStyles(
  m: IFieldMetadata | undefined,
  steps: IFormStepConfig[],
  fields: IFormFieldConfig[]
): { background: string; border: string } | undefined {
  if (!m || !m.Required) return undefined;
  const ok = requiredListFieldIsSatisfied(m, steps, fields);
  return {
    background: ok ? REQ_FIELD_BG_OK : REQ_FIELD_BG_BAD,
    border: `1px solid ${ok ? REQ_FIELD_BORDER_OK : REQ_FIELD_BORDER_BAD}`,
  };
}

function requiredListFieldsMissingFromSteps(meta: IFieldMetadata[], steps: IFormStepConfig[]): IFieldMetadata[] {
  if (!hasAnyFieldInAnyStep(steps)) return [];
  const inSteps = new Set<string>();
  for (let s = 0; s < steps.length; s++) {
    const fn = steps[s].fieldNames;
    for (let i = 0; i < fn.length; i++) inSteps.add(fn[i]);
  }
  const missing: IFieldMetadata[] = [];
  for (let m = 0; m < meta.length; m++) {
    const f = meta[m];
    if (!f.Required) continue;
    if (!inSteps.has(f.InternalName)) missing.push(f);
  }
  return missing;
}

/** Campos de sistema que não entram no pool / dropdowns de campos do gestor. */
const FORM_CONFIG_UI_EXCLUDED_FIELD_INTERNALS = new Set([
  'Attachments',
  'ContentType',
  'ContentTypeId',
]);

function isFormConfigSelectableField(m: IFieldMetadata): boolean {
  return !FORM_CONFIG_UI_EXCLUDED_FIELD_INTERNALS.has(m.InternalName);
}

const DND_FIELD = 'fm/field:';
const DND_STEP = 'fm/step:';
const DND_MCOL = 'fm/mcol:';
const DND_POOL = 'fm/pool:';
const DND_FS = 'fm/fs:';
const DND_BTN = 'fm/btn:';

function dragPayload(kind: string, index: number): string {
  return kind + String(index);
}

function parseDragIndex(data: string, prefix: string): number | undefined {
  if (data.indexOf(prefix) !== 0) return undefined;
  const n = parseInt(data.slice(prefix.length), 10);
  return isNaN(n) ? undefined : n;
}

function dragPayloadPool(internalName: string): string {
  return DND_POOL + encodeURIComponent(internalName);
}

function parsePoolDrag(data: string): string | undefined {
  if (data.indexOf(DND_POOL) !== 0) return undefined;
  try {
    return decodeURIComponent(data.slice(DND_POOL.length));
  } catch {
    return undefined;
  }
}

function dragPayloadFieldInStep(stepIdx: number, idxInStep: number, internalName: string): string {
  return DND_FS + String(stepIdx) + ':' + String(idxInStep) + ':' + encodeURIComponent(internalName);
}

function parseFieldInStepDrag(data: string): { fromStep: number; fromIdx: number; name: string } | undefined {
  if (data.indexOf(DND_FS) !== 0) return undefined;
  const rest = data.slice(DND_FS.length);
  const p1 = rest.indexOf(':');
  const p2 = rest.indexOf(':', p1 + 1);
  if (p1 === -1 || p2 === -1) return undefined;
  const fromStep = parseInt(rest.slice(0, p1), 10);
  const fromIdx = parseInt(rest.slice(p1 + 1, p2), 10);
  let name = '';
  try {
    name = decodeURIComponent(rest.slice(p2 + 1));
  } catch {
    return undefined;
  }
  if (isNaN(fromStep) || isNaN(fromIdx) || !name) return undefined;
  return { fromStep, fromIdx, name };
}

function insertFieldNameIntoStep(
  st: IFormStepConfig[],
  fieldName: string,
  toStepIdx: number,
  insertBefore: number
): IFormStepConfig[] {
  const next = st.map((s) => ({
    ...s,
    fieldNames: s.fieldNames.filter((n) => n !== fieldName),
  }));
  const tgt = next[toStepIdx];
  if (!tgt) return next;
  const fn = tgt.fieldNames.slice();
  const pos = Math.max(0, Math.min(insertBefore, fn.length));
  fn.splice(pos, 0, fieldName);
  next[toStepIdx] = { ...tgt, fieldNames: fn };
  return next;
}

function fieldsAlignedToSteps(flds: IFormFieldConfig[], st: IFormStepConfig[]): IFormFieldConfig[] {
  const byName: Record<string, IFormFieldConfig> = {};
  for (let i = 0; i < flds.length; i++) {
    byName[flds[i].internalName] = flds[i];
  }
  const out: IFormFieldConfig[] = [];
  const seen: Record<string, boolean> = {};
  for (let s = 0; s < st.length; s++) {
    const sid = st[s].id;
    for (let j = 0; j < st[s].fieldNames.length; j++) {
      const n = st[s].fieldNames[j];
      const fc = byName[n];
      if (fc) {
        out.push({ ...fc, sectionId: sid });
        seen[n] = true;
      }
    }
  }
  for (let i = 0; i < flds.length; i++) {
    const n = flds[i].internalName;
    if (!seen[n]) {
      out.push({ ...flds[i], sectionId: st[0]?.id ?? flds[i].sectionId });
    }
  }
  return out;
}

function resyncStepsOrderFromFields(flds: IFormFieldConfig[], st: IFormStepConfig[]): IFormStepConfig[] {
  const orderMap: Record<string, number> = {};
  for (let i = 0; i < flds.length; i++) {
    orderMap[flds[i].internalName] = i;
  }
  return st.map((s) => ({
    ...s,
    fieldNames: s.fieldNames.slice().sort((a, b) => (orderMap[a] ?? 99999) - (orderMap[b] ?? 99999)),
  }));
}

function numOpt(s: string): number | undefined {
  const t = s.trim();
  if (!t) return undefined;
  const n = Number(t);
  return isNaN(n) ? undefined : n;
}

function defaultWhenUi(meta: IFieldMetadata[]): IWhenUi {
  const f = meta[0]?.InternalName ?? 'Title';
  return { field: f, op: 'eq', compareKind: 'literal', compareValue: '' };
}

function emptyEffect(): IConditionalEffectUi {
  return { kind: 'showField', targetField: '' };
}

function parseCsvFieldNames(s: string): string[] {
  return s
    .split(/[,;]/)
    .map((x) => x.trim())
    .filter(Boolean);
}

function fieldNamesToCsv(names: string[]): string {
  return names.join(', ');
}

function buttonSetFieldValueChoiceDropdown(
  fieldInternalName: string,
  valueTemplate: string | undefined,
  fieldMeta: IFieldMetadata[]
): { options: IDropdownOption[]; selectedKey: string } | null {
  const tpl = valueTemplate ?? '';
  const low = tpl.trim().toLowerCase();
  if (low.length >= 4 && low.slice(0, 4) === 'str:') {
    return null;
  }
  let fm: IFieldMetadata | undefined;
  for (let i = 0; i < fieldMeta.length; i++) {
    if (fieldMeta[i].InternalName === fieldInternalName) {
      fm = fieldMeta[i];
      break;
    }
  }
  const choices =
    fm && fm.MappedType === 'choice' && fm.Choices && fm.Choices.length > 0 ? fm.Choices : null;
  if (!choices) {
    return null;
  }
  const opts: IDropdownOption[] = [{ key: '', text: '—' }];
  for (let i = 0; i < choices.length; i++) {
    const c = choices[i];
    opts.push({ key: c, text: c });
  }
  if (tpl && choices.indexOf(tpl) === -1) {
    opts.push({ key: tpl, text: `${tpl} (valor atual)` });
  }
  return { options: opts, selectedKey: tpl };
}

const REDIRECT_KEY_FORM = '__FORM__';
const REDIRECT_KEY_FORMID = '__FORMID__';

function redirectTokenForKey(key: string): string {
  if (key === REDIRECT_KEY_FORM) return '{{Form}}';
  if (key === REDIRECT_KEY_FORMID) return '{{FormID}}';
  return `{{${key}}}`;
}

function replaceFirstEmptyRedirectBrace(url: string, key: string): string {
  return url.replace(/\{\{\s*\}\}/, redirectTokenForKey(key));
}

const BUTTON_OPERATION_OPTIONS: IDropdownOption[] = [
  { key: 'legacy', text: 'Ações em cadeia (rascunho, enviar, fechar…)' },
  { key: 'redirect', text: 'Redirecionar (URL com {{campo}})' },
  { key: 'add', text: 'Adicionar — criar novo item na lista' },
  { key: 'update', text: 'Atualizar — gravar o item atual (Form/FormID)' },
  { key: 'delete', text: 'Eliminar — apagar o item atual' },
];

const BUTTON_BEHAVIOR_OPTIONS: IDropdownOption[] = [
  { key: 'actionsOnly', text: 'Só executar ações' },
  { key: 'draft', text: 'Ações e depois rascunho' },
  { key: 'submit', text: 'Ações e depois enviar' },
  { key: 'close', text: 'Ações e depois fechar formulário' },
];

const BUTTON_ACTION_KIND_OPTIONS: IDropdownOption[] = [
  { key: 'showFields', text: 'Mostrar campos' },
  { key: 'hideFields', text: 'Ocultar campos' },
  { key: 'setFieldValue', text: 'Definir valor de um campo' },
  { key: 'joinFields', text: 'Juntar vários campos num campo' },
];

function defaultActionForKind(kind: TFormButtonAction['kind']): TFormButtonAction {
  switch (kind) {
    case 'hideFields':
      return { kind: 'hideFields', fields: [] };
    case 'setFieldValue':
      return { kind: 'setFieldValue', field: '', valueTemplate: '' };
    case 'joinFields':
      return { kind: 'joinFields', targetField: '', sourceFields: [], separator: ' ' };
    default:
      return { kind: 'showFields', fields: [] };
  }
}

function modesFromCheckboxes(c: boolean, e: boolean, v: boolean): TFormManagerFormMode[] | undefined {
  if (c && e && v) return undefined;
  const out: TFormManagerFormMode[] = [];
  if (c) out.push('create');
  if (e) out.push('edit');
  if (v) out.push('view');
  return out;
}

function checkboxesFromModes(modes: TFormManagerFormMode[] | undefined): {
  c: boolean;
  e: boolean;
  v: boolean;
} {
  if (!modes || modes.length === 0) return { c: true, e: true, v: true };
  return {
    c: modes.indexOf('create') !== -1,
    e: modes.indexOf('edit') !== -1,
    v: modes.indexOf('view') !== -1,
  };
}

function sectionsFromSteps(steps: IFormStepConfig[]): IFormSectionConfig[] {
  const out: IFormSectionConfig[] = [];
  for (let i = 0; i < steps.length; i++) {
    out.push({ id: steps[i].id, title: steps[i].title, visible: true });
  }
  return out;
}

function inferStepsFromLegacy(sections: IFormSectionConfig[], flds: IFormFieldConfig[]): IFormStepConfig[] {
  const out: IFormStepConfig[] = [];
  const defaultSid = sections[0]?.id ?? 'main';
  for (let i = 0; i < sections.length; i++) {
    const sec = sections[i];
    const fieldNames: string[] = [];
    for (let j = 0; j < flds.length; j++) {
      const sid = flds[j].sectionId ?? defaultSid;
      if (sid === sec.id) fieldNames.push(flds[j].internalName);
    }
    out.push({ id: sec.id, title: sec.title, fieldNames: fieldNames.slice() });
  }
  if (out.length === 0) {
    const fn: string[] = [];
    for (let k = 0; k < flds.length; k++) {
      fn.push(flds[k].internalName);
    }
    return ensureCoreSteps([{ id: 'main', title: 'Geral', fieldNames: fn }]);
  }
  return ensureCoreSteps(out);
}

function ensureCoreSteps(st: IFormStepConfig[]): IFormStepConfig[] {
  const out = st.map((s) => ({ ...s, fieldNames: s.fieldNames.slice() }));
  if (out.length === 0) {
    return [
      { id: 'main', title: 'Geral', fieldNames: [] },
      { id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] },
    ];
  }
  if (!out.some((s) => s.id === 'main')) {
    out.unshift({ id: 'main', title: 'Geral', fieldNames: [] });
  }
  if (!out.some((s) => s.id === FORM_OCULTOS_STEP_ID)) {
    out.push({ id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] });
  }
  return out;
}

function buildInitialFieldsAndSteps(v: IFormManagerConfig): {
  fields: IFormFieldConfig[];
  steps: IFormStepConfig[];
} {
  const stepsSrc =
    v.steps && v.steps.length > 0
      ? v.steps.map((st) => ({ ...st, fieldNames: st.fieldNames.slice() }))
      : inferStepsFromLegacy(v.sections, v.fields);
  return normalizeFieldsIntoSteps(
    v.fields.map((f) => ({ ...f })),
    ensureCoreSteps(stepsSrc)
  );
}

function normalizeFieldsIntoSteps(
  flds: IFormFieldConfig[],
  stepsIn: IFormStepConfig[]
): { fields: IFormFieldConfig[]; steps: IFormStepConfig[] } {
  const base = ensureCoreSteps(
    stepsIn.map((s) => ({ ...s, fieldNames: s.fieldNames.slice() }))
  );
  const nextSteps = base.map((s) => ({ ...s, fieldNames: [] as string[] }));
  const nextFields = flds.map((f) => ({ ...f }));
  for (let i = 0; i < nextFields.length; i++) {
    const name = nextFields[i].internalName;
    let stepIdx = 0;
    let assigned = false;
    for (let j = 0; j < base.length; j++) {
      if (base[j].fieldNames.indexOf(name) !== -1) {
        stepIdx = j;
        assigned = true;
        break;
      }
    }
    if (!assigned) {
      const sid = nextFields[i].sectionId;
      stepIdx = -1;
      if (sid) {
        for (let k = 0; k < base.length; k++) {
          if (base[k].id === sid) {
            stepIdx = k;
            break;
          }
        }
      }
      if (stepIdx < 0) {
        const oi = base.findIndex((x) => x.id === FORM_OCULTOS_STEP_ID);
        stepIdx = oi >= 0 ? oi : 0;
      }
    }
    nextSteps[stepIdx].fieldNames.push(name);
    nextFields[i].sectionId = nextSteps[stepIdx].id;
  }
  return { fields: nextFields, steps: nextSteps };
}

function buildStepNavigationForSave(
  requireFilled: boolean,
  fullVal: boolean,
  allowBack: boolean
): IFormStepNavigationConfig | undefined {
  const sn: IFormStepNavigationConfig = {};
  if (requireFilled) sn.requireFilledRequiredToAdvance = true;
  if (fullVal) sn.fullValidationOnAdvance = true;
  if (!allowBack) sn.allowBackWithoutValidation = false;
  if (Object.keys(sn).length === 0) return undefined;
  return sn;
}

export interface IFormManagerConfigPanelProps {
  isOpen: boolean;
  listTitle: string;
  value: IFormManagerConfig;
  onSave: (next: IFormManagerConfig) => void;
  onDismiss: () => void;
}

export const FormManagerConfigPanel: React.FC<IFormManagerConfigPanelProps> = ({
  isOpen,
  listTitle,
  value,
  onSave,
  onDismiss,
}) => {
  const [fields, setFields] = useState<IFormFieldConfig[]>(() => buildInitialFieldsAndSteps(value).fields);
  const [rules, setRules] = useState<TFormRule[]>(() => value.rules ?? []);
  const [steps, setSteps] = useState<IFormStepConfig[]>(() => buildInitialFieldsAndSteps(value).steps);
  const [helpJson, setHelpJson] = useState(() => JSON.stringify(value.dynamicHelp ?? [], null, 2));
  const [managerColumnFields, setManagerColumnFields] = useState<string[]>(() => value.managerColumnFields ?? []);
  const [customButtons, setCustomButtons] = useState<IFormCustomButtonConfig[]>(() =>
    (value.customButtons ?? []).map((b) => ({
      ...b,
      actions: b.actions.map((a) => ({ ...a })),
    }))
  );
  const [showDefaultFormButtons, setShowDefaultFormButtons] = useState(() => value.showDefaultFormButtons === true);
  const [stepLayout, setStepLayout] = useState<TFormStepLayoutKind>(() => value.stepLayout ?? 'segmented');
  const [stepNavButtons, setStepNavButtons] = useState<TFormStepNavButtonsKind>(
    () => value.stepNavButtons ?? 'fluent'
  );
  const [formDataLoadingKind, setFormDataLoadingKind] = useState<TFormDataLoadingUiKind>(
    () => value.formDataLoadingKind ?? 'spinner'
  );
  const [defaultSubmitLoadingKind, setDefaultSubmitLoadingKind] = useState<TFormSubmitLoadingUiKind>(
    () => value.defaultSubmitLoadingKind ?? 'overlay'
  );
  const [stepRequireFilledToAdvance, setStepRequireFilledToAdvance] = useState(
    () => value.stepNavigation?.requireFilledRequiredToAdvance === true
  );
  const [stepFullValOnAdvance, setStepFullValOnAdvance] = useState(
    () => value.stepNavigation?.fullValidationOnAdvance === true
  );
  const [stepAllowBackWithoutVal, setStepAllowBackWithoutVal] = useState(
    () => value.stepNavigation?.allowBackWithoutValidation !== false
  );
  const [stepSectionOpen, setStepSectionOpen] = useState<Record<string, boolean>>({});
  const [buttonSectionOpen, setButtonSectionOpen] = useState<Record<string, boolean>>({});
  const [attachMin, setAttachMin] = useState('');
  const [attachMax, setAttachMax] = useState('');
  const [attachMsg, setAttachMsg] = useState('');
  const [meta, setMeta] = useState<IFieldMetadata[]>([]);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState<string | undefined>(undefined);
  const [jsonOpen, setJsonOpen] = useState(false);
  const [fieldPanelName, setFieldPanelName] = useState<string | null>(null);
  const [redirectReplaceBraceForBtnId, setRedirectReplaceBraceForBtnId] = useState<string | null>(null);
  const [redirectInsertNonceByBtn, setRedirectInsertNonceByBtn] = useState<Record<string, number>>({});
  const [redirectReplaceNonceByBtn, setRedirectReplaceNonceByBtn] = useState<Record<string, number>>({});

  const fieldsService = useMemo(() => new FieldsService(), []);

  useEffect(() => {
    if (!isOpen) return;
    const norm = buildInitialFieldsAndSteps(value);
    setFields(norm.fields);
    setSteps(norm.steps);
    setRules(value.rules ?? []);
    setHelpJson(JSON.stringify(value.dynamicHelp ?? [], null, 2));
    setManagerColumnFields(value.managerColumnFields ?? []);
    setCustomButtons(
      (value.customButtons ?? []).map((b) => ({
        ...b,
        actions: b.actions.map((a) => ({ ...a })),
      }))
    );
    setShowDefaultFormButtons(value.showDefaultFormButtons === true);
    setStepLayout(value.stepLayout ?? 'segmented');
    setStepNavButtons(value.stepNavButtons ?? 'fluent');
    setFormDataLoadingKind(value.formDataLoadingKind ?? 'spinner');
    setDefaultSubmitLoadingKind(value.defaultSubmitLoadingKind ?? 'overlay');
    setStepRequireFilledToAdvance(value.stepNavigation?.requireFilledRequiredToAdvance === true);
    setStepFullValOnAdvance(value.stepNavigation?.fullValidationOnAdvance === true);
    setStepAllowBackWithoutVal(value.stepNavigation?.allowBackWithoutValidation !== false);
    const att = parseAttachmentUiRule(value.rules ?? []);
    setAttachMin(att.minCount);
    setAttachMax(att.maxCount);
    setAttachMsg(att.message);
    setErr(undefined);
    setFieldPanelName(null);
    setStepSectionOpen({});
    setButtonSectionOpen({});
  }, [isOpen, value]);

  useEffect(() => {
    if (!isOpen || !listTitle.trim()) return;
    setLoading(true);
    fieldsService
      .getVisibleFields(listTitle.trim())
      .then((f) => {
        setMeta(f);
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, [isOpen, listTitle, fieldsService]);

  const fieldOptions: IDropdownOption[] = useMemo(
    () =>
      meta
        .filter(isFormConfigSelectableField)
        .map((f) => ({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` })),
    [meta]
  );

  const conditionalCards = useMemo(() => parseConditionalCardsFromRules(rules).cards, [rules]);

  const customs = useMemo(() => customRulesOnly(rules), [rules]);

  const requiredFieldsMissingFromSteps = useMemo(
    () => requiredListFieldsMissingFromSteps(meta, steps),
    [meta, steps]
  );

  const anyStepHasFields = useMemo(() => hasAnyFieldInAnyStep(steps), [steps]);

  const metaSortedForPool = useMemo(() => {
    return meta
      .filter(isFormConfigSelectableField)
      .slice()
      .sort((a, b) => {
        if (a.Required !== b.Required) return a.Required ? -1 : 1;
        return a.Title.localeCompare(b.Title, 'pt');
      });
  }, [meta]);

  const redirectDynamicFieldOptions = useMemo((): IDropdownOption[] => {
    const base: IDropdownOption[] = [
      { key: REDIRECT_KEY_FORM, text: '{{Form}} — modo (Display / Edit / New)' },
      { key: REDIRECT_KEY_FORMID, text: '{{FormID}} — id do item na lista' },
    ];
    return base.concat(
      meta.filter(isFormConfigSelectableField).map((m) => ({
        key: m.InternalName,
        text: `${m.Title}  →  {{${m.InternalName}}}`,
      }))
    );
  }, [meta]);

  const setCardsAndRules = useCallback((cards: IConditionalRuleCard[]) => {
    setRules((r) => mergeCardRulesIntoAll(r, cards));
  }, []);

  const addField = (internalName: string): void => {
    if (!internalName) return;
    setSteps((prevSteps) => {
      const st = ensureCoreSteps(prevSteps);
      let already = false;
      for (let s = 0; s < st.length; s++) {
        if (st[s].fieldNames.indexOf(internalName) !== -1) {
          already = true;
          break;
        }
      }
      if (already) return st;
      const oi = st.findIndex((x) => x.id === FORM_OCULTOS_STEP_ID);
      const idx = oi >= 0 ? oi : 0;
      const sid = st[idx].id;
      const nextSteps = st.map((s, i) =>
        i === idx ? { ...s, fieldNames: s.fieldNames.concat([internalName]) } : s
      );
      setFields((prev) => {
        const withF = prev.some((f) => f.internalName === internalName)
          ? prev
          : prev.concat([{ internalName, sectionId: sid }]);
        return fieldsAlignedToSteps(withF, nextSteps);
      });
      return nextSteps;
    });
  };

  const removeField = (internalName: string): void => {
    setErr(undefined);
    if (hasAnyFieldInAnyStep(steps)) {
      for (let mi = 0; mi < meta.length; mi++) {
        if (meta[mi].InternalName === internalName && meta[mi].Required) {
          setErr(
            `O campo «${meta[mi].Title}» é obrigatório na lista e tem de constar em alguma etapa.`
          );
          return;
        }
      }
    }
    setFields((prev) => prev.filter((f) => f.internalName !== internalName));
    setSteps((prev) =>
      prev.map((s) => ({
        ...s,
        fieldNames: s.fieldNames.filter((n) => n !== internalName),
      }))
    );
  };

  const reorderField = (from: number, to: number): void => {
    setFields((prev) => {
      const next = reorderByIndex(prev, from, to);
      setSteps((st) => resyncStepsOrderFromFields(next, st));
      return next;
    });
  };

  const updateFieldAt = (internalName: string, patch: Partial<IFormFieldConfig>): void => {
    setFields((prev) => prev.map((f) => (f.internalName === internalName ? { ...f, ...patch } : f)));
  };

  const handleStructureFieldDrop = useCallback((toStepIdx: number, insertBefore: number) => {
    return (e: React.DragEvent<HTMLElement>): void => {
      e.preventDefault();
      e.stopPropagation();
      const d = e.dataTransfer.getData('text/plain');
      const poolName = parsePoolDrag(d);
      if (poolName) {
        setSteps((prevSteps) => {
          const nextSteps = insertFieldNameIntoStep(prevSteps, poolName, toStepIdx, insertBefore);
          setFields((prevFields) => {
            let f = prevFields;
            const sid = nextSteps[toStepIdx] ? nextSteps[toStepIdx].id : '';
            let has = false;
            for (let i = 0; i < f.length; i++) {
              if (f[i].internalName === poolName) {
                has = true;
                break;
              }
            }
            if (!has) {
              f = f.concat([{ internalName: poolName, sectionId: sid }]);
            }
            return fieldsAlignedToSteps(f, nextSteps);
          });
          return nextSteps;
        });
        return;
      }
      const fs = parseFieldInStepDrag(d);
      if (fs) {
        setSteps((prevSteps) => {
          const nextSteps = insertFieldNameIntoStep(prevSteps, fs.name, toStepIdx, insertBefore);
          setFields((prevFields) => fieldsAlignedToSteps(prevFields, nextSteps));
          return nextSteps;
        });
      }
    };
  }, []);

  const handleSave = (): void => {
    setErr(undefined);
    let dynamicHelp: IFormManagerConfig['dynamicHelp'];
    try {
      const h = JSON.parse(helpJson || '[]');
      dynamicHelp = Array.isArray(h) && h.length > 0 ? h : undefined;
    } catch {
      setErr('JSON de ajuda dinâmica inválido.');
      return;
    }
    if (meta.length > 0) {
      const missingReq = requiredListFieldsMissingFromSteps(meta, steps);
      if (missingReq.length > 0) {
        setErr(
          'Campos obrigatórios na lista têm de constar em alguma etapa: ' +
            missingReq.map((f) => `${f.Title} (${f.InternalName})`).join(', ')
        );
        return;
      }
    }
    const withRules = mergeAttachmentUiRule(rules, {
      minCount: numOpt(attachMin),
      maxCount: numOpt(attachMax),
      message: attachMsg,
    });
    const sectionsOut = sectionsFromSteps(steps);
    const stepNavigation = buildStepNavigationForSave(
      stepRequireFilledToAdvance,
      stepFullValOnAdvance,
      stepAllowBackWithoutVal
    );
    const raw: IFormManagerConfig = {
      sections: sectionsOut,
      fields,
      rules: withRules,
      steps,
      ...(dynamicHelp ? { dynamicHelp } : {}),
      ...(managerColumnFields.length ? { managerColumnFields } : {}),
      ...(customButtons.length ? { customButtons } : {}),
      stepLayout,
      ...(stepNavButtons && stepNavButtons !== 'fluent' ? { stepNavButtons } : {}),
      ...(formDataLoadingKind && formDataLoadingKind !== 'spinner' ? { formDataLoadingKind } : {}),
      ...(defaultSubmitLoadingKind && defaultSubmitLoadingKind !== 'overlay'
        ? { defaultSubmitLoadingKind }
        : {}),
      ...(showDefaultFormButtons ? { showDefaultFormButtons: true } : {}),
      ...(stepNavigation ? { stepNavigation } : {}),
    };
    const sanitized = sanitizeFormManagerConfig(raw);
    if (!sanitized) {
      setErr('Configuração inválida.');
      return;
    }
    onSave(sanitized);
    onDismiss();
  };

  const addStep = (): void => {
    setSteps((prev) => [...prev, { id: newId('step'), title: 'Nova etapa', fieldNames: [] }]);
  };

  const updateStep = (i: number, patch: Partial<IFormStepConfig>): void => {
    setSteps((prev) => prev.map((s, j) => (j === i ? { ...s, ...patch } : s)));
  };

  const reorderStep = (from: number, to: number): void => {
    setSteps((prev) => {
      const n = reorderByIndex(prev, from, to);
      setFields((flds) => fieldsAlignedToSteps(flds, n));
      return n;
    });
  };

  const removeStep = (i: number): void => {
    setSteps((prev) => {
      if (prev.length <= 1) return prev;
      const removed = prev[i];
      if (!removed) return prev;
      if (removed.id === FORM_OCULTOS_STEP_ID) return prev;
      const next = prev.filter((_, j) => j !== i);
      const t0 = next[0];
      if (!t0) return prev;
      const merged = t0.fieldNames.slice();
      for (let k = 0; k < removed.fieldNames.length; k++) {
        const n = removed.fieldNames[k];
        if (merged.indexOf(n) === -1) merged.push(n);
      }
      next[0] = { ...t0, fieldNames: merged };
      setFields((pf) =>
        fieldsAlignedToSteps(
          pf.map((f) =>
            removed.fieldNames.indexOf(f.internalName) !== -1
              ? { ...f, sectionId: next[0].id }
              : f
          ),
          next
        )
      );
      return ensureCoreSteps(next);
    });
  };

  const toggleManagerCol = (internalName: string, checked: boolean): void => {
    setManagerColumnFields((prev) => {
      if (checked) {
        if (prev.indexOf(internalName) !== -1) return prev;
        return prev.concat([internalName]);
      }
      return prev.filter((x) => x !== internalName);
    });
  };

  const reorderManagerCol = (from: number, to: number): void => {
    setManagerColumnFields((prev) => reorderByIndex(prev, from, to));
  };

  const addCustomButton = (): void => {
    setCustomButtons((b) =>
      b.concat([
        {
          id: newId('btn'),
          label: 'Novo botão',
          appearance: 'default',
          operation: 'legacy',
          behavior: 'actionsOnly',
          actions: [],
        },
      ])
    );
  };

  const patchCustomButton = (i: number, patch: Partial<IFormCustomButtonConfig>): void => {
    setCustomButtons((prev) => prev.map((x, j) => (j === i ? { ...x, ...patch } : x)));
  };

  const patchButtonWhenUi = (bi: number, partial: Partial<IWhenUi>): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) => {
        if (j !== bi) return b;
        const baseLeaf = b.when ? whenNodeToUi(b.when) : undefined;
        const base: IWhenUi = baseLeaf ?? defaultWhenUi(meta);
        const merged: IWhenUi = { ...base, ...partial };
        return { ...b, when: whenUiToNode(merged) };
      })
    );
  };

  const patchButtonActionWhenUi = (bi: number, ai: number, partial: Partial<IWhenUi>): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) => {
        if (j !== bi) return b;
        const acts = b.actions.map((a, k) => {
          if (k !== ai) return a;
          const baseLeaf = a.when ? whenNodeToUi(a.when) : undefined;
          const base: IWhenUi = baseLeaf ?? defaultWhenUi(meta);
          const merged: IWhenUi = { ...base, ...partial };
          return { ...a, when: whenUiToNode(merged) } as TFormButtonAction;
        });
        return { ...b, actions: acts };
      })
    );
  };

  const removeCustomButton = (i: number): void => {
    setCustomButtons((prev) => prev.filter((_, j) => j !== i));
  };

  const reorderCustomButton = (from: number, to: number): void => {
    setCustomButtons((prev) => reorderByIndex(prev, from, to));
  };

  const addButtonAction = (bi: number): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) =>
        j === bi ? { ...b, actions: b.actions.concat([{ kind: 'showFields', fields: [] }]) } : b
      )
    );
  };

  const patchButtonAction = (bi: number, ai: number, next: TFormButtonAction): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) => {
        if (j !== bi) return b;
        const acts = b.actions.map((a, k) => (k === ai ? next : a));
        return { ...b, actions: acts };
      })
    );
  };

  const removeButtonAction = (bi: number, ai: number): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) =>
        j === bi ? { ...b, actions: b.actions.filter((_, k) => k !== ai) } : b
      )
    );
  };

  const setButtonModesFromTriState = (bi: number, c: boolean, e: boolean, v: boolean): void => {
    patchCustomButton(bi, { modes: modesFromCheckboxes(c, e, v) });
  };

  let fieldPanelConfig: IFormFieldConfig | undefined;
  let fieldPanelMeta: IFieldMetadata | undefined;
  if (fieldPanelName) {
    for (let i = 0; i < fields.length; i++) {
      if (fields[i].internalName === fieldPanelName) {
        fieldPanelConfig = fields[i];
        break;
      }
    }
    for (let j = 0; j < meta.length; j++) {
      if (meta[j].InternalName === fieldPanelName) {
        fieldPanelMeta = meta[j];
        break;
      }
    }
  }

  const previewConfigJson = useMemo(() => {
    const withRules = mergeAttachmentUiRule(rules, {
      minCount: numOpt(attachMin),
      maxCount: numOpt(attachMax),
      message: attachMsg,
    });
    let dynamicHelp: IFormManagerConfig['dynamicHelp'];
    try {
      const h = JSON.parse(helpJson || '[]');
      dynamicHelp = Array.isArray(h) && h.length > 0 ? h : undefined;
    } catch {
      dynamicHelp = undefined;
    }
    const sectionsOut = sectionsFromSteps(steps);
    const stepNavigation = buildStepNavigationForSave(
      stepRequireFilledToAdvance,
      stepFullValOnAdvance,
      stepAllowBackWithoutVal
    );
    const raw: IFormManagerConfig = {
      sections: sectionsOut,
      fields,
      rules: withRules,
      steps,
      ...(dynamicHelp ? { dynamicHelp } : {}),
      ...(managerColumnFields.length ? { managerColumnFields } : {}),
      ...(customButtons.length ? { customButtons } : {}),
      stepLayout,
      ...(stepNavButtons && stepNavButtons !== 'fluent' ? { stepNavButtons } : {}),
      ...(formDataLoadingKind && formDataLoadingKind !== 'spinner' ? { formDataLoadingKind } : {}),
      ...(defaultSubmitLoadingKind && defaultSubmitLoadingKind !== 'overlay'
        ? { defaultSubmitLoadingKind }
        : {}),
      ...(showDefaultFormButtons ? { showDefaultFormButtons: true } : {}),
      ...(stepNavigation ? { stepNavigation } : {}),
    };
    return JSON.stringify(raw, null, 2);
  }, [
    fields,
    rules,
    steps,
    helpJson,
    managerColumnFields,
    customButtons,
    stepLayout,
    stepNavButtons,
    formDataLoadingKind,
    defaultSubmitLoadingKind,
    showDefaultFormButtons,
    stepRequireFilledToAdvance,
    stepFullValOnAdvance,
    stepAllowBackWithoutVal,
    attachMin,
    attachMax,
    attachMsg,
  ]);

  const addConditionalCard = (): void => {
    const card: IConditionalRuleCard = {
      id: newCardId(),
      when: defaultWhenUi(meta),
      effects: [emptyEffect()],
    };
    setCardsAndRules(conditionalCards.concat([card]));
  };

  const patchCard = (index: number, patch: Partial<IConditionalRuleCard>): void => {
    const next = conditionalCards.map((c, i) => (i === index ? { ...c, ...patch } : c));
    setCardsAndRules(next);
  };

  const patchWhen = (index: number, w: Partial<IWhenUi>): void => {
    const c = conditionalCards[index];
    if (!c) return;
    patchCard(index, { when: { ...c.when, ...w } });
  };

  const patchEffect = (cardIndex: number, effIndex: number, patch: Partial<IConditionalEffectUi>): void => {
    const c = conditionalCards[cardIndex];
    if (!c) return;
    const effects = c.effects.map((e, i) => (i === effIndex ? { ...e, ...patch } : e));
    patchCard(cardIndex, { effects });
  };

  const addEffect = (cardIndex: number): void => {
    const c = conditionalCards[cardIndex];
    if (!c) return;
    patchCard(cardIndex, { effects: c.effects.concat([emptyEffect()]) });
  };

  const removeEffect = (cardIndex: number, effIndex: number): void => {
    const c = conditionalCards[cardIndex];
    if (!c) return;
    patchCard(cardIndex, { effects: c.effects.filter((_, i) => i !== effIndex) });
  };

  const duplicateCard = (index: number): void => {
    const c = conditionalCards[index];
    if (!c) return;
    const copy: IConditionalRuleCard = {
      ...c,
      id: newCardId(),
      effects: c.effects.map((e) => ({ ...e })),
    };
    const next = conditionalCards.slice();
    next.splice(index + 1, 0, copy);
    setCardsAndRules(next);
  };

  const removeCard = (index: number): void => {
    setCardsAndRules(conditionalCards.filter((_, i) => i !== index));
  };

  const applyPresetConditional = (preset: 'showWhenEq' | 'choiceRequire'): void => {
    const a = meta[0]?.InternalName ?? 'A';
    const b = meta[1]?.InternalName ?? 'B';
    if (preset === 'showWhenEq') {
      const card = templateConditionalShowWhenEquals(a, '', b);
      card.when.compareValue = '';
      setCardsAndRules(conditionalCards.concat([card]));
    } else {
      setCardsAndRules(conditionalCards.concat([templateFieldRulesChoiceRequiresOther(a, '', b)]));
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.large}
      headerText="Configurar formulário e regras"
      onDismiss={onDismiss}
    >
      {loading && <Spinner label="Campos da lista..." />}
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Link onClick={() => setJsonOpen(true)}>Ver JSON gerado</Link>
      </Stack>
      <Pivot>
        <PivotItem headerText="Estrutura">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            {requiredFieldsMissingFromSteps.length > 0 && (
              <MessageBar messageBarType={MessageBarType.warning}>
                Campos marcados como obrigatórios na lista ainda não estão em nenhuma etapa:{' '}
                {requiredFieldsMissingFromSteps.map((f) => `${f.Title} (${f.InternalName})`).join(', ')}
              </MessageBar>
            )}
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Arraste campos para dentro de cada etapa e reordene-os pela alça. O id da etapa é gravado como seção no JSON. Reordene etapas pela alça no cabeçalho. Obrigatórios na lista: verde só quando incluídos no formulário (marcados); com campos nas etapas, têm de estar numa etapa. Layout do passador e botões anterior/próximo: aba Componentes.
            </Text>
            <Stack
              tokens={{ childrenGap: 10 }}
              styles={{ root: { border: '1px solid #edebe9', padding: 12, borderRadius: 4 } }}
            >
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                Navegação entre etapas (formulário)
              </Text>
              <Toggle
                label="Exigir obrigatórios preenchidos para avançar (Próximo / etapa à frente)"
                checked={stepRequireFilledToAdvance}
                onChange={(_, c) => setStepRequireFilledToAdvance(!!c)}
              />
              <Toggle
                label="Ao avançar, aplicar todas as regras de validação nos campos da etapa (não só obrigatório)"
                checked={stepFullValOnAdvance}
                onChange={(_, c) => setStepFullValOnAdvance(!!c)}
                disabled={!stepRequireFilledToAdvance}
              />
              <Toggle
                label="Permitir voltar etapa sem validar a atual"
                checked={stepAllowBackWithoutVal}
                onChange={(_, c) => setStepAllowBackWithoutVal(!!c)}
                disabled={!stepRequireFilledToAdvance}
              />
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Ideias futuras: etapa de revisão antes de enviar; desativar salto direto no passador; barra de progresso por etapa.
              </Text>
            </Stack>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Nova etapa" onClick={addStep} />
            </Stack>
            {steps.map((st, si) => {
              const panelOpen = stepSectionOpen[st.id] === true;
              return (
              <Stack
                key={st.id}
                styles={{ root: { border: '1px solid #edebe9', padding: 12, borderRadius: 4 } }}
                tokens={{ childrenGap: 8 }}
              >
                <Stack
                  horizontal
                  verticalAlign="end"
                  tokens={{ childrenGap: 8 }}
                  wrap
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    e.dataTransfer.dropEffect = 'move';
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    const from = parseDragIndex(e.dataTransfer.getData('text/plain'), DND_STEP);
                    if (from === undefined || from === si) return;
                    reorderStep(from, si);
                  }}
                >
                  <IconButton
                    iconProps={{ iconName: panelOpen ? 'ChevronDown' : 'ChevronRight' }}
                    title={panelOpen ? 'Recolher' : 'Expandir'}
                    aria-expanded={panelOpen}
                    onClick={() => {
                      setStepSectionOpen((p) => ({
                        ...p,
                        [st.id]: !panelOpen,
                      }));
                    }}
                  />
                  <span
                    draggable
                    title="Arrastar etapa"
                    onDragStart={(e) => {
                      e.dataTransfer.setData('text/plain', dragPayload(DND_STEP, si));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                  >
                    <Icon iconName="GripperBarVertical" />
                  </span>
                  <TextField
                    label={`Título da etapa (${st.id})`}
                    value={st.title}
                    onChange={(_, v) => updateStep(si, { title: v ?? '' })}
                  />
                  {!panelOpen && (
                    <Text variant="small" styles={{ root: { color: '#605e5c', alignSelf: 'center' } }}>
                      {st.fieldNames.length} campo(s)
                    </Text>
                  )}
                  <DefaultButton text="Remover etapa" onClick={() => removeStep(si)} />
                </Stack>
                {panelOpen && (
                <Stack tokens={{ childrenGap: 6 }} styles={{ root: { marginTop: 4 } }}>
                  {st.fieldNames.map((fname, fIdx) => {
                    let mm: IFieldMetadata | undefined;
                    for (let mi = 0; mi < meta.length; mi++) {
                      if (meta[mi].InternalName === fname) {
                        mm = meta[mi];
                        break;
                      }
                    }
                    const reqStyles = requiredFieldRowStyles(mm, steps, fields);
                    return (
                      <Stack
                        key={fname}
                        horizontal
                        verticalAlign="center"
                        tokens={{ childrenGap: 8 }}
                        wrap
                        styles={{
                          root: {
                            padding: '8px 10px',
                            borderRadius: 4,
                            ...(reqStyles ?? { background: '#faf9f8', border: '1px solid #edebe9' }),
                          },
                        }}
                        onDragOver={(e) => {
                          e.preventDefault();
                          e.stopPropagation();
                          e.dataTransfer.dropEffect = 'move';
                        }}
                        onDrop={handleStructureFieldDrop(si, fIdx)}
                      >
                        <span
                          draggable
                          title="Arrastar campo"
                          onDragStart={(e) => {
                            e.dataTransfer.setData('text/plain', dragPayloadFieldInStep(si, fIdx, fname));
                            e.dataTransfer.effectAllowed = 'move';
                          }}
                          style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                        >
                          <Icon iconName="GripperBarVertical" />
                        </span>
                        <Text styles={{ root: { fontWeight: 600, minWidth: 120 } }}>
                          {mm ? mm.Title : fname === FORM_ATTACHMENTS_FIELD_INTERNAL ? 'Anexos ao item' : fname}
                        </Text>
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                          {fname} · {fname === FORM_ATTACHMENTS_FIELD_INTERNAL ? 'anexos' : mm ? mm.MappedType : '—'}
                          {mm?.Required ? ' · obrigatório na lista' : ''}
                        </Text>
                        {fname !== FORM_ATTACHMENTS_FIELD_INTERNAL && (
                          <DefaultButton text="Regras…" onClick={() => setFieldPanelName(fname)} />
                        )}
                        <DefaultButton
                          text="Remover"
                          onClick={() => removeField(fname)}
                          disabled={anyStepHasFields && mm?.Required === true}
                          title={
                            anyStepHasFields && mm?.Required === true
                              ? 'Obrigatório na lista: tem de permanecer numa etapa'
                              : undefined
                          }
                        />
                      </Stack>
                    );
                  })}
                  <Stack
                    styles={{
                      root: {
                        minHeight: 40,
                        padding: 8,
                        borderRadius: 4,
                        border: '1px dashed #c8c6c4',
                        background: '#ffffff',
                      },
                    }}
                    onDragOver={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      e.dataTransfer.dropEffect = 'move';
                    }}
                    onDrop={handleStructureFieldDrop(si, st.fieldNames.length)}
                  >
                    <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                      Soltar aqui para colocar no fim desta etapa
                    </Text>
                  </Stack>
                </Stack>
                )}
              </Stack>
              );
            })}
            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>Campos fora do formulário</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Arraste um campo para uma etapa acima ou marque para incluir na primeira etapa.
            </Text>
            {(() => {
              let attInPool = false;
              for (let i = 0; i < fields.length; i++) {
                if (fields[i].internalName === FORM_ATTACHMENTS_FIELD_INTERNAL) {
                  attInPool = true;
                  break;
                }
              }
              if (attInPool) return null;
              return (
                <Stack
                  key={FORM_ATTACHMENTS_FIELD_INTERNAL}
                  horizontal
                  verticalAlign="center"
                  tokens={{ childrenGap: 8 }}
                  wrap
                >
                  <span
                    draggable
                    title="Arrastar para uma etapa"
                    onDragStart={(e) => {
                      e.dataTransfer.setData('text/plain', dragPayloadPool(FORM_ATTACHMENTS_FIELD_INTERNAL));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                  >
                    <Icon iconName="GripperBarVertical" />
                  </span>
                  <Checkbox
                    label="Anexos ao item (ficheiros)"
                    checked={false}
                    onChange={(_, c) => (c ? addField(FORM_ATTACHMENTS_FIELD_INTERNAL) : undefined)}
                  />
                  <Text variant="small" styles={{ root: { minWidth: 80 } }}>
                    anexos
                  </Text>
                </Stack>
              );
            })()}
            {metaSortedForPool.map((m) => {
              let inForm = false;
              for (let i = 0; i < fields.length; i++) {
                if (fields[i].internalName === m.InternalName) {
                  inForm = true;
                  break;
                }
              }
              if (inForm) return null;
              const poolReqStyles = requiredFieldRowStyles(m, steps, fields);
              return (
                <Stack
                  key={m.InternalName}
                  horizontal
                  verticalAlign="center"
                  tokens={{ childrenGap: 8 }}
                  wrap
                  styles={{
                    root: poolReqStyles
                      ? { padding: '8px 10px', borderRadius: 4, ...poolReqStyles }
                      : undefined,
                  }}
                >
                  <span
                    draggable
                    title="Arrastar para uma etapa"
                    onDragStart={(e) => {
                      e.dataTransfer.setData('text/plain', dragPayloadPool(m.InternalName));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                  >
                    <Icon iconName="GripperBarVertical" />
                  </span>
                  <Checkbox
                    label={`${m.Title} (${m.InternalName})${m.Required ? ' *' : ''}`}
                    checked={false}
                    onChange={(_, c) => (c ? addField(m.InternalName) : undefined)}
                  />
                  <Text variant="small" styles={{ root: { minWidth: 80 } }}>
                    {m.MappedType}
                    {m.Required ? ' · obrig. lista' : ''}
                  </Text>
                </Stack>
              );
            })}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Componentes">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <FormManagerComponentsTabContent
              loading={loading}
              stepLayout={stepLayout}
              onStepLayoutChange={setStepLayout}
              stepNavButtons={stepNavButtons}
              onStepNavButtonsChange={setStepNavButtons}
              formDataLoadingKind={formDataLoadingKind}
              onFormDataLoadingKindChange={setFormDataLoadingKind}
              defaultSubmitLoadingKind={defaultSubmitLoadingKind}
              onDefaultSubmitLoadingKindChange={setDefaultSubmitLoadingKind}
            />
          </Stack>
        </PivotItem>
        <PivotItem headerText="Botões">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Botões extra no rodapé do formulário. Ao clicar, as ações executam por ordem (mostrar/ocultar campos,
              preencher valores). Para texto composto a partir de campos, use o prefixo str: e placeholders no formato
              {' {{NomeInterno}} '} (igual à expressão de texto da regra de valor calculado). Visibilidade por grupo e
              por condição em dados (ex.: coluna Status) usa os campos abaixo em cada botão. Condições compostas
              (várias cláusulas) só em JSON avançado.
            </Text>
            <Checkbox
              label="Mostrar também os botões padrão (Enviar, Rascunho, Fechar)"
              checked={showDefaultFormButtons}
              onChange={(_, c) => setShowDefaultFormButtons(!!c)}
            />
            <PrimaryButton text="Adicionar botão" onClick={addCustomButton} />
            {customButtons.map((btn, bi) => {
              const chk = checkboxesFromModes(btn.modes);
              const leafWhen = btn.when ? whenNodeToUi(btn.when) : undefined;
              const panelOpen = buttonSectionOpen[btn.id] === true;
              return (
                <Stack
                  key={btn.id}
                  styles={{ root: { border: '1px solid #edebe9', padding: 12, borderRadius: 4 } }}
                  tokens={{ childrenGap: 10 }}
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    const from = parseDragIndex(e.dataTransfer.getData('text/plain'), DND_BTN);
                    if (from === undefined || from === bi) return;
                    reorderCustomButton(from, bi);
                  }}
                >
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                      <IconButton
                        iconProps={{ iconName: panelOpen ? 'ChevronDown' : 'ChevronRight' }}
                        title={panelOpen ? 'Recolher' : 'Expandir'}
                        aria-expanded={panelOpen}
                        onClick={() => {
                          setButtonSectionOpen((p) => ({
                            ...p,
                            [btn.id]: !panelOpen,
                          }));
                        }}
                      />
                      <span
                        draggable
                        title="Arrastar para reordenar"
                        onDragStart={(e) => {
                          e.dataTransfer.setData('text/plain', dragPayload(DND_BTN, bi));
                          e.dataTransfer.effectAllowed = 'move';
                        }}
                        style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                      >
                        <Icon iconName="GripperBarVertical" />
                      </span>
                      <Text styles={{ root: { fontWeight: 600 } }}>{btn.label || btn.id}</Text>
                    </Stack>
                    <DefaultButton text="Remover botão" onClick={() => removeCustomButton(bi)} />
                  </Stack>
                  {panelOpen && (
                  <>
                  <TextField
                    label="Texto do botão"
                    value={btn.label}
                    onChange={(_, v) => patchCustomButton(bi, { label: v ?? '' })}
                  />
                  <Dropdown
                    label="Tipo de operação"
                    options={BUTTON_OPERATION_OPTIONS}
                    selectedKey={(btn.operation ?? 'legacy') as string}
                    onChange={(_, o) => {
                      if (!o) return;
                      const k = String(o.key) as TFormCustomButtonOperation;
                      patchCustomButton(bi, {
                        operation: k,
                        ...(k === 'redirect'
                          ? { redirectUrlTemplate: btn.redirectUrlTemplate ?? '', actions: [] }
                          : {}),
                      });
                    }}
                  />
                  <Dropdown
                    label="Loading ao gravar"
                    options={[
                      { key: FORM_SUBMIT_LOADING_INHERIT_KEY, text: 'Padrão (aba Componentes)' },
                      ...FORM_SUBMIT_LOADING_DROPDOWN_OPTIONS,
                    ]}
                    selectedKey={btn.submitLoadingKind ?? FORM_SUBMIT_LOADING_INHERIT_KEY}
                    onChange={(_, o) => {
                      if (!o) return;
                      const k = String(o.key);
                      setCustomButtons((prev) =>
                        prev.map((b, j) => {
                          if (j !== bi) return b;
                          if (k === FORM_SUBMIT_LOADING_INHERIT_KEY) {
                            const { submitLoadingKind: _rm, ...rest } = b;
                            return rest;
                          }
                          return { ...b, submitLoadingKind: k as TFormSubmitLoadingUiKind };
                        })
                      );
                    }}
                  />
                  {(btn.operation ?? 'legacy') === 'redirect' && (
                    <Stack tokens={{ childrenGap: 10 }}>
                      <TextField
                        label="URL de destino"
                        description="Escreva o endereço. Use {{}} vazio para escolher um campo na lista abaixo, ou o menu «Inserir valor dinâmico»."
                        multiline
                        rows={3}
                        value={btn.redirectUrlTemplate ?? ''}
                        onChange={(_, v) => {
                          const next = v ?? '';
                          patchCustomButton(bi, { redirectUrlTemplate: next });
                          if (/\{\{\s*\}\}/.test(next)) {
                            setRedirectReplaceBraceForBtnId(btn.id);
                          } else if (redirectReplaceBraceForBtnId === btn.id) {
                            setRedirectReplaceBraceForBtnId(null);
                          }
                        }}
                      />
                      <Dropdown
                        key={`redirect-ins-${btn.id}-${redirectInsertNonceByBtn[btn.id] ?? 0}`}
                        label="Inserir valor dinâmico (no fim do URL)"
                        options={[{ key: '', text: '— escolher campo —' }, ...redirectDynamicFieldOptions]}
                        selectedKey=""
                        onChange={(_, o) => {
                          if (!o || o.key === '') return;
                          const tok = redirectTokenForKey(String(o.key));
                          patchCustomButton(bi, {
                            redirectUrlTemplate: (btn.redirectUrlTemplate ?? '') + tok,
                          });
                          setRedirectInsertNonceByBtn((p) => ({
                            ...p,
                            [btn.id]: (p[btn.id] ?? 0) + 1,
                          }));
                        }}
                      />
                      {redirectReplaceBraceForBtnId === btn.id && (
                        <Stack
                          tokens={{ childrenGap: 8 }}
                          styles={{
                            root: {
                              padding: 12,
                              background: '#f3f9ff',
                              borderRadius: 4,
                              border: '1px solid #0078d4',
                            },
                          }}
                        >
                          <Text variant="small" styles={{ root: { fontWeight: 600, color: '#0078d4' } }}>
                            Placeholder {'{{}}'} detetado — escolha o valor dinâmico (substitui o primeiro {'{{}}'} vazio):
                          </Text>
                          <Dropdown
                            key={`redirect-repl-${btn.id}-${redirectReplaceNonceByBtn[btn.id] ?? 0}`}
                            label="Campo ou token"
                            options={[{ key: '', text: '— selecionar —' }, ...redirectDynamicFieldOptions]}
                            selectedKey=""
                            onChange={(_, o) => {
                              if (!o || o.key === '') return;
                              const cur = btn.redirectUrlTemplate ?? '';
                              const next = replaceFirstEmptyRedirectBrace(cur, String(o.key));
                              patchCustomButton(bi, { redirectUrlTemplate: next });
                              setRedirectReplaceNonceByBtn((p) => ({
                                ...p,
                                [btn.id]: (p[btn.id] ?? 0) + 1,
                              }));
                              if (!/\{\{\s*\}\}/.test(next)) {
                                setRedirectReplaceBraceForBtnId(null);
                              }
                            }}
                          />
                        </Stack>
                      )}
                    </Stack>
                  )}
                  {(btn.operation ?? 'legacy') === 'add' && (
                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                      Cria um novo item na lista com os valores atuais do formulário (validação igual a «Enviar»).
                    </Text>
                  )}
                  {(btn.operation ?? 'legacy') === 'update' && (
                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                      Atualiza o item aberto. O sistema usa o contexto da página (ex.: FormID na query). Modo novo não aplica.
                    </Text>
                  )}
                  {(btn.operation ?? 'legacy') === 'delete' && (
                    <Stack tokens={{ childrenGap: 8 }}>
                      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                        Mostrar o botão eliminar em:
                      </Text>
                      <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                        <Checkbox
                          label="Modo ver (Disp)"
                          checked={btn.deleteShowInView !== false}
                          onChange={(_, c) => patchCustomButton(bi, { deleteShowInView: !!c })}
                        />
                        <Checkbox
                          label="Modo editar"
                          checked={btn.deleteShowInEdit !== false}
                          onChange={(_, c) => patchCustomButton(bi, { deleteShowInEdit: !!c })}
                        />
                      </Stack>
                    </Stack>
                  )}
                  <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="end">
                    <Dropdown
                      label="Estilo"
                      options={[
                        { key: 'default', text: 'Secundário' },
                        { key: 'primary', text: 'Primário' },
                      ]}
                      selectedKey={btn.appearance === 'primary' ? 'primary' : 'default'}
                      onChange={(_, o) =>
                        o && patchCustomButton(bi, { appearance: o.key === 'primary' ? 'primary' : 'default' })
                      }
                    />
                    {(btn.operation ?? 'legacy') === 'legacy' && (
                      <Dropdown
                        label="Depois das ações"
                        options={BUTTON_BEHAVIOR_OPTIONS}
                        selectedKey={(btn.behavior ?? 'actionsOnly') as string}
                        onChange={(_, o) =>
                          o &&
                          patchCustomButton(bi, {
                            behavior: String(o.key) as TFormCustomButtonBehavior,
                          })
                        }
                      />
                    )}
                  </Stack>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Modos (vazio = todos)
                  </Text>
                  <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                    <Checkbox
                      label="Criar"
                      checked={chk.c}
                      onChange={(_, c) => setButtonModesFromTriState(bi, !!c, chk.e, chk.v)}
                    />
                    <Checkbox
                      label="Editar"
                      checked={chk.e}
                      onChange={(_, c) => setButtonModesFromTriState(bi, chk.c, !!c, chk.v)}
                    />
                    <Checkbox
                      label="Ver"
                      checked={chk.v}
                      onChange={(_, c) => setButtonModesFromTriState(bi, chk.c, chk.e, !!c)}
                    />
                  </Stack>
                  <Checkbox
                    label="Botão ativo"
                    checked={btn.enabled !== false}
                    onChange={(_, c) => patchCustomButton(bi, { enabled: c ? undefined : false })}
                  />
                  <TextField
                    label="Grupos do SharePoint (títulos, vírgula)"
                    description="Vazio = qualquer utilizador. Igual às regras: compara o título do grupo, sem diferenciar maiúsculas."
                    value={fieldNamesToCsv(btn.groupTitles ?? [])}
                    onChange={(_, v) => {
                      const parsed = parseCsvFieldNames(v ?? '');
                      patchCustomButton(bi, { groupTitles: parsed.length ? parsed : undefined });
                    }}
                  />
                  <Checkbox
                    label="Mostrar só quando a condição abaixo for verdadeira"
                    checked={!!btn.when}
                    onChange={(_, c) => {
                      if (c) patchCustomButton(bi, { when: whenUiToNode(defaultWhenUi(meta)) });
                      else patchCustomButton(bi, { when: undefined });
                    }}
                  />
                  {btn.when && !leafWhen && (
                    <MessageBar messageBarType={MessageBarType.warning}>
                      Condição composta (várias cláusulas). Edição completa: JSON avançado. Desmarque a caixa acima
                      para remover a condição.
                    </MessageBar>
                  )}
                  {btn.when && leafWhen && (
                    <Stack tokens={{ childrenGap: 8 }}>
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        Condição nos dados do formulário
                      </Text>
                      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                        <Dropdown
                          label="Campo"
                          options={fieldOptions}
                          selectedKey={leafWhen.field}
                          onChange={(_, o) => o && patchButtonWhenUi(bi, { field: String(o.key) })}
                        />
                        <Dropdown
                          label="Operador"
                          options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                          selectedKey={leafWhen.op}
                          onChange={(_, o) => o && patchButtonWhenUi(bi, { op: o.key as TFormConditionOp })}
                        />
                        <Dropdown
                          label="Comparar com"
                          options={[
                            { key: 'literal', text: 'Texto fixo' },
                            { key: 'field', text: 'Outro campo' },
                            { key: 'token', text: 'Token' },
                          ]}
                          selectedKey={leafWhen.compareKind}
                          onChange={(_, o) =>
                            o && patchButtonWhenUi(bi, { compareKind: o.key as IWhenUi['compareKind'] })
                          }
                        />
                        <TextField
                          label="Valor"
                          value={leafWhen.compareValue}
                          onChange={(_, v) => patchButtonWhenUi(bi, { compareValue: v ?? '' })}
                          disabled={
                            leafWhen.op === 'isEmpty' ||
                            leafWhen.op === 'isFilled' ||
                            leafWhen.op === 'isTrue' ||
                            leafWhen.op === 'isFalse'
                          }
                        />
                      </Stack>
                    </Stack>
                  )}
                  {(btn.operation ?? 'legacy') !== 'redirect' && (
                    <>
                      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Ações (por ordem)</Text>
                      {btn.actions.map((act, ai) => (
                        <Stack
                          key={ai}
                          styles={{ root: { background: '#faf9f8', padding: 8, borderRadius: 4 } }}
                          tokens={{ childrenGap: 8 }}
                        >
                          <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                            <Dropdown
                              label="Tipo"
                              options={BUTTON_ACTION_KIND_OPTIONS}
                              selectedKey={act.kind}
                              onChange={(_, o) => {
                                if (!o) return;
                                patchButtonAction(
                                  bi,
                                  ai,
                                  defaultActionForKind(String(o.key) as TFormButtonAction['kind'])
                                );
                              }}
                            />
                            <DefaultButton text="Remover ação" onClick={() => removeButtonAction(bi, ai)} />
                          </Stack>
                          {(act.kind === 'showFields' || act.kind === 'hideFields') && (
                            <TextField
                              label="Campos (internal name, vírgula)"
                              multiline
                              rows={2}
                              value={fieldNamesToCsv(act.fields)}
                              onChange={(_, v) =>
                                patchButtonAction(bi, ai, {
                                  ...act,
                                  fields: parseCsvFieldNames(v ?? ''),
                                })
                              }
                            />
                          )}
                          {act.kind === 'setFieldValue' && (() => {
                            const choiceVal = buttonSetFieldValueChoiceDropdown(
                              act.field,
                              act.valueTemplate,
                              meta
                            );
                            return (
                              <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                                <Dropdown
                                  label="Campo"
                                  options={[{ key: '', text: '—' }, ...fieldOptions]}
                                  selectedKey={act.field || ''}
                                  onChange={(_, o) =>
                                    patchButtonAction(bi, ai, {
                                      ...act,
                                      field: o ? String(o.key) : '',
                                    })
                                  }
                                />
                                {choiceVal ? (
                                  <Dropdown
                                    label="Valor"
                                    styles={{ root: { minWidth: 280 } }}
                                    options={choiceVal.options}
                                    selectedKey={choiceVal.selectedKey}
                                    onChange={(_, o) =>
                                      patchButtonAction(bi, ai, {
                                        ...act,
                                        valueTemplate: o ? String(o.key) : '',
                                      })
                                    }
                                  />
                                ) : (
                                  <TextField
                                    label="Valor fixo ou str:{{Campo}}"
                                    styles={{ root: { minWidth: 280 } }}
                                    value={act.valueTemplate}
                                    onChange={(_, v) =>
                                      patchButtonAction(bi, ai, { ...act, valueTemplate: v ?? '' })
                                    }
                                  />
                                )}
                              </Stack>
                            );
                          })()}
                          {act.kind === 'joinFields' && (
                            <Stack tokens={{ childrenGap: 8 }}>
                              <Dropdown
                                label="Campo destino"
                                options={[{ key: '', text: '—' }, ...fieldOptions]}
                                selectedKey={act.targetField || ''}
                                onChange={(_, o) =>
                                  patchButtonAction(bi, ai, {
                                    ...act,
                                    targetField: o ? String(o.key) : '',
                                  })
                                }
                              />
                              <TextField
                                label="Separador"
                                value={act.separator}
                                onChange={(_, v) =>
                                  patchButtonAction(bi, ai, { ...act, separator: v ?? ' ' })
                                }
                              />
                              <TextField
                                label="Campos origem (vírgula)"
                                multiline
                                rows={2}
                                value={fieldNamesToCsv(act.sourceFields)}
                                onChange={(_, v) =>
                                  patchButtonAction(bi, ai, {
                                    ...act,
                                    sourceFields: parseCsvFieldNames(v ?? ''),
                                  })
                                }
                              />
                            </Stack>
                          )}
                          <Checkbox
                            label="Só executar esta ação se (avalia valores já alterados pelas ações acima)"
                            checked={!!act.when}
                            onChange={(_, c) => {
                              if (c) {
                                patchButtonAction(bi, ai, {
                                  ...act,
                                  when: whenUiToNode(defaultWhenUi(meta)),
                                });
                              } else {
                                const { when: _rm, ...rest } = act as TFormButtonAction & {
                                  when?: TFormConditionNode;
                                };
                                patchButtonAction(bi, ai, rest as TFormButtonAction);
                              }
                            }}
                          />
                          {act.when &&
                            (() => {
                              const leafActWhen = whenNodeToUi(act.when);
                              return !leafActWhen ? (
                                <MessageBar messageBarType={MessageBarType.warning}>
                                  Condição composta nesta ação: use o JSON do gestor ou uma única condição
                                  simples.
                                </MessageBar>
                              ) : (
                                <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 4 } }}>
                                  <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                                    <Dropdown
                                      label="Campo"
                                      options={fieldOptions}
                                      selectedKey={leafActWhen.field}
                                      onChange={(_, o) =>
                                        o && patchButtonActionWhenUi(bi, ai, { field: String(o.key) })
                                      }
                                    />
                                    <Dropdown
                                      label="Operador"
                                      options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                                      selectedKey={leafActWhen.op}
                                      onChange={(_, o) =>
                                        o && patchButtonActionWhenUi(bi, ai, { op: o.key as TFormConditionOp })
                                      }
                                    />
                                    <Dropdown
                                      label="Comparar com"
                                      options={[
                                        { key: 'literal', text: 'Texto fixo' },
                                        { key: 'field', text: 'Outro campo' },
                                        { key: 'token', text: 'Token' },
                                      ]}
                                      selectedKey={leafActWhen.compareKind}
                                      onChange={(_, o) =>
                                        o &&
                                        patchButtonActionWhenUi(bi, ai, {
                                          compareKind: o.key as IWhenUi['compareKind'],
                                        })
                                      }
                                    />
                                    <TextField
                                      label="Valor"
                                      value={leafActWhen.compareValue}
                                      onChange={(_, v) =>
                                        patchButtonActionWhenUi(bi, ai, { compareValue: v ?? '' })
                                      }
                                      disabled={
                                        leafActWhen.op === 'isEmpty' ||
                                        leafActWhen.op === 'isFilled' ||
                                        leafActWhen.op === 'isTrue' ||
                                        leafActWhen.op === 'isFalse'
                                      }
                                    />
                                  </Stack>
                                </Stack>
                              );
                            })()}
                        </Stack>
                      ))}
                      <DefaultButton text="Adicionar ação" onClick={() => addButtonAction(bi)} />
                    </>
                  )}
                  </>
                  )}
                </Stack>
              );
            })}
            {!customButtons.length && <Text>Nenhum botão personalizado.</Text>}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Regras rápidas">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Ajustes base por campo. Reordene linhas arrastando a alça. Regras geradas pela UI aparecem no motor com prefixo ui_f_.
            </Text>
            {fields.map((fc, fIdx) => {
                let m: IFieldMetadata | undefined;
                for (let mi = 0; mi < meta.length; mi++) {
                  if (meta[mi].InternalName === fc.internalName) {
                    m = meta[mi];
                    break;
                  }
                }
              const n = countFieldUiRules(fc.internalName, rules);
              const def = fieldRuleStateFromRules(fc.internalName, rules).defaultValue;
              return (
                <Stack
                  key={fc.internalName}
                  horizontal
                  tokens={{ childrenGap: 8 }}
                  verticalAlign="end"
                  wrap
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    const from = parseDragIndex(e.dataTransfer.getData('text/plain'), DND_FIELD);
                    if (from === undefined || from === fIdx) return;
                    reorderField(from, fIdx);
                  }}
                >
                  <span
                    draggable
                    title="Arrastar para reordenar"
                    onDragStart={(e) => {
                      e.dataTransfer.setData('text/plain', dragPayload(DND_FIELD, fIdx));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                  >
                    <Icon iconName="GripperBarVertical" />
                  </span>
                  <Text styles={{ root: { minWidth: 140, fontWeight: 600 } }}>{fc.internalName}</Text>
                  <Checkbox
                    label="Visível"
                    checked={fc.visible !== false}
                    onChange={(_, c) => updateFieldAt(fc.internalName, { visible: !!c })}
                  />
                  {fc.internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL && (
                    <Checkbox
                      label="Obrigatório"
                      checked={fc.required === true}
                      onChange={(_, c) => updateFieldAt(fc.internalName, { required: !!c })}
                    />
                  )}
                  <Checkbox
                    label="Só leitura"
                    checked={fc.readOnly === true}
                    onChange={(_, c) => updateFieldAt(fc.internalName, { readOnly: !!c })}
                  />
                  <Checkbox
                    label="Desativado"
                    checked={fc.disabled === true}
                    onChange={(_, c) => updateFieldAt(fc.internalName, { disabled: !!c })}
                  />
                  <TextField
                    label="Ajuda"
                    value={fc.helpText ?? ''}
                    onChange={(_, v) => updateFieldAt(fc.internalName, { helpText: v || undefined })}
                  />
                  {fc.internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL && (
                    <TextField
                      label="Padrão (texto/token)"
                      value={def}
                      onChange={(_, v) => {
                        const st = fieldRuleStateFromRules(fc.internalName, rules);
                        st.defaultValue = v ?? '';
                        setRules((r) => mergeFieldRules(r, fc.internalName, buildFieldUiRules(fc.internalName, st)));
                      }}
                    />
                  )}
                  {fc.internalName === FORM_ATTACHMENTS_FIELD_INTERNAL ? (
                    <Text variant="small" styles={{ root: { color: '#605e5c', maxWidth: 280 } }}>
                      Obrigatoriedade e limites de ficheiros: aba Gestor (anexos).
                    </Text>
                  ) : (
                    <DefaultButton
                      text={n ? `${n} regra(s)` : 'Regras…'}
                      onClick={() => setFieldPanelName(fc.internalName)}
                    />
                  )}
                  {fc.internalName === FORM_ATTACHMENTS_FIELD_INTERNAL ? (
                    <Text variant="small">(anexos)</Text>
                  ) : (
                    m && <Text variant="small">({m.MappedType})</Text>
                  )}
                </Stack>
              );
            })}
            {!fields.length && <Text>Nenhum campo no formulário. Use a aba Estrutura.</Text>}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Regras condicionais">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Nova regra" onClick={addConditionalCard} />
              <DefaultButton
                text="Modelo: mostrar B quando A = valor"
                onClick={() => applyPresetConditional('showWhenEq')}
              />
              <DefaultButton
                text="Modelo: obrigar B quando A = valor"
                onClick={() => applyPresetConditional('choiceRequire')}
              />
            </Stack>
            {conditionalCards.map((card, ci) => (
              <Stack
                key={card.id}
                styles={{ root: { border: '1px solid #edebe9', padding: 12, borderRadius: 4 } }}
                tokens={{ childrenGap: 8 }}
              >
                <Stack horizontal horizontalAlign="space-between">
                  <Text styles={{ root: { fontWeight: 600 } }}>{describeConditionalCardPT(card)}</Text>
                  <Stack horizontal tokens={{ childrenGap: 4 }}>
                    <DefaultButton text="Duplicar" onClick={() => duplicateCard(ci)} />
                    <DefaultButton text="Excluir" onClick={() => removeCard(ci)} />
                  </Stack>
                </Stack>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>Quando</Text>
                <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                  <Dropdown
                    label="Campo"
                    options={fieldOptions}
                    selectedKey={card.when.field}
                    onChange={(_, o) => o && patchWhen(ci, { field: String(o.key) })}
                  />
                  <Dropdown
                    label="Operador"
                    options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                    selectedKey={card.when.op}
                    onChange={(_, o) => o && patchWhen(ci, { op: o.key as TFormConditionOp })}
                  />
                  <Dropdown
                    label="Comparar com"
                    options={[
                      { key: 'literal', text: 'Texto fixo' },
                      { key: 'field', text: 'Outro campo' },
                      { key: 'token', text: 'Token' },
                    ]}
                    selectedKey={card.when.compareKind}
                    onChange={(_, o) => o && patchWhen(ci, { compareKind: o.key as IWhenUi['compareKind'] })}
                  />
                  <TextField
                    label="Valor"
                    value={card.when.compareValue}
                    onChange={(_, v) => patchWhen(ci, { compareValue: v ?? '' })}
                    disabled={
                      card.when.op === 'isEmpty' ||
                      card.when.op === 'isFilled' ||
                      card.when.op === 'isTrue' ||
                      card.when.op === 'isFalse'
                    }
                  />
                </Stack>
                <TextField
                  label="Grupos do SharePoint (títulos, vírgula)"
                  description="Vazio = qualquer utilizador. As regras geradas só aplicam se o utilizador pertencer a um destes grupos."
                  value={fieldNamesToCsv(card.groupTitles ?? [])}
                  onChange={(_, v) => {
                    const parsed = parseCsvFieldNames(v ?? '');
                    patchCard(ci, { groupTitles: parsed.length ? parsed : undefined });
                  }}
                />
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>Então</Text>
                {card.effects.map((eff, ei) => (
                  <Stack key={ei} horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                    <Dropdown
                      label="Efeito"
                      options={CONDITIONAL_EFFECT_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                      selectedKey={eff.kind}
                      onChange={(_, o) =>
                        o && patchEffect(ci, ei, { kind: o.key as TConditionalEffectKind })
                      }
                    />
                    {eff.kind !== 'message' && (
                      <Dropdown
                        label="Campo alvo"
                        options={[{ key: '', text: '—' }, ...fieldOptions]}
                        selectedKey={eff.targetField ?? ''}
                        onChange={(_, o) =>
                          patchEffect(ci, ei, { targetField: o ? String(o.key) : undefined })
                        }
                      />
                    )}
                    {eff.kind === 'message' && (
                      <>
                        <Dropdown
                          label="Tipo"
                          options={[
                            { key: 'info', text: 'Info' },
                            { key: 'warning', text: 'Aviso' },
                            { key: 'error', text: 'Erro' },
                          ]}
                          selectedKey={eff.messageVariant ?? 'info'}
                          onChange={(_, o) =>
                            o &&
                            patchEffect(ci, ei, { messageVariant: o.key as 'info' | 'warning' | 'error' })
                          }
                        />
                        <TextField
                          label="Texto"
                          value={eff.messageText ?? ''}
                          onChange={(_, v) => patchEffect(ci, ei, { messageText: v ?? '' })}
                        />
                      </>
                    )}
                    <DefaultButton text="Remover efeito" onClick={() => removeEffect(ci, ei)} />
                  </Stack>
                ))}
                <DefaultButton text="Adicionar efeito" onClick={() => addEffect(ci)} />
                <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                  Prévia: {compileConditionalCard(card).length} regra(s) gerada(s)
                </Text>
              </Stack>
            ))}
            {!conditionalCards.length && <Text>Nenhuma regra condicional. Use &quot;Nova regra&quot;.</Text>}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Ajuda dinâmica">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Etapas e campos: aba Estrutura. Layout das etapas e botões de navegação: aba Componentes. Aqui: ajuda condicional (JSON avançado).
            </Text>
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Ajuda dinâmica (JSON)</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Array de {'{'} field, when, helpText {'}'} — formato avançado.
            </Text>
            <TextField multiline rows={10} value={helpJson} onChange={(_, v) => setHelpJson(v ?? '')} />
          </Stack>
        </PivotItem>
        <PivotItem headerText="Gestor">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Colunas da grade gestor. Ordem abaixo.
            </Text>
            {meta.map((m) => (
              <Checkbox
                key={m.InternalName}
                label={`${m.Title} (${m.InternalName})`}
                checked={managerColumnFields.indexOf(m.InternalName) !== -1}
                onChange={(_, c) => toggleManagerCol(m.InternalName, !!c)}
              />
            ))}
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Ordem das colunas selecionadas (arraste pela alça)</Text>
            {managerColumnFields.map((name, mi) => (
              <Stack
                key={name}
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 8 }}
                onDragOver={(e) => {
                  e.preventDefault();
                  e.dataTransfer.dropEffect = 'move';
                }}
                onDrop={(e) => {
                  e.preventDefault();
                  const from = parseDragIndex(e.dataTransfer.getData('text/plain'), DND_MCOL);
                  if (from === undefined || from === mi) return;
                  reorderManagerCol(from, mi);
                }}
              >
                <span
                  draggable
                  title="Arrastar para reordenar"
                  onDragStart={(e) => {
                    e.dataTransfer.setData('text/plain', dragPayload(DND_MCOL, mi));
                    e.dataTransfer.effectAllowed = 'move';
                  }}
                  style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                >
                  <Icon iconName="GripperBarVertical" />
                </span>
                <Text styles={{ root: { minWidth: 160 } }}>{name}</Text>
              </Stack>
            ))}
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Anexos (formulário)</Text>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
              <TextField label="Mín. arquivos" value={attachMin} onChange={(_, v) => setAttachMin(v ?? '')} />
              <TextField label="Máx. arquivos" value={attachMax} onChange={(_, v) => setAttachMax(v ?? '')} />
              <TextField
                label="Mensagem"
                value={attachMsg}
                onChange={(_, v) => setAttachMsg(v ?? '')}
                styles={{ root: { minWidth: 280 } }}
              />
            </Stack>
          </Stack>
        </PivotItem>
      </Pivot>
      {!!customs.length && (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
          <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
            Regras só no motor (não editadas por esta UI)
          </Text>
          {customs.map((r) => (
            <Text key={r.id} variant="small" styles={{ root: { color: '#605e5c' } }}>
              {r.id}: {describeRule(r)}
            </Text>
          ))}
        </Stack>
      )}
      <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 24 } }}>
        <PrimaryButton text="Salvar" onClick={handleSave} disabled={loading} />
        <DefaultButton
          text="Restaurar padrão (estrutura)"
          onClick={() => {
            const d = getDefaultFormManagerConfig();
            const st = d.steps && d.steps.length ? d.steps : [{ id: 'main', title: 'Geral', fieldNames: [] }];
            setSteps(st.map((x) => ({ ...x, fieldNames: x.fieldNames.slice() })));
            setFields(d.fields.slice());
          }}
        />
        <DefaultButton text="Cancelar" onClick={onDismiss} />
      </Stack>
      <Panel
        isOpen={jsonOpen}
        type={PanelType.medium}
        headerText="JSON gerado (somente leitura)"
        onDismiss={() => setJsonOpen(false)}
      >
        <TextField multiline readOnly rows={22} value={previewConfigJson} />
        <DefaultButton styles={{ root: { marginTop: 12 } }} text="Fechar" onClick={() => setJsonOpen(false)} />
      </Panel>
      {fieldPanelName && fieldPanelConfig && (
        <FormFieldRulesPanel
          isOpen={true}
          internalName={fieldPanelName}
          fieldConfig={fieldPanelConfig}
          meta={fieldPanelMeta}
          rules={rules}
          fieldOptions={fieldOptions}
          onDismiss={() => setFieldPanelName(null)}
          onApply={(nextFc, editor) => {
            setFields((prev) => prev.map((f) => (f.internalName === fieldPanelName ? { ...f, ...nextFc } : f)));
            setRules((r) => mergeFieldRules(r, fieldPanelName, buildFieldUiRules(fieldPanelName, editor)));
          }}
        />
      )}
    </Panel>
  );
};
