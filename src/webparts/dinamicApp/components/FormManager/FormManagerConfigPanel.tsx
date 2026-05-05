import * as React from 'react';
import { useState, useEffect, useLayoutEffect, useMemo, useCallback, useRef } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  Modal,
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
  keyframes,
  mergeStyles,
} from '@fluentui/react';
import { FieldsService, GroupsService, filterSiteGroupsByNameQuery, mergeSystemMetadataFields } from '../../../../services';
import type { FieldMappedType, IFieldMetadata, IGroupDetails } from '../../../../services';
import type {
  IFormManagerConfig,
  IFormManagerActionLogConfig,
  IFormLinkedChildFormConfig,
  IFormStepNavigationConfig,
  IFormFieldConfig,
  IFormSectionConfig,
  IFormStepConfig,
  IFormCustomButtonConfig,
  TFormCustomButtonConfirmKind,
  TFormButtonAction,
  TFormConditionNode,
  TFormConditionOp,
  TFormCustomButtonBehavior,
  TFormCustomButtonOperation,
  TFormCustomButtonPaletteSlot,
  TFormManagerFormMode,
  TFormRule,
  TFormStepLayoutKind,
  TFormStepNavButtonsKind,
  TFormDataLoadingUiKind,
  TFormSubmitLoadingUiKind,
  TFormAttachmentUploadLayoutKind,
  TFormAttachmentFilePreviewKind,
  TFormHistoryPresentationKind,
  TFormHistoryLayoutKind,
  TFormHistoryButtonKind,
  TFormRootWidthMode,
  TFormRootHorizontalAlign,
  TFormAttachmentStorageKind,
  TFormCustomButtonsBarVertical,
  TFormCustomButtonsBarHorizontal,
  IAttachmentLibraryFolderTreeNode,
  IFormManagerAttachmentLibraryConfig,
  IFormManagerPermissionBreakConfig,
  TFormBannerPlacement,
  TFormAlertVariant,
} from '../../core/config/types/formManager';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_BANNER_INTERNAL_PREFIX,
  FORM_OCULTOS_STEP_ID,
  FORM_FIXOS_STEP_ID,
  FORM_BUILTIN_HISTORY_BUTTON_ID,
  FORM_SYSTEM_LIST_METADATA_INTERNAL_NAMES,
  isFormAlertFieldConfig,
  isFormBannerFieldConfig,
  resolveAlertPlacement,
  resolveAlertVariant,
  resolveBannerPlacement,
  resolveBannerWidthPercent,
  resolveFixedPlacement,
  resolveChromePositionMode,
  resolveFieldColumnSpan,
  type TFixedChromePlacement,
  type TChromePositionMode,
  type TFormFieldColumnSpan,
} from '../../core/config/types/formManager';
import { getDefaultFormManagerConfig } from '../../core/config/utils';
import { resolveFormCustomButtonPaletteSlot } from '../../core/formManager/formCustomButtonTheme';
import { mergeFormFieldConfigFromRulesPanel } from '../../core/formManager/mergeFormFieldConfigFromRulesPanel';
import { sanitizeFormManagerConfig } from '../../core/formManager/sanitizeFormManagerConfig';
import {
  attachmentFolderNodePathLabel,
  flattenFolderTreeNodes,
  loadFolderTreeFromAttachmentLibrary,
} from '../../core/formManager/attachmentFolderTree';
import { ALL_FORM_MANAGER_MODES, toggleStepShowInFormMode } from '../../core/formManager/stepFormMode';
import {
  buildFieldUiRules,
  customRulesOnly,
  describeRule,
  fieldRuleStateFromRules,
  isSetComputedAllowedForMappedType,
  mergeAttachmentUiRule,
  mergeFieldRules,
  parseAttachmentUiRule,
  CONDITION_OP_OPTIONS,
  type IWhenUi,
  whenUiToNode,
  whenNodeToUi,
  summarizeConditionTreePt,
} from '../../core/formManager/formManagerVisualModel';
import { FormFieldRulesPanel, FORM_FIELD_RULES_MENTION_PORTAL_ATTR } from './FormFieldRulesPanel';
import { FormManagerComponentsTabContent, FormManagerCollapseSection } from './FormManagerComponentsTab';
import { ThemePaletteSlotDropdown } from './ThemePaletteSlotDropdown';
import { FormManagerAttachmentsTabContent } from './FormManagerAttachmentsTab';
import type { IFolderVisibilityEditorProps } from './FormManagerFolderTreeEditor';
import {
  FORM_SUBMIT_LOADING_DROPDOWN_OPTIONS,
  FORM_SUBMIT_LOADING_INHERIT_KEY,
} from './FormLoadingUi';
import { FormManagerActionLogTabContent } from './FormManagerActionLogTab';
import { FormManagerLinkedChildFormsTabContent } from './FormManagerLinkedChildFormsTab';
import { FormManagerPermissionBreakTabContent } from './FormManagerPermissionBreakTab';
import { FormManagerChainedActionsBlock } from './FormManagerChainedActionsBlock';
import { isFormAttachmentLibraryRuntime } from '../../core/formManager/formAttachmentLibrary';
import { isConfirmPromptEligibleField } from '../../core/formManager/confirmPromptFieldHelpers';

const FIELD_RULES_TAB_SORT_TYPE_ORDER: readonly FieldMappedType[] = [
  'text',
  'multiline',
  'choice',
  'multichoice',
  'number',
  'currency',
  'boolean',
  'datetime',
  'url',
  'lookup',
  'lookupmulti',
  'user',
  'usermulti',
  'calculated',
  'taxonomy',
  'taxonomymulti',
  'unknown',
];

function fieldRulesTabMappedTypeOrderIndex(t: FieldMappedType | undefined): number {
  if (t === undefined) return 1000;
  const i = FIELD_RULES_TAB_SORT_TYPE_ORDER.indexOf(t);
  return i === -1 ? 999 : i;
}

const poolBulkMovePulse = keyframes({
  '0%, 100%': { transform: 'scale(1)', opacity: 1 },
  '50%': { transform: 'scale(1.12)', opacity: 0.82 },
});

const poolBulkMoveIconClassName = mergeStyles({
  animation: `${poolBulkMovePulse} 1.15s ease-in-out infinite`,
  borderRadius: 6,
  background: '#edebe9',
});

function structureStepMenuLabel(step: IFormStepConfig, stepIdx: number): string {
  if (step.id === FORM_OCULTOS_STEP_ID) return 'Ocultos';
  if (step.id === FORM_FIXOS_STEP_ID) return 'Fixos';
  const t = step.title?.trim();
  return t || `Etapa ${stepIdx + 1}`;
}

function buildRulesCloneFieldPatch(src: IFormFieldConfig): Partial<IFormFieldConfig> {
  const patch: Partial<IFormFieldConfig> = {};
  const keys: (keyof IFormFieldConfig)[] = [
    'textConditionalVisibility',
    'textInputMaskKind',
    'textInputMaskCustomPattern',
    'textValueTransform',
    'visible',
  ];
  for (let i = 0; i < keys.length; i++) {
    const k = keys[i];
    const v = src[k];
    if (v !== undefined) {
      (patch as Record<string, unknown>)[k as string] =
        v !== null && typeof v === 'object' ? JSON.parse(JSON.stringify(v)) : v;
    }
  }
  if ('readOnly' in src) {
    patch.readOnly = src.readOnly;
  }
  return patch;
}

function attachmentLibraryFromPanelState(
  libraryTitle: string,
  sourceListLookupFieldInternalName: string,
  folderTree: IAttachmentLibraryFolderTreeNode[]
): IFormManagerAttachmentLibraryConfig | undefined {
  const lt = libraryTitle.trim();
  const lk = sourceListLookupFieldInternalName.trim();
  if (!lt && !lk && !folderTree.length) return undefined;
  return {
    ...(lt ? { libraryTitle: lt } : {}),
    ...(lk ? { sourceListLookupFieldInternalName: lk } : {}),
    ...(folderTree.length ? { folderTree } : {}),
  };
}

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

function mergePermissionBreakFromCfg(raw?: IFormManagerPermissionBreakConfig): IFormManagerPermissionBreakConfig {
  return {
    enabled: raw?.enabled === true,
    copyInheritedAssignments: raw?.copyInheritedAssignments === true,
    retainAuthor: raw?.retainAuthor !== false,
    authorRoleDefinitionName: (raw?.authorRoleDefinitionName ?? 'Contribuir').trim().slice(0, 120),
    targets: {
      mainListItem: raw?.targets?.mainListItem !== false,
      linkedChildFormIds: raw?.targets?.linkedChildFormIds?.slice(),
      mainAttachmentLibraryFiles: raw?.targets?.mainAttachmentLibraryFiles === true,
      linkedAttachmentLibraryFilesByFormId: raw?.targets?.linkedAttachmentLibraryFilesByFormId?.slice(),
    },
    assignments: (raw?.assignments ?? []).map((a) => ({ ...a })),
  };
}

function cloneLinkedChildFormConfig(c: IFormLinkedChildFormConfig): IFormLinkedChildFormConfig {
  const lib = c.childAttachmentLibrary;
  return {
    ...c,
    sections: c.sections.map((s) => ({ ...s })),
    fields: c.fields.map((f) => ({ ...f })),
    rules: (c.rules ?? []).map((r) => JSON.parse(JSON.stringify(r)) as TFormRule),
    steps: (c.steps ?? []).map((s) => ({ ...s, fieldNames: [...s.fieldNames] })),
    ...(lib
      ? {
          childAttachmentLibrary: {
            ...lib,
            folderTree: loadFolderTreeFromAttachmentLibrary(lib),
          },
        }
      : {}),
  };
}

function patchLinkedChildFormById(
  arr: IFormLinkedChildFormConfig[],
  id: string,
  patch: Partial<IFormLinkedChildFormConfig>
): IFormLinkedChildFormConfig[] {
  return arr.map((c) => (c.id === id ? { ...c, ...patch } : c));
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

const DND_STEP = 'fm/step:';
const DND_POOL = 'fm/pool:';
const DND_FS = 'fm/fs:';
const DND_BTN = 'fm/btn:';

const BANNER_PLACEMENT_DROPDOWN_OPTIONS: IDropdownOption[] = [
  { key: 'inStep', text: 'Na etapa (ordem com os campos)' },
  { key: 'topFixed', text: 'Fixo no topo (sticky)' },
  { key: 'bottomFixed', text: 'Fixo em baixo (sticky)' },
];

const FIXED_CHROME_PLACEMENT_OPTIONS: IDropdownOption[] = [
  { key: 'top', text: 'Fixo no topo' },
  { key: 'bottom', text: 'Fixo em baixo' },
];

const CHROME_POSITION_MODE_OPTIONS: IDropdownOption[] = [
  { key: 'sticky', text: 'Fixo (acompanha ao scroll)' },
  { key: 'absolute', text: 'Absoluto (sobre o conteúdo)' },
  { key: 'flow', text: 'No espaço (fluxo normal)' },
];

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

function normalizeFixosFieldConfigs(flds: IFormFieldConfig[]): IFormFieldConfig[] {
  return flds.map((fc) => {
    if (fc.sectionId === FORM_FIXOS_STEP_ID) {
      if (!fc.fixedPlacement) return { ...fc, fixedPlacement: 'top' };
      return fc;
    }
    if (fc.fixedPlacement) {
      const { fixedPlacement: _fp, ...rest } = fc;
      return rest;
    }
    return fc;
  });
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
  return normalizeFixosFieldConfigs(out);
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

function alertWhenUiFromNode(node: TFormConditionNode | undefined, meta: IFieldMetadata[]): IWhenUi {
  return whenNodeToUi(node) ?? defaultWhenUi(meta);
}

function parseButtonWhenToRows(
  w: TFormConditionNode | undefined,
  meta: IFieldMetadata[]
): { combiner: 'all' | 'any'; rows: IWhenUi[] } {
  if (!w) return { combiner: 'all', rows: [defaultWhenUi(meta)] };
  if (w.kind === 'leaf') {
    const u = whenNodeToUi(w);
    return { combiner: 'all', rows: u ? [u] : [defaultWhenUi(meta)] };
  }
  if (w.kind === 'all' || w.kind === 'any') {
    const rows: IWhenUi[] = [];
    for (let i = 0; i < w.children.length; i++) {
      const u = whenNodeToUi(w.children[i]);
      if (u) rows.push(u);
    }
    return { combiner: w.kind, rows: rows.length ? rows : [defaultWhenUi(meta)] };
  }
  return { combiner: 'all', rows: [defaultWhenUi(meta)] };
}

function whenUiRowCompletesCondition(r: IWhenUi): boolean {
  if (r.compareKind === 'spGroupMember' || r.compareKind === 'spGroupNotMember') {
    return r.compareValue.trim().length > 0;
  }
  return r.field.trim().length > 0;
}

function buildButtonWhenFromRows(combiner: 'all' | 'any', rows: IWhenUi[]): TFormConditionNode | undefined {
  const valid = rows.filter(whenUiRowCompletesCondition);
  if (valid.length === 0) return undefined;
  const nodes = valid.map(whenUiToNode);
  if (nodes.length === 1) return nodes[0];
  return { kind: combiner, children: nodes };
}

function formatStepModesHint(step: IFormStepConfig): string {
  const sel = step.showInFormModes;
  if (!sel?.length) return 'Modos: todos';
  const labels: Record<TFormManagerFormMode, string> = { create: 'Criar', edit: 'Editar', view: 'Ver' };
  return `Modos: ${sel.map((m) => labels[m]).join(', ')}`;
}

function normSpGroupTitle(s: string): string {
  return s.trim().toLowerCase();
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

const BUTTON_OPERATION_OPTIONS_BASE: IDropdownOption[] = [
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

const BUTTON_FINISH_AFTER_OPTIONS: IDropdownOption[] = [
  { key: 'none', text: 'Nada' },
  { key: 'redirect', text: 'Redirecionar' },
  { key: 'clearForm', text: 'Limpar o formulário' },
];

const ESTRUTURA_COLLAPSE_IDS = {
  formLayout: 'estruturaFormLayout',
  stepNav: 'estruturaStepNav',
} as const;

const FORM_ROOT_WIDTH_OPTIONS: IDropdownOption[] = [
  { key: 'percent', text: 'Percentagem da área disponível' },
  { key: 'full', text: 'Largura total (100%)' },
];

const FORM_ROOT_ALIGN_OPTIONS: IDropdownOption[] = [
  { key: 'start', text: 'Início (esquerda)' },
  { key: 'center', text: 'Centro' },
  { key: 'end', text: 'Fim (direita)' },
];

const FIELD_COLUMN_SPAN_OPTIONS: IDropdownOption[] = [
  { key: '12', text: '12 — linha inteira' },
  { key: '8', text: '8 — ex.: 8+4' },
  { key: '6', text: '6 — ex.: 6+6' },
  { key: '4', text: '4 — ex.: 4+4+4' },
  { key: '3', text: '3 — ex.: 3+3+3+3' },
];

function formatFieldColumnSpanConfigSummary(fc: IFormFieldConfig | undefined, fname: string): string {
  const base: Pick<IFormFieldConfig, 'internalName' | 'columnSpan' | 'width' | 'columnSpanByMode'> = fc ?? {
    internalName: fname,
  };
  const n = resolveFieldColumnSpan(base, 'create');
  const v = resolveFieldColumnSpan(base, 'view');
  const e = resolveFieldColumnSpan(base, 'edit');
  if (n === v && v === e) return String(n);
  return `N${n} · V${v} · E${e}`;
}

const COLUMN_SPAN_BY_MODE_TABS: { mode: TFormManagerFormMode; headerText: string }[] = [
  { mode: 'create', headerText: 'Novo' },
  { mode: 'view', headerText: 'Ver' },
  { mode: 'edit', headerText: 'Editar' },
];

function clampFormRootPercentInput(s: string): number {
  const n = Number(String(s).replace(',', '.').trim());
  if (!isFinite(n)) return 100;
  return Math.min(100, Math.max(1, Math.round(n)));
}

function clampFormRootPaddingInput(s: string): number {
  const t = String(s).replace(',', '.').trim();
  if (t === '') return 0;
  const n = Number(t);
  if (!isFinite(n) || n < 1) return 0;
  return Math.min(160, Math.max(1, Math.round(n)));
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

function pinOcultosStepFirst(st: IFormStepConfig[]): IFormStepConfig[] {
  const oi = st.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
  if (oi <= 0) return st.map((s) => ({ ...s, fieldNames: s.fieldNames.slice() }));
  const out = st.map((s) => ({ ...s, fieldNames: s.fieldNames.slice() }));
  const [oc] = out.splice(oi, 1);
  out.unshift(oc);
  return out;
}

function pinFixosAfterOcultos(st: IFormStepConfig[]): IFormStepConfig[] {
  const out = st.map((s) => ({ ...s, fieldNames: s.fieldNames.slice() }));
  const fi = out.findIndex((s) => s.id === FORM_FIXOS_STEP_ID);
  if (fi < 0) return out;
  const oi = out.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
  const wantIdx = oi >= 0 ? oi + 1 : 0;
  if (fi === wantIdx) return out;
  const [fx] = out.splice(fi, 1);
  const insertAt = fi < wantIdx ? wantIdx - 1 : wantIdx;
  out.splice(insertAt, 0, fx);
  return out;
}

function pinCoreStepsOrder(st: IFormStepConfig[]): IFormStepConfig[] {
  return pinFixosAfterOcultos(pinOcultosStepFirst(st));
}

function ensureCoreSteps(st: IFormStepConfig[]): IFormStepConfig[] {
  if (st.length === 0) {
    return [
      { id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] },
      { id: FORM_FIXOS_STEP_ID, title: 'Fixos', fieldNames: [] },
      { id: 'main', title: 'Geral', fieldNames: [] },
    ];
  }
  let out = st.map((s) => ({ ...s, fieldNames: s.fieldNames.slice() }));
  if (!out.some((s) => s.id === 'main')) {
    out.push({ id: 'main', title: 'Geral', fieldNames: [] });
  }
  if (!out.some((s) => s.id === FORM_OCULTOS_STEP_ID)) {
    out.unshift({ id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] });
  } else {
    out = pinOcultosStepFirst(out);
  }
  if (!out.some((s) => s.id === FORM_FIXOS_STEP_ID)) {
    const oi = out.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
    out.splice(oi >= 0 ? oi + 1 : 0, 0, { id: FORM_FIXOS_STEP_ID, title: 'Fixos', fieldNames: [] });
  } else {
    out = pinFixosAfterOcultos(out);
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

/**
 * Persistência (modo formulário): cada aba do Pivot alimenta `handleSave` → `IFormManagerConfig` →
 * `sanitizeFormManagerConfig` → `configJson` da webpart. Ao recarregar, `parseConfig` volta a sanitizar o JSON.
 *
 * | Aba | Chaves principais em `IFormManagerConfig` |
 * | --- | --- |
 * | Estrutura | `steps`, `sections`, `fields`, `rules` (merge anexos), `stepNavigation` |
 * | Componentes | `stepLayout`, `stepAccentPaletteSlot`, `stepNavButtons`, `formDataLoadingKind`, `defaultSubmitLoadingKind`, `formRootWidthMode`, `formRootWidthPercent`, `formRootHorizontalAlign`, `formRootPaddingPx`, `managerColumnFields`, `dynamicHelp`, `attachmentUploadLayout`, `attachmentFilePreview`, `historyEnabled`, `historyPresentationKind`, `historyLayoutKind`, `historyButtonKind`, `historyButtonLabel`, `historyButtonIcon`, `historyPanelSubtitle`, `historyGroupTitles` |
 * | Anexos | `attachmentStorageKind` (`itemAttachments` \| `documentLibrary`), `attachmentLibrary` |
 * | Botões | `customButtons`, `customButtonsBarVertical`, `customButtonsBarHorizontal` |
 * | Lista de logs | `actionLog` (lista, captação, textos por botão) |
 * | Listas vinculadas | `linkedChildForms` |
 * | Quebra de permissões | `permissionBreak` |
 * | Regras dos campos | regras por campo (painel) + resto de `rules` no motor |
 * | JSON | mesmo modelo (JSON inválido ou regras com `action` desconhecida são descartadas no sanitize) |
 */
export interface IFormManagerConfigPanelProps {
  isOpen: boolean;
  listTitle: string;
  listWebServerRelativeUrl?: string;
  value: IFormManagerConfig;
  onSave: (next: IFormManagerConfig) => void;
  onDismiss: () => void;
}

export const FormManagerConfigPanel: React.FC<IFormManagerConfigPanelProps> = ({
  isOpen,
  listTitle,
  listWebServerRelativeUrl,
  value,
  onSave,
  onDismiss,
}) => {
  const lw = listWebServerRelativeUrl?.trim() || undefined;
  const [fields, setFields] = useState<IFormFieldConfig[]>(() => buildInitialFieldsAndSteps(value).fields);
  const [rules, setRules] = useState<TFormRule[]>(() => value.rules ?? []);
  const [steps, setSteps] = useState<IFormStepConfig[]>(() => buildInitialFieldsAndSteps(value).steps);
  const [helpJson, setHelpJson] = useState(() => JSON.stringify(value.dynamicHelp ?? [], null, 2));
  const [managerColumnFields, setManagerColumnFields] = useState<string[]>(() => value.managerColumnFields ?? []);
  const [customButtons, setCustomButtons] = useState<IFormCustomButtonConfig[]>(() =>
    (value.customButtons ?? []).map((b) => ({
      ...b,
      actions: (b.actions ?? []).map((a) => ({ ...a })),
    }))
  );
  const [stepLayout, setStepLayout] = useState<TFormStepLayoutKind>(() => value.stepLayout ?? 'segmented');
  const [stepNavButtons, setStepNavButtons] = useState<TFormStepNavButtonsKind>(
    () => value.stepNavButtons ?? 'fluent'
  );
  const [stepAccentPaletteSlot, setStepAccentPaletteSlot] = useState<
    TFormCustomButtonPaletteSlot | undefined
  >(() => value.stepAccentPaletteSlot);
  const [formDataLoadingKind, setFormDataLoadingKind] = useState<TFormDataLoadingUiKind>(
    () => value.formDataLoadingKind ?? 'spinner'
  );
  const [defaultSubmitLoadingKind, setDefaultSubmitLoadingKind] = useState<TFormSubmitLoadingUiKind>(
    () => value.defaultSubmitLoadingKind ?? 'overlay'
  );
  const [attachmentUploadLayout, setAttachmentUploadLayout] = useState<TFormAttachmentUploadLayoutKind>(
    () => value.attachmentUploadLayout ?? 'default'
  );
  const [attachmentFilePreview, setAttachmentFilePreview] = useState<TFormAttachmentFilePreviewKind>(
    () => value.attachmentFilePreview ?? 'nameAndSize'
  );
  const [attachmentStorageKind, setAttachmentStorageKind] = useState<TFormAttachmentStorageKind>(
    () => value.attachmentStorageKind ?? 'itemAttachments'
  );
  const [attachmentLibLibraryTitle, setAttachmentLibLibraryTitle] = useState(
    () => value.attachmentLibrary?.libraryTitle ?? ''
  );
  const [attachmentLibLookupField, setAttachmentLibLookupField] = useState(
    () => value.attachmentLibrary?.sourceListLookupFieldInternalName ?? ''
  );
  const [attachmentLibFolderTree, setAttachmentLibFolderTree] = useState<IAttachmentLibraryFolderTreeNode[]>(() =>
    loadFolderTreeFromAttachmentLibrary(value.attachmentLibrary)
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
  const [formRootWidthMode, setFormRootWidthMode] = useState<TFormRootWidthMode>(
    () => value.formRootWidthMode ?? 'percent'
  );
  const [formRootWidthPercent, setFormRootWidthPercent] = useState(() =>
    String(value.formRootWidthPercent ?? 100)
  );
  const [formRootHorizontalAlign, setFormRootHorizontalAlign] = useState<TFormRootHorizontalAlign>(
    () => value.formRootHorizontalAlign ?? 'start'
  );
  const [formRootPaddingPx, setFormRootPaddingPx] = useState(() =>
    value.formRootPaddingPx != null && value.formRootPaddingPx >= 1
      ? String(value.formRootPaddingPx)
      : ''
  );
  const [estruturaOpen, setEstruturaOpen] = useState<Record<string, boolean>>({});
  const toggleEstruturaSection = (id: string): void => {
    setEstruturaOpen((prev) => ({ ...prev, [id]: !prev[id] }));
  };
  const isEstruturaOpen = (id: string): boolean => estruturaOpen[id] === true;
  const [stepSectionOpen, setStepSectionOpen] = useState<Record<string, boolean>>({});
  const [stepVisibilityPanelStepId, setStepVisibilityPanelStepId] = useState<string | null>(null);
  const [buttonSectionOpen, setButtonSectionOpen] = useState<Record<string, boolean>>({});
  const [customButtonsBarVertical, setCustomButtonsBarVertical] = useState<TFormCustomButtonsBarVertical>(
    () => value.customButtonsBarVertical ?? 'bottom'
  );
  const [customButtonsBarHorizontal, setCustomButtonsBarHorizontal] = useState<TFormCustomButtonsBarHorizontal>(
    () => value.customButtonsBarHorizontal ?? 'left'
  );
  const [attachMin, setAttachMin] = useState('');
  const [attachMax, setAttachMax] = useState('');
  const [attachMsg, setAttachMsg] = useState('');
  const [attachAllowedExt, setAttachAllowedExt] = useState<string[]>(() =>
    parseAttachmentUiRule(value.rules ?? []).allowedFileExtensions ?? []
  );
  const [meta, setMeta] = useState<IFieldMetadata[]>([]);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState<string | undefined>(undefined);
  const [jsonOpen, setJsonOpen] = useState(false);
  const [jsonPanelText, setJsonPanelText] = useState('');
  const [jsonPanelErr, setJsonPanelErr] = useState<string | undefined>(undefined);
  const [fieldPanelName, setFieldPanelName] = useState<string | null>(null);
  const [fieldRulesTabSort, setFieldRulesTabSort] = useState<'asc' | 'desc' | 'type'>('asc');
  const [cloneRulesModalTarget, setCloneRulesModalTarget] = useState<string | null>(null);
  const [cloneRulesSourceKey, setCloneRulesSourceKey] = useState<string | undefined>(undefined);
  const [columnSpanModalField, setColumnSpanModalField] = useState<string | null>(null);
  const [structurePoolSelected, setStructurePoolSelected] = useState<string[]>([]);
  const [structureFieldOpen, setStructureFieldOpen] = useState<Record<string, boolean>>({});
  const structurePoolSelectedRef = useRef<string[]>([]);
  structurePoolSelectedRef.current = structurePoolSelected;
  const [redirectReplaceBraceForBtnId, setRedirectReplaceBraceForBtnId] = useState<string | null>(null);
  const [redirectInsertNonceByBtn, setRedirectInsertNonceByBtn] = useState<Record<string, number>>({});
  const [redirectReplaceNonceByBtn, setRedirectReplaceNonceByBtn] = useState<Record<string, number>>({});
  const [siteGroups, setSiteGroups] = useState<IGroupDetails[]>([]);
  const [siteGroupsLoading, setSiteGroupsLoading] = useState(false);
  const [siteGroupsErr, setSiteGroupsErr] = useState<string | undefined>(undefined);
  const [customButtonGroupNameFilter, setCustomButtonGroupNameFilter] = useState('');
  const [actionLogCaptureEnabled, setActionLogCaptureEnabled] = useState(false);
  const [actionLogListTitle, setActionLogListTitle] = useState('');
  const [actionLogFieldInternalName, setActionLogFieldInternalName] = useState('');
  const [actionLogSourceListLookupFieldInternalName, setActionLogSourceListLookupFieldInternalName] =
    useState('');
  const [actionLogDescById, setActionLogDescById] = useState<Record<string, string>>({});
  const [actionLogPaletteSlotById, setActionLogPaletteSlotById] = useState<
    Record<string, TFormCustomButtonPaletteSlot>
  >({});
  const [actionLogAutomaticChangesOnUpdate, setActionLogAutomaticChangesOnUpdate] = useState(false);
  const [historyEnabled, setHistoryEnabled] = useState(() => value.historyEnabled === true);
  const [historyPresentationKind, setHistoryPresentationKind] = useState<TFormHistoryPresentationKind>(
    () => value.historyPresentationKind ?? 'panel'
  );
  const [historyLayoutKind, setHistoryLayoutKind] = useState<TFormHistoryLayoutKind>(
    () => value.historyLayoutKind ?? 'list'
  );
  const [historyButtonKind, setHistoryButtonKind] = useState<TFormHistoryButtonKind>(
    () => value.historyButtonKind ?? 'text'
  );
  const [historyButtonLabel, setHistoryButtonLabel] = useState(() => value.historyButtonLabel ?? 'Histórico');
  const [historyButtonIcon, setHistoryButtonIcon] = useState(() => value.historyButtonIcon ?? 'History');
  const [historyPanelSubtitle, setHistoryPanelSubtitle] = useState(() => value.historyPanelSubtitle ?? '');
  const [historyGroupTitles, setHistoryGroupTitles] = useState<string[]>(() => value.historyGroupTitles ?? []);
  const [linkedChildForms, setLinkedChildForms] = useState<IFormLinkedChildFormConfig[]>(() =>
    (value.linkedChildForms ?? []).map(cloneLinkedChildFormConfig)
  );
  const [permissionBreak, setPermissionBreak] = useState<IFormManagerPermissionBreakConfig>(() =>
    mergePermissionBreakFromCfg(value.permissionBreak)
  );

  const fieldsService = useMemo(() => new FieldsService(), []);
  const groupsService = useMemo(() => new GroupsService(), []);
  const attachmentFolderStepOptions = useMemo(
    () =>
      steps
        .filter((s) => s.id !== FORM_OCULTOS_STEP_ID && s.id !== FORM_FIXOS_STEP_ID)
        .map((s) => ({ id: s.id, title: s.title })),
    [steps]
  );

  const linkedMainStepPlacementOptions = useMemo((): IDropdownOption[] => {
    return steps
      .filter((s) => s.id !== FORM_OCULTOS_STEP_ID && s.id !== FORM_FIXOS_STEP_ID)
      .map((s) => ({ key: s.id, text: `${s.title} (${s.id})` }));
  }, [steps]);

  const linkedChildFormsSortedForStructure = useMemo(
    () => linkedChildForms.slice().sort((a, b) => (a.order ?? 0) - (b.order ?? 0)),
    [linkedChildForms]
  );

  const formManagerForPermissionBreakResolve = useMemo(
    (): IFormManagerConfig => ({
      ...getDefaultFormManagerConfig(),
      attachmentStorageKind,
      attachmentLibrary: attachmentLibraryFromPanelState(
        attachmentLibLibraryTitle,
        attachmentLibLookupField,
        attachmentLibFolderTree
      ),
      linkedChildForms: linkedChildForms.map(cloneLinkedChildFormConfig),
    }),
    [
      attachmentStorageKind,
      attachmentLibLibraryTitle,
      attachmentLibLookupField,
      attachmentLibFolderTree,
      linkedChildForms,
    ]
  );

  const mainAttachmentLibraryEnabledTab = useMemo(
    () => isFormAttachmentLibraryRuntime(formManagerForPermissionBreakResolve),
    [formManagerForPermissionBreakResolve]
  );

  const linkedMainStepDefaultKey = useMemo(() => {
    const k = linkedMainStepPlacementOptions[0]?.key;
    return typeof k === 'string' ? k : 'main';
  }, [linkedMainStepPlacementOptions]);

  const attachmentFolderOptionsForFieldRules = useMemo(() => {
    if (attachmentStorageKind !== 'documentLibrary' || !attachmentLibFolderTree.length) return [];
    return flattenFolderTreeNodes(attachmentLibFolderTree).map((n) => ({
      key: n.id,
      text: attachmentFolderNodePathLabel(attachmentLibFolderTree, n.id),
    }));
  }, [attachmentStorageKind, attachmentLibFolderTree]);

  const hydrateFromFormManagerConfig = useCallback((cfg: IFormManagerConfig) => {
    const norm = buildInitialFieldsAndSteps(cfg);
    setFields(norm.fields);
    setSteps(norm.steps);
    setRules(cfg.rules ?? []);
    setHelpJson(JSON.stringify(cfg.dynamicHelp ?? [], null, 2));
    setManagerColumnFields(cfg.managerColumnFields ?? []);
    setCustomButtons(
      (cfg.customButtons ?? []).map((b) => ({
        ...b,
        actions: (b.actions ?? []).map((a) => ({ ...a })),
      }))
    );
    setStepLayout(cfg.stepLayout ?? 'segmented');
    setStepAccentPaletteSlot(cfg.stepAccentPaletteSlot);
    setStepNavButtons(cfg.stepNavButtons ?? 'fluent');
    setFormDataLoadingKind(cfg.formDataLoadingKind ?? 'spinner');
    setDefaultSubmitLoadingKind(cfg.defaultSubmitLoadingKind ?? 'overlay');
    setAttachmentUploadLayout(cfg.attachmentUploadLayout ?? 'default');
    setAttachmentFilePreview(cfg.attachmentFilePreview ?? 'nameAndSize');
    setAttachmentStorageKind(cfg.attachmentStorageKind ?? 'itemAttachments');
    setAttachmentLibLibraryTitle(cfg.attachmentLibrary?.libraryTitle ?? '');
    setAttachmentLibLookupField(cfg.attachmentLibrary?.sourceListLookupFieldInternalName ?? '');
    setAttachmentLibFolderTree(loadFolderTreeFromAttachmentLibrary(cfg.attachmentLibrary));
    setStepRequireFilledToAdvance(cfg.stepNavigation?.requireFilledRequiredToAdvance === true);
    setStepFullValOnAdvance(cfg.stepNavigation?.fullValidationOnAdvance === true);
    setStepAllowBackWithoutVal(cfg.stepNavigation?.allowBackWithoutValidation !== false);
    setFormRootWidthMode(cfg.formRootWidthMode ?? 'percent');
    setFormRootWidthPercent(String(cfg.formRootWidthPercent ?? 100));
    setFormRootHorizontalAlign(cfg.formRootHorizontalAlign ?? 'start');
    setFormRootPaddingPx(
      cfg.formRootPaddingPx != null && cfg.formRootPaddingPx >= 1 ? String(cfg.formRootPaddingPx) : ''
    );
    const att = parseAttachmentUiRule(cfg.rules ?? []);
    setAttachMin(att.minCount);
    setAttachMax(att.maxCount);
    setAttachMsg(att.message);
    setAttachAllowedExt(att.allowedFileExtensions ?? []);
    setErr(undefined);
    setFieldPanelName(null);
    setStepSectionOpen({});
    setButtonSectionOpen({});
    setCustomButtonsBarVertical(cfg.customButtonsBarVertical ?? 'bottom');
    setCustomButtonsBarHorizontal(cfg.customButtonsBarHorizontal ?? 'left');
    setActionLogCaptureEnabled(cfg.actionLog?.captureEnabled === true);
    setActionLogListTitle(cfg.actionLog?.listTitle ?? '');
    setActionLogFieldInternalName(cfg.actionLog?.actionFieldInternalName ?? '');
    setActionLogSourceListLookupFieldInternalName(cfg.actionLog?.sourceListLookupFieldInternalName ?? '');
    setActionLogDescById(
      cfg.actionLog?.descriptionsHtmlByButtonId
        ? { ...cfg.actionLog.descriptionsHtmlByButtonId }
        : {}
    );
    setActionLogPaletteSlotById(
      cfg.actionLog?.descriptionPaletteSlotByButtonId
        ? { ...cfg.actionLog.descriptionPaletteSlotByButtonId }
        : {}
    );
    setActionLogAutomaticChangesOnUpdate(cfg.actionLog?.automaticChangesOnUpdate === true);
    setHistoryEnabled(cfg.historyEnabled === true);
    setHistoryPresentationKind(cfg.historyPresentationKind ?? 'panel');
    setHistoryLayoutKind(cfg.historyLayoutKind ?? 'list');
    setHistoryButtonKind(cfg.historyButtonKind ?? 'text');
    setHistoryButtonLabel(cfg.historyButtonLabel ?? 'Histórico');
    setHistoryButtonIcon(cfg.historyButtonIcon ?? 'History');
    setHistoryPanelSubtitle(cfg.historyPanelSubtitle ?? '');
    setHistoryGroupTitles(cfg.historyGroupTitles ?? []);
    setLinkedChildForms((cfg.linkedChildForms ?? []).map(cloneLinkedChildFormConfig));
    setPermissionBreak(mergePermissionBreakFromCfg(cfg.permissionBreak));
  }, []);

  const formManagerValueRef = useRef(value);
  formManagerValueRef.current = value;
  useLayoutEffect(() => {
    if (!isOpen) return;
    hydrateFromFormManagerConfig(value);
  }, [isOpen, value, hydrateFromFormManagerConfig]);

  useEffect(() => {
    setActionLogDescById((prev) => {
      const next = { ...prev };
      if (historyEnabled) {
        if (!(FORM_BUILTIN_HISTORY_BUTTON_ID in next)) {
          next[FORM_BUILTIN_HISTORY_BUTTON_ID] = prev[FORM_BUILTIN_HISTORY_BUTTON_ID] ?? '';
        }
      } else {
        delete next[FORM_BUILTIN_HISTORY_BUTTON_ID];
      }
      for (let i = 0; i < customButtons.length; i++) {
        const id = customButtons[i].id;
        if (!(id in next)) next[id] = '';
      }
      const keys = Object.keys(next);
      for (let k = 0; k < keys.length; k++) {
        const key = keys[k];
        if (key === FORM_BUILTIN_HISTORY_BUTTON_ID) continue;
        let found = false;
        for (let j = 0; j < customButtons.length; j++) {
          if (customButtons[j].id === key) {
            found = true;
            break;
          }
        }
        if (!found) delete next[key];
      }
      return next;
    });
  }, [customButtons, historyEnabled]);

  useEffect(() => {
    setActionLogPaletteSlotById((prev) => {
      const next = { ...prev };
      if (historyEnabled) {
        if (!(FORM_BUILTIN_HISTORY_BUTTON_ID in next)) {
          next[FORM_BUILTIN_HISTORY_BUTTON_ID] = 'themePrimary';
        }
      } else {
        delete next[FORM_BUILTIN_HISTORY_BUTTON_ID];
      }
      for (let i = 0; i < customButtons.length; i++) {
        const id = customButtons[i].id;
        if (!(id in next)) next[id] = 'themePrimary';
      }
      const keys = Object.keys(next);
      for (let k = 0; k < keys.length; k++) {
        const key = keys[k];
        if (key === FORM_BUILTIN_HISTORY_BUTTON_ID) continue;
        let found = false;
        for (let j = 0; j < customButtons.length; j++) {
          if (customButtons[j].id === key) {
            found = true;
            break;
          }
        }
        if (!found) delete next[key];
      }
      return next;
    });
  }, [customButtons, historyEnabled]);

  useEffect(() => {
    if (!isOpen || !listTitle.trim()) return;
    setLoading(true);
    fieldsService
      .getVisibleFields(listTitle.trim(), lw)
      .then((f) => {
        setMeta(mergeSystemMetadataFields(f));
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, [isOpen, listTitle, lw]);

  const loadSiteGroups = useCallback((): void => {
    setSiteGroupsErr(undefined);
    setSiteGroupsLoading(true);
    groupsService
      .getSiteGroups()
      .then((g) => {
        setSiteGroups(g);
        setSiteGroupsLoading(false);
      })
      .catch((e) => {
        setSiteGroups([]);
        setSiteGroupsLoading(false);
        setSiteGroupsErr(e instanceof Error ? e.message : String(e));
      });
  }, [groupsService]);

  useEffect(() => {
    if (!isOpen) return;
    loadSiteGroups();
  }, [isOpen, loadSiteGroups]);

  const siteGroupsSorted = useMemo(() => {
    const g = siteGroups.slice();
    g.sort((a, b) => (a.Title < b.Title ? -1 : a.Title > b.Title ? 1 : 0));
    return g;
  }, [siteGroups]);

  const siteGroupsSortedForCustomButtons = useMemo(
    () => filterSiteGroupsByNameQuery(siteGroupsSorted, customButtonGroupNameFilter),
    [siteGroupsSorted, customButtonGroupNameFilter]
  );

  const buttonOperationDropdownOptions = useMemo((): IDropdownOption[] => {
    const opts = BUTTON_OPERATION_OPTIONS_BASE.slice();
    if (customButtons.some((b) => b.operation === 'history')) {
      opts.push({ key: 'history', text: 'Histórico (legado — use Componentes + Lista de logs)' });
    }
    return opts;
  }, [customButtons]);

  const fieldOptions: IDropdownOption[] = useMemo(
    () =>
      meta
        .filter(isFormConfigSelectableField)
        .map((f) => ({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` })),
    [meta]
  );

  const confirmModalPromptFieldOptions = useMemo((): IDropdownOption[] => {
    const rows = meta.filter(isConfirmPromptEligibleField).map((m) => ({
      key: m.InternalName,
      text: `${m.Title} (${m.InternalName})`,
    }));
    return [{ key: '', text: '— Sem campo extra no modal —' }, ...rows];
  }, [meta]);

  const attachmentFolderVisibilityEditor = useMemo((): IFolderVisibilityEditorProps => {
    return {
      fieldOptions,
      defaultConditionFieldName: meta[0]?.InternalName ?? 'Title',
      siteGroupsSorted,
      siteGroups,
      siteGroupsLoading,
      siteGroupsErr,
      onReloadSiteGroups: loadSiteGroups,
    };
  }, [fieldOptions, meta, siteGroupsSorted, siteGroups, siteGroupsLoading, siteGroupsErr, loadSiteGroups]);

  const customs = useMemo(() => customRulesOnly(rules), [rules]);

  const fieldsListedForRulesTab = useMemo(() => {
    const inSomeStep = (internalName: string): boolean => {
      for (let si = 0; si < steps.length; si++) {
        if (steps[si].fieldNames.indexOf(internalName) !== -1) return true;
      }
      return false;
    };
    return fields.filter((fc) => {
      if (fc.internalName === FORM_ATTACHMENTS_FIELD_INTERNAL || isFormBannerFieldConfig(fc) || isFormAlertFieldConfig(fc))
        return false;
      if (FORM_SYSTEM_LIST_METADATA_INTERNAL_NAMES.has(fc.internalName) && !inSomeStep(fc.internalName))
        return false;
      return true;
    });
  }, [fields, meta, steps]);

  const fieldsForRulesTabDisplay = useMemo(() => {
    const byName = new Map(meta.map((x) => [x.InternalName, x]));
    const list = fieldsListedForRulesTab.slice();
    const titleOf = (fc: IFormFieldConfig): string =>
      (byName.get(fc.internalName)?.Title ?? fc.internalName).trim();
    if (fieldRulesTabSort === 'asc') {
      list.sort((a, b) => titleOf(a).localeCompare(titleOf(b), 'pt'));
      return list;
    }
    if (fieldRulesTabSort === 'desc') {
      list.sort((a, b) => titleOf(b).localeCompare(titleOf(a), 'pt'));
      return list;
    }
    list.sort((a, b) => {
      const ma = byName.get(a.internalName)?.MappedType;
      const mb = byName.get(b.internalName)?.MappedType;
      const ia = fieldRulesTabMappedTypeOrderIndex(ma);
      const ib = fieldRulesTabMappedTypeOrderIndex(mb);
      if (ia !== ib) return ia - ib;
      return titleOf(a).localeCompare(titleOf(b), 'pt');
    });
    return list;
  }, [fieldsListedForRulesTab, meta, fieldRulesTabSort]);

  const cloneRulesSourceOptions = useMemo((): IDropdownOption[] => {
    if (!cloneRulesModalTarget) return [];
    return fieldsListedForRulesTab
      .filter((fc) => fc.internalName !== cloneRulesModalTarget)
      .map((fc) => {
        const mm = meta.find((m) => m.InternalName === fc.internalName);
        return {
          key: fc.internalName,
          text: `${mm?.Title ?? fc.internalName} (${fc.internalName})`,
        };
      });
  }, [cloneRulesModalTarget, fieldsListedForRulesTab, meta]);

  const applyCloneFieldRules = useCallback((): void => {
    const targetName = cloneRulesModalTarget;
    const sourceName = cloneRulesSourceKey;
    if (!targetName || !sourceName || sourceName === targetName) return;
    const srcFc = fields.find((f) => f.internalName === sourceName);
    const tgtFc = fields.find((f) => f.internalName === targetName);
    if (!srcFc || !tgtFc) return;
    const tgtMeta = meta.find((m) => m.InternalName === targetName);
    const mtp = tgtMeta?.MappedType ?? 'unknown';
    let st = fieldRuleStateFromRules(sourceName, rules);
    if (!isSetComputedAllowedForMappedType(mtp)) {
      st = {
        ...st,
        computedExpression: '',
        computedAttachmentFolderNodeId: '',
        computedLiveInEditView: false,
      };
    }
    const textVis = srcFc.textConditionalVisibility
      ? JSON.parse(JSON.stringify(srcFc.textConditionalVisibility))
      : undefined;
    const newRules = buildFieldUiRules(
      targetName,
      st,
      { textConditionalVisibility: textVis },
      { mappedType: mtp }
    );
    setRules((r) => mergeFieldRules(r, targetName, newRules));
    setFields((prev) =>
      prev.map((f) =>
        f.internalName === targetName
          ? mergeFormFieldConfigFromRulesPanel(f, buildRulesCloneFieldPatch(srcFc) as IFormFieldConfig)
          : f
      )
    );
    setCloneRulesModalTarget(null);
    setCloneRulesSourceKey(undefined);
  }, [cloneRulesModalTarget, cloneRulesSourceKey, fields, meta, rules]);

  const dismissCloneRulesModal = useCallback((): void => {
    setCloneRulesModalTarget(null);
    setCloneRulesSourceKey(undefined);
  }, []);

  const applyFieldColumnSpanForMode = useCallback(
    (fname: string, mode: TFormManagerFormMode, span: TFormFieldColumnSpan) => {
      setFields((prev) => {
        const ix = prev.findIndex((f) => f.internalName === fname);
        const applyOne = (base: IFormFieldConfig): IFormFieldConfig => {
          const next: IFormFieldConfig = { ...base };
          const by: Partial<Record<TFormManagerFormMode, TFormFieldColumnSpan>> = {
            ...(next.columnSpanByMode ?? {}),
          };
          by[mode] = span;
          next.columnSpanByMode = by;
          return next;
        };
        if (ix >= 0) return prev.map((f, j) => (j === ix ? applyOne(f) : f));
        return [...prev, applyOne({ internalName: fname })];
      });
    },
    []
  );

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

  const toggleStructurePoolSelect = useCallback((internalName: string, selected: boolean): void => {
    setStructurePoolSelected((prev) => {
      const next = new Set(prev);
      if (selected) next.add(internalName);
      else next.delete(internalName);
      return Array.from(next);
    });
  }, []);

  const toggleStructureField = useCallback((internalName: string): void => {
    setStructureFieldOpen((prev) => ({ ...prev, [internalName]: !prev[internalName] }));
  }, []);

  const placeSelectedFieldsIntoStep = useCallback((targetStepIdx: number): void => {
    const names = structurePoolSelectedRef.current.slice();
    if (!names.length) return;
    setSteps((prevSteps) => {
      let next = ensureCoreSteps(prevSteps);
      if (targetStepIdx < 0 || targetStepIdx >= next.length) return prevSteps;
      const sid = next[targetStepIdx].id;
      const isFixos = sid === FORM_FIXOS_STEP_ID;

      for (let i = 0; i < names.length; i++) {
        const internalName = names[i];
        const insertBefore = next[targetStepIdx].fieldNames.length;
        next = insertFieldNameIntoStep(next, internalName, targetStepIdx, insertBefore);
      }

      setFields((prevFields) => {
        let f = prevFields;
        for (let i = 0; i < names.length; i++) {
          const internalName = names[i];
          let exists = false;
          for (let j = 0; j < f.length; j++) {
            if (f[j].internalName === internalName) {
              exists = true;
              break;
            }
          }
          if (exists) continue;
          if (isFixos) {
            f = f.concat([
              {
                internalName,
                sectionId: sid,
                fixedPlacement: 'top',
                chromePositionMode: 'sticky',
              },
            ]);
          } else {
            f = f.concat([{ internalName, sectionId: sid }]);
          }
        }
        return fieldsAlignedToSteps(f, next);
      });

      return next;
    });
    setStructurePoolSelected([]);
  }, []);

  const addBannerField = (): void => {
    const internalName = `${FORM_BANNER_INTERNAL_PREFIX}${Date.now().toString(36)}_${Math.random()
      .toString(36)
      .slice(2, 9)}`;
    setSteps((prevSteps) => {
      const st = ensureCoreSteps(prevSteps);
      const oi = st.findIndex((x) => x.id === FORM_OCULTOS_STEP_ID);
      const idx = oi >= 0 ? oi : 0;
      const sid = st[idx].id;
      const nextSteps = st.map((s, i) =>
        i === idx ? { ...s, fieldNames: s.fieldNames.concat([internalName]) } : s
      );
      setFields((prev) => {
        const withF = prev.some((f) => f.internalName === internalName)
          ? prev
          : prev.concat([
              {
                internalName,
                sectionId: sid,
                fieldKind: 'banner',
                label: 'Banner',
                bannerImageUrl: '',
                bannerPlacement: 'inStep',
                bannerWidthPercent: 100,
                bannerHeightPx: 240,
              },
            ]);
        return fieldsAlignedToSteps(withF, nextSteps);
      });
      return nextSteps;
    });
  };

  const addAlertField = (): void => {
    const internalName = `__formAlert_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
    setSteps((prevSteps) => {
      const st = ensureCoreSteps(prevSteps);
      const oi = st.findIndex((x) => x.id === FORM_OCULTOS_STEP_ID);
      const idx = oi >= 0 ? oi : 0;
      const sid = st[idx].id;
      const nextSteps = st.map((s, i) =>
        i === idx ? { ...s, fieldNames: s.fieldNames.concat([internalName]) } : s
      );
      setFields((prev) => {
        const withF = prev.some((f) => f.internalName === internalName)
          ? prev
          : prev.concat([
              {
                internalName,
                sectionId: sid,
                fieldKind: 'alert',
                label: 'Alerta',
                alertVariant: 'info',
                alertPlacement: 'inStep',
                alertFields: [],
              },
            ]);
        return fieldsAlignedToSteps(withF, nextSteps);
      });
      return nextSteps;
    });
  };

  const addBannerFieldToFixos = (): void => {
    const internalName = `${FORM_BANNER_INTERNAL_PREFIX}${Date.now().toString(36)}_${Math.random()
      .toString(36)
      .slice(2, 9)}`;
    setSteps((prevSteps) => {
      const st = ensureCoreSteps(prevSteps);
      const fi = st.findIndex((x) => x.id === FORM_FIXOS_STEP_ID);
      const idx = fi >= 0 ? fi : 0;
      const sid = st[idx].id;
      const nextSteps = st.map((s, i) =>
        i === idx ? { ...s, fieldNames: s.fieldNames.concat([internalName]) } : s
      );
      setFields((prev) => {
        const withF = prev.some((f) => f.internalName === internalName)
          ? prev
          : prev.concat([
              {
                internalName,
                sectionId: sid,
                fieldKind: 'banner',
                label: 'Banner',
                bannerImageUrl: '',
                bannerPlacement: 'inStep',
                bannerWidthPercent: 100,
                bannerHeightPx: 240,
                fixedPlacement: 'top',
                chromePositionMode: 'sticky',
              },
            ]);
        return fieldsAlignedToSteps(withF, nextSteps);
      });
      return nextSteps;
    });
  };

  const addAlertFieldToFixos = (): void => {
    const internalName = `__formAlert_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`;
    setSteps((prevSteps) => {
      const st = ensureCoreSteps(prevSteps);
      const fi = st.findIndex((x) => x.id === FORM_FIXOS_STEP_ID);
      const idx = fi >= 0 ? fi : 0;
      const sid = st[idx].id;
      const nextSteps = st.map((s, i) =>
        i === idx ? { ...s, fieldNames: s.fieldNames.concat([internalName]) } : s
      );
      setFields((prev) => {
        const withF = prev.some((f) => f.internalName === internalName)
          ? prev
          : prev.concat([
              {
                internalName,
                sectionId: sid,
                fieldKind: 'alert',
                label: 'Alerta',
                alertVariant: 'info',
                alertPlacement: 'inStep',
                fixedPlacement: 'top',
                chromePositionMode: 'sticky',
                alertFields: [],
              },
            ]);
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
    if (actionLogCaptureEnabled) {
      if (
        !actionLogListTitle.trim() ||
        !actionLogFieldInternalName.trim() ||
        !actionLogSourceListLookupFieldInternalName.trim()
      ) {
        setErr(
          'Para captação de logs ativa, indique a lista de log, o campo multilinhas da ação e o Lookup de vínculo à lista principal.'
        );
        return;
      }
    }
    if (attachmentStorageKind === 'documentLibrary') {
      if (!attachmentLibLibraryTitle.trim() || !attachmentLibLookupField.trim()) {
        setErr(
          'No modo Biblioteca, indique a biblioteca de documentos e o campo Lookup que relaciona com a lista principal.'
        );
        return;
      }
    }
    const withRules = mergeAttachmentUiRule(rules, {
      minCount: numOpt(attachMin),
      maxCount: numOpt(attachMax),
      message: attachMsg,
      allowedFileExtensions: attachAllowedExt.length ? attachAllowedExt : undefined,
    });
    const sectionsOut = sectionsFromSteps(steps);
    const stepNavigation = buildStepNavigationForSave(
      stepRequireFilledToAdvance,
      stepFullValOnAdvance,
      stepAllowBackWithoutVal
    );
    const actionLogPayload: IFormManagerActionLogConfig = {};
    if (actionLogCaptureEnabled) actionLogPayload.captureEnabled = true;
    if (actionLogListTitle.trim()) actionLogPayload.listTitle = actionLogListTitle.trim();
    if (actionLogFieldInternalName.trim()) {
      actionLogPayload.actionFieldInternalName = actionLogFieldInternalName.trim();
    }
    if (actionLogSourceListLookupFieldInternalName.trim()) {
      actionLogPayload.sourceListLookupFieldInternalName = actionLogSourceListLookupFieldInternalName.trim();
    }
    const descEntries = Object.entries(actionLogDescById).filter(([, v]) => (v || '').trim());
    if (descEntries.length) {
      actionLogPayload.descriptionsHtmlByButtonId = Object.fromEntries(descEntries);
    }
    const paletteEntries = Object.entries(actionLogPaletteSlotById).filter(
      ([, slot]) => slot && slot !== 'themePrimary'
    );
    if (paletteEntries.length) {
      actionLogPayload.descriptionPaletteSlotByButtonId = Object.fromEntries(paletteEntries);
    }
    if (actionLogAutomaticChangesOnUpdate) actionLogPayload.automaticChangesOnUpdate = true;
    const hasActionLog = !!(
      actionLogCaptureEnabled ||
      actionLogPayload.listTitle ||
      actionLogPayload.actionFieldInternalName ||
      actionLogPayload.sourceListLookupFieldInternalName ||
      actionLogPayload.descriptionsHtmlByButtonId ||
      actionLogPayload.descriptionPaletteSlotByButtonId ||
      actionLogPayload.automaticChangesOnUpdate
    );
    const attachmentLibStashed = attachmentLibraryFromPanelState(
      attachmentLibLibraryTitle,
      attachmentLibLookupField,
      attachmentLibFolderTree
    );
    const attachmentLibPayload =
      attachmentStorageKind === 'documentLibrary'
        ? {
            libraryTitle: attachmentLibLibraryTitle.trim(),
            sourceListLookupFieldInternalName: attachmentLibLookupField.trim(),
            ...(attachmentLibFolderTree.length ? { folderTree: attachmentLibFolderTree } : {}),
          }
        : undefined;
    const raw: IFormManagerConfig = {
      sections: sectionsOut,
      fields,
      rules: withRules,
      steps,
      formRootWidthMode,
      formRootWidthPercent: clampFormRootPercentInput(formRootWidthPercent),
      formRootHorizontalAlign,
      ...(clampFormRootPaddingInput(formRootPaddingPx) > 0
        ? { formRootPaddingPx: clampFormRootPaddingInput(formRootPaddingPx) }
        : {}),
      ...(dynamicHelp ? { dynamicHelp } : {}),
      ...(managerColumnFields.length ? { managerColumnFields } : {}),
      ...(customButtons.length
        ? {
            customButtons: customButtons.map((b) => ({
              ...b,
              actions: b.actions ?? [],
            })),
          }
        : {}),
      ...(customButtonsBarVertical === 'top' ? { customButtonsBarVertical: 'top' as const } : {}),
      ...(customButtonsBarHorizontal === 'right' ? { customButtonsBarHorizontal: 'right' as const } : {}),
      stepLayout,
      ...(stepAccentPaletteSlot ? { stepAccentPaletteSlot } : {}),
      ...(stepNavButtons && stepNavButtons !== 'fluent' ? { stepNavButtons } : {}),
      ...(formDataLoadingKind && formDataLoadingKind !== 'spinner' ? { formDataLoadingKind } : {}),
      ...(defaultSubmitLoadingKind && defaultSubmitLoadingKind !== 'overlay'
        ? { defaultSubmitLoadingKind }
        : {}),
      ...(stepNavigation ? { stepNavigation } : {}),
      ...(attachmentUploadLayout && attachmentUploadLayout !== 'default' ? { attachmentUploadLayout } : {}),
      ...(attachmentFilePreview && attachmentFilePreview !== 'nameAndSize' ? { attachmentFilePreview } : {}),
      ...(attachmentStorageKind === 'documentLibrary' && attachmentLibPayload
        ? { attachmentStorageKind: 'documentLibrary', attachmentLibrary: attachmentLibPayload }
        : {
            ...(attachmentStorageKind === 'itemAttachments' ? { attachmentStorageKind: 'itemAttachments' } : {}),
            ...(attachmentLibStashed ? { attachmentLibrary: attachmentLibStashed } : {}),
          }),
      ...(hasActionLog ? { actionLog: actionLogPayload } : {}),
      ...(historyLayoutKind && historyLayoutKind !== 'list' ? { historyLayoutKind } : {}),
      ...(historyEnabled
        ? {
            historyEnabled: true,
            ...(historyPresentationKind !== 'panel' ? { historyPresentationKind } : {}),
            ...(historyButtonKind !== 'text' ? { historyButtonKind } : {}),
            historyButtonLabel: (historyButtonLabel.trim() || 'Histórico').slice(0, 120),
            historyButtonIcon: (historyButtonIcon.trim() || 'History').slice(0, 80),
            ...(historyPanelSubtitle.trim() ? { historyPanelSubtitle: historyPanelSubtitle.trim() } : {}),
            ...(historyGroupTitles.length ? { historyGroupTitles: historyGroupTitles.slice() } : {}),
          }
        : {}),
      linkedChildForms: linkedChildForms.map(cloneLinkedChildFormConfig),
      permissionBreak,
    };
    const sanitized = sanitizeFormManagerConfig(raw);
    if (!sanitized) {
      setErr('Configuração inválida.');
      return;
    }
    const jsonStr = JSON.stringify(sanitized, null, 2);
    void (async (): Promise<void> => {
      try {
        if (typeof navigator !== 'undefined' && navigator.clipboard?.writeText) {
          await navigator.clipboard.writeText(jsonStr);
        }
      } catch {
        /* ignore */
      }
    })();
    onSave(sanitized);
    onDismiss();
  };

  const addStep = (): void => {
    setSteps((prev) => [...prev, { id: newId('step'), title: 'Nova etapa', fieldNames: [] }]);
  };

  const updateStep = (i: number, patch: Partial<IFormStepConfig>): void => {
    setSteps((prev) => prev.map((s, j) => (j === i ? { ...s, ...patch } : s)));
  };

  const patchStepById = useCallback((id: string, patch: Partial<IFormStepConfig>): void => {
    setSteps((prev) => prev.map((s) => (s.id === id ? { ...s, ...patch } : s)));
  }, []);

  const patchStepVisibilityWhenRow = useCallback(
    (stepId: string, ri: number, partial: Partial<IWhenUi>): void => {
      setSteps((prev) =>
        prev.map((st) => {
          if (st.id !== stepId) return st;
          const { combiner, rows } = parseButtonWhenToRows(st.showStepWhen, meta);
          const nextRows = rows.map((r, i) => (i === ri ? { ...r, ...partial } : r));
          const showStepWhen = buildButtonWhenFromRows(combiner, nextRows);
          return { ...st, showStepWhen };
        })
      );
    },
    [meta]
  );

  const setStepVisibilityWhenCombiner = useCallback((stepId: string, combiner: 'all' | 'any'): void => {
    setSteps((prev) =>
      prev.map((st) => {
        if (st.id !== stepId) return st;
        const { rows } = parseButtonWhenToRows(st.showStepWhen, meta);
        const showStepWhen = buildButtonWhenFromRows(combiner, rows);
        return { ...st, showStepWhen };
      })
    );
  }, [meta]);

  const addStepVisibilityWhenRow = useCallback(
    (stepId: string): void => {
      setSteps((prev) =>
        prev.map((st) => {
          if (st.id !== stepId) return st;
          const { combiner, rows } = parseButtonWhenToRows(st.showStepWhen, meta);
          const nextRows = rows.concat([defaultWhenUi(meta)]);
          const showStepWhen = buildButtonWhenFromRows(combiner, nextRows);
          return { ...st, showStepWhen };
        })
      );
    },
    [meta]
  );

  const removeStepVisibilityWhenRow = useCallback(
    (stepId: string, ri: number): void => {
      setSteps((prev) =>
        prev.map((st) => {
          if (st.id !== stepId) return st;
          const { combiner, rows } = parseButtonWhenToRows(st.showStepWhen, meta);
          if (rows.length <= 1) return st;
          const nextRows = rows.filter((_, i) => i !== ri);
          const showStepWhen = buildButtonWhenFromRows(combiner, nextRows);
          return { ...st, showStepWhen };
        })
      );
    },
    [meta]
  );

  const reorderStep = (from: number, to: number): void => {
    setSteps((prev) => {
      const n = pinCoreStepsOrder(reorderByIndex(prev, from, to));
      setFields((flds) => fieldsAlignedToSteps(flds, n));
      return n;
    });
  };

  const removeStep = (i: number): void => {
    setSteps((prev) => {
      if (prev.length <= 1) return prev;
      const removed = prev[i];
      if (!removed) return prev;
      if (removed.id === FORM_OCULTOS_STEP_ID || removed.id === FORM_FIXOS_STEP_ID) return prev;
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

  const patchButtonWhenRow = (bi: number, ri: number, partial: Partial<IWhenUi>): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) => {
        if (j !== bi) return b;
        const { combiner, rows } = parseButtonWhenToRows(b.when, meta);
        const nextRows = rows.map((r, i) => (i === ri ? { ...r, ...partial } : r));
        const when = buildButtonWhenFromRows(combiner, nextRows);
        return { ...b, when };
      })
    );
  };

  const setButtonWhenCombiner = (bi: number, combiner: 'all' | 'any'): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) => {
        if (j !== bi) return b;
        const { rows } = parseButtonWhenToRows(b.when, meta);
        const when = buildButtonWhenFromRows(combiner, rows);
        return { ...b, when };
      })
    );
  };

  const addButtonWhenRow = (bi: number): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) => {
        if (j !== bi) return b;
        const { combiner, rows } = parseButtonWhenToRows(b.when, meta);
        const nextRows = rows.concat([defaultWhenUi(meta)]);
        const when = buildButtonWhenFromRows(combiner, nextRows);
        return { ...b, when };
      })
    );
  };

  const removeButtonWhenRow = (bi: number, ri: number): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) => {
        if (j !== bi) return b;
        const { combiner, rows } = parseButtonWhenToRows(b.when, meta);
        if (rows.length <= 1) return b;
        const nextRows = rows.filter((_, i) => i !== ri);
        const when = buildButtonWhenFromRows(combiner, nextRows);
        return { ...b, when };
      })
    );
  };

  const patchButtonActionCondition = (bi: number, ai: number, when: TFormConditionNode | undefined): void => {
    setCustomButtons((prev) =>
      prev.map((b, j) => {
        if (j !== bi) return b;
        const acts = b.actions.map((a, k) => {
          if (k !== ai) return a;
          if (!when) {
            const { when: _rm, ...rest } = a as TFormButtonAction & { when?: TFormConditionNode };
            return rest as TFormButtonAction;
          }
          return { ...a, when } as TFormButtonAction;
        });
        return { ...b, actions: acts };
      })
    );
  };

  const removeCustomButton = (i: number): void => {
    setCustomButtons((prev) => prev.filter((_, j) => j !== i));
  };

  const cloneCustomButton = (i: number): void => {
    setCustomButtons((prev) => {
      const src = prev[i];
      if (!src) return prev;
      const copy = JSON.parse(JSON.stringify(src)) as IFormCustomButtonConfig;
      copy.id = newId('btn');
      const base = (copy.label || '').trim() || src.label || 'Botão';
      copy.label = `${base} (cópia)`;
      const next = prev.slice();
      next.splice(i + 1, 0, copy);
      return next;
    });
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
      allowedFileExtensions: attachAllowedExt.length ? attachAllowedExt : undefined,
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
    const actionLogPreview: IFormManagerActionLogConfig = {};
    if (actionLogCaptureEnabled) actionLogPreview.captureEnabled = true;
    if (actionLogListTitle.trim()) actionLogPreview.listTitle = actionLogListTitle.trim();
    if (actionLogFieldInternalName.trim()) {
      actionLogPreview.actionFieldInternalName = actionLogFieldInternalName.trim();
    }
    if (actionLogSourceListLookupFieldInternalName.trim()) {
      actionLogPreview.sourceListLookupFieldInternalName = actionLogSourceListLookupFieldInternalName.trim();
    }
    const descPrev = Object.entries(actionLogDescById).filter(([, v]) => (v || '').trim());
    if (descPrev.length) {
      actionLogPreview.descriptionsHtmlByButtonId = Object.fromEntries(descPrev);
    }
    const palettePrev = Object.entries(actionLogPaletteSlotById).filter(
      ([, slot]) => slot && slot !== 'themePrimary'
    );
    if (palettePrev.length) {
      actionLogPreview.descriptionPaletteSlotByButtonId = Object.fromEntries(palettePrev);
    }
    if (actionLogAutomaticChangesOnUpdate) actionLogPreview.automaticChangesOnUpdate = true;
    const hasActionLogPreview = !!(
      actionLogCaptureEnabled ||
      actionLogPreview.listTitle ||
      actionLogPreview.actionFieldInternalName ||
      actionLogPreview.sourceListLookupFieldInternalName ||
      actionLogPreview.descriptionsHtmlByButtonId ||
      actionLogPreview.descriptionPaletteSlotByButtonId ||
      actionLogPreview.automaticChangesOnUpdate
    );
    const attachmentLibStashedPreview = attachmentLibraryFromPanelState(
      attachmentLibLibraryTitle,
      attachmentLibLookupField,
      attachmentLibFolderTree
    );
    const attachmentLibPreview =
      attachmentStorageKind === 'documentLibrary'
        ? {
            libraryTitle: attachmentLibLibraryTitle.trim(),
            sourceListLookupFieldInternalName: attachmentLibLookupField.trim(),
            ...(attachmentLibFolderTree.length ? { folderTree: attachmentLibFolderTree } : {}),
          }
        : undefined;
    const raw: IFormManagerConfig = {
      sections: sectionsOut,
      fields,
      rules: withRules,
      steps,
      formRootWidthMode,
      formRootWidthPercent: clampFormRootPercentInput(formRootWidthPercent),
      formRootHorizontalAlign,
      ...(clampFormRootPaddingInput(formRootPaddingPx) > 0
        ? { formRootPaddingPx: clampFormRootPaddingInput(formRootPaddingPx) }
        : {}),
      ...(dynamicHelp ? { dynamicHelp } : {}),
      ...(managerColumnFields.length ? { managerColumnFields } : {}),
      ...(customButtons.length
        ? {
            customButtons: customButtons.map((b) => ({
              ...b,
              actions: b.actions ?? [],
            })),
          }
        : {}),
      ...(customButtonsBarVertical === 'top' ? { customButtonsBarVertical: 'top' as const } : {}),
      ...(customButtonsBarHorizontal === 'right' ? { customButtonsBarHorizontal: 'right' as const } : {}),
      stepLayout,
      ...(stepAccentPaletteSlot ? { stepAccentPaletteSlot } : {}),
      ...(stepNavButtons && stepNavButtons !== 'fluent' ? { stepNavButtons } : {}),
      ...(formDataLoadingKind && formDataLoadingKind !== 'spinner' ? { formDataLoadingKind } : {}),
      ...(defaultSubmitLoadingKind && defaultSubmitLoadingKind !== 'overlay'
        ? { defaultSubmitLoadingKind }
        : {}),
      ...(stepNavigation ? { stepNavigation } : {}),
      ...(attachmentUploadLayout && attachmentUploadLayout !== 'default' ? { attachmentUploadLayout } : {}),
      ...(attachmentFilePreview && attachmentFilePreview !== 'nameAndSize' ? { attachmentFilePreview } : {}),
      ...(attachmentStorageKind === 'documentLibrary' && attachmentLibPreview
        ? { attachmentStorageKind: 'documentLibrary', attachmentLibrary: attachmentLibPreview }
        : {
            ...(attachmentStorageKind === 'itemAttachments' ? { attachmentStorageKind: 'itemAttachments' } : {}),
            ...(attachmentLibStashedPreview ? { attachmentLibrary: attachmentLibStashedPreview } : {}),
          }),
      ...(hasActionLogPreview ? { actionLog: actionLogPreview } : {}),
      ...(historyLayoutKind && historyLayoutKind !== 'list' ? { historyLayoutKind } : {}),
      ...(historyEnabled
        ? {
            historyEnabled: true,
            ...(historyPresentationKind !== 'panel' ? { historyPresentationKind } : {}),
            ...(historyButtonKind !== 'text' ? { historyButtonKind } : {}),
            historyButtonLabel: (historyButtonLabel.trim() || 'Histórico').slice(0, 120),
            historyButtonIcon: (historyButtonIcon.trim() || 'History').slice(0, 80),
            ...(historyPanelSubtitle.trim() ? { historyPanelSubtitle: historyPanelSubtitle.trim() } : {}),
            ...(historyGroupTitles.length ? { historyGroupTitles: historyGroupTitles.slice() } : {}),
          }
        : {}),
      linkedChildForms: linkedChildForms.map(cloneLinkedChildFormConfig),
      permissionBreak,
    };
    return JSON.stringify(raw, null, 2);
  }, [
    fields,
    rules,
    steps,
    helpJson,
    managerColumnFields,
    customButtons,
    linkedChildForms,
    permissionBreak,
    stepLayout,
    stepAccentPaletteSlot,
    stepNavButtons,
    formDataLoadingKind,
    defaultSubmitLoadingKind,
    value,
    attachmentUploadLayout,
    attachmentFilePreview,
    attachmentStorageKind,
    attachmentLibLibraryTitle,
    attachmentLibLookupField,
    attachmentLibFolderTree,
    stepRequireFilledToAdvance,
    stepFullValOnAdvance,
    stepAllowBackWithoutVal,
    formRootWidthMode,
    formRootWidthPercent,
    formRootHorizontalAlign,
    formRootPaddingPx,
    attachMin,
    attachMax,
    attachMsg,
    attachAllowedExt,
    actionLogCaptureEnabled,
    actionLogListTitle,
    actionLogFieldInternalName,
    actionLogSourceListLookupFieldInternalName,
    actionLogDescById,
    actionLogPaletteSlotById,
    historyEnabled,
    historyPresentationKind,
    historyLayoutKind,
    historyButtonKind,
    historyButtonLabel,
    historyButtonIcon,
    historyPanelSubtitle,
    historyGroupTitles,
    customButtonsBarVertical,
    customButtonsBarHorizontal,
  ]);

  const previewConfigJsonRef = useRef(previewConfigJson);
  previewConfigJsonRef.current = previewConfigJson;
  useEffect(() => {
    if (jsonOpen) {
      setJsonPanelText(previewConfigJsonRef.current);
      setJsonPanelErr(undefined);
    }
  }, [jsonOpen]);

  useEffect(() => {
    if (stepVisibilityPanelStepId === null) return;
    if (!steps.some((s) => s.id === stepVisibilityPanelStepId)) {
      setStepVisibilityPanelStepId(null);
    }
  }, [steps, stepVisibilityPanelStepId]);

  const applyJsonFromPanel = useCallback(() => {
    setJsonPanelErr(undefined);
    try {
      const parsed = JSON.parse(jsonPanelText) as unknown;
      const sanitized = sanitizeFormManagerConfig(parsed);
      if (!sanitized) {
        setJsonPanelErr('JSON inválido ou estrutura não reconhecida.');
        return;
      }
      hydrateFromFormManagerConfig(sanitized);
      setJsonPanelText(JSON.stringify(sanitized, null, 2));
    } catch (e) {
      setJsonPanelErr(e instanceof Error ? e.message : String(e));
    }
  }, [jsonPanelText, hydrateFromFormManagerConfig]);

  return (
    <>
      <Panel
      isOpen={isOpen}
      type={PanelType.large}
      headerText="Configurar formulário e regras"
      onDismiss={onDismiss}
      onOuterClick={(ev) => {
        const t = ev?.target;
        if (t instanceof Element && t.closest(`[${FORM_FIELD_RULES_MENTION_PORTAL_ATTR}]`)) return;
        onDismiss();
      }}
      isFooterAtBottom
      onRenderFooterContent={() => (
        <Stack
          horizontal
          horizontalAlign="start"
          verticalAlign="center"
          tokens={{ childrenGap: 8 }}
          wrap
          styles={{ root: { width: '100%' } }}
        >
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
      )}
      styles={{
        main: {
          display: 'flex',
          flexDirection: 'column',
          maxHeight: '100%',
          overflow: 'hidden',
        },
        content: {
          flex: 1,
          minHeight: 0,
          overflowY: 'auto',
          overflowX: 'hidden',
          WebkitOverflowScrolling: 'touch',
        },
        footer: {
          flexShrink: 0,
          borderTop: '1px solid #edebe9',
          paddingTop: 16,
          paddingBottom: 16,
          background: '#faf9f8',
        },
      }}
    >
      {loading && <Spinner label="Campos da lista..." />}
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Link onClick={() => setJsonOpen(true)}>JSON (ver / colar)</Link>
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
            {structurePoolSelected.length > 0 && (
              <Stack
                horizontal
                verticalAlign="center"
                wrap
                tokens={{ childrenGap: 10 }}
                styles={{
                  root: {
                    padding: '10px 12px',
                    borderRadius: 4,
                    border: '1px solid #c8c6c4',
                    background: '#fff4ce',
                  },
                }}
              >
                <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                  {structurePoolSelected.length} campo(s) selecionado(s)
                </Text>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Escolha a etapa de destino:
                </Text>
                <IconButton
                  className={poolBulkMoveIconClassName}
                  iconProps={{ iconName: 'Forward' }}
                  title="Mover para etapa…"
                  ariaLabel="Abrir lista de etapas para mover os campos selecionados"
                  menuProps={{
                    items: steps.map((step, stepIdx) => ({
                      key: step.id,
                      text: structureStepMenuLabel(step, stepIdx),
                      onClick: () => placeSelectedFieldsIntoStep(stepIdx),
                    })),
                  }}
                />
                <DefaultButton text="Limpar seleção" onClick={() => setStructurePoolSelected([])} />
              </Stack>
            )}
            <FormManagerCollapseSection
              title="Layout do formulário (vista)"
              isOpen={isEstruturaOpen(ESTRUTURA_COLLAPSE_IDS.formLayout)}
              onToggle={() => toggleEstruturaSection(ESTRUTURA_COLLAPSE_IDS.formLayout)}
            >
           
              <Dropdown
                label="Largura"
                options={FORM_ROOT_WIDTH_OPTIONS}
                selectedKey={formRootWidthMode}
                onChange={(_, o) => o && setFormRootWidthMode(String(o.key) as TFormRootWidthMode)}
              />
              {formRootWidthMode === 'percent' && (
                <TextField
                  label="Percentagem da largura (1–100)"
                  value={formRootWidthPercent}
                  onChange={(_, v) => setFormRootWidthPercent(v ?? '')}
                  description="Ex.: 80 para ocupar 80% da largura disponível."
                />
              )}
              <Dropdown
                label="Alinhamento horizontal"
                options={FORM_ROOT_ALIGN_OPTIONS}
                selectedKey={formRootHorizontalAlign}
                onChange={(_, o) =>
                  o && setFormRootHorizontalAlign(String(o.key) as TFormRootHorizontalAlign)
                }
              />
              <TextField
                label="Padding (px)"
                value={formRootPaddingPx}
                onChange={(_, v) => setFormRootPaddingPx(v ?? '')}
                description="Espaço interior em torno do conteúdo (1–160). Vazio = sem padding extra."
              />
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Navegação entre etapas (formulário)"
              isOpen={isEstruturaOpen(ESTRUTURA_COLLAPSE_IDS.stepNav)}
              onToggle={() => toggleEstruturaSection(ESTRUTURA_COLLAPSE_IDS.stepNav)}
            >
          
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
   
            </FormManagerCollapseSection>
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
                    draggable={st.id !== FORM_OCULTOS_STEP_ID && st.id !== FORM_FIXOS_STEP_ID}
                    title={
                      st.id === FORM_OCULTOS_STEP_ID
                        ? 'Ocultos permanece sempre na primeira posição'
                        : st.id === FORM_FIXOS_STEP_ID
                          ? 'Fixos permanece após Ocultos'
                          : 'Arrastar etapa'
                    }
                    onDragStart={(e) => {
                      if (st.id === FORM_OCULTOS_STEP_ID || st.id === FORM_FIXOS_STEP_ID) return;
                      e.dataTransfer.setData('text/plain', dragPayload(DND_STEP, si));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    style={{
                      cursor:
                        st.id === FORM_OCULTOS_STEP_ID || st.id === FORM_FIXOS_STEP_ID ? 'default' : 'grab',
                      display: 'flex',
                      alignItems: 'center',
                      color: '#605e5c',
                    }}
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
                  {st.id === FORM_OCULTOS_STEP_ID ? (
                    <Text variant="small" styles={{ root: { color: '#605e5c', alignSelf: 'center' } }}>
                      Não entra no passador (reserva de campos)
                    </Text>
                  ) : (
                    <Stack tokens={{ childrenGap: 6 }} styles={{ root: { alignItems: 'flex-start' } }}>
                      {st.id === FORM_FIXOS_STEP_ID && (
                        <Text variant="small" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
                          Topo ou rodapé fixo
                        </Text>
                      )}
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        {formatStepModesHint(st)}
                        {st.showStepWhen ? ` · ${summarizeConditionTreePt(st.showStepWhen)}` : ''}
                      </Text>
                      <DefaultButton text="Configurar" onClick={() => setStepVisibilityPanelStepId(st.id)} />
                    </Stack>
                  )}
                  {st.id !== FORM_OCULTOS_STEP_ID && st.id !== FORM_FIXOS_STEP_ID && (
                    <DefaultButton text="Remover etapa" onClick={() => removeStep(si)} />
                  )}
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
                    let fcRow: IFormFieldConfig | undefined;
                    for (let fi = 0; fi < fields.length; fi++) {
                      if (fields[fi].internalName === fname) {
                        fcRow = fields[fi];
                        break;
                      }
                    }
                    const isAlert = fcRow !== undefined && isFormAlertFieldConfig(fcRow);
                    const isBanner = fcRow !== undefined && isFormBannerFieldConfig(fcRow);
                    const reqStyles = requiredFieldRowStyles(mm, steps, fields);
                    if (isAlert && fcRow) {
                      const alertWhenUi = alertWhenUiFromNode(fcRow.alertWhen, meta);
                      const alertWhenEnabled = fcRow.alertWhen !== undefined;
                      const alertOpen = structureFieldOpen[fname] === true;
                      const currentAlertFields =
                        (fcRow as IFormFieldConfig & { alertFields?: string[] }).alertFields ?? [];
                      return (
                        <Stack
                          key={fname}
                          tokens={{ childrenGap: 8 }}
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
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
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
                            <Checkbox
                              checked={structurePoolSelected.indexOf(fname) !== -1}
                              onChange={(_, c) => toggleStructurePoolSelect(fname, !!c)}
                              styles={{ text: { display: 'none' } }}
                              title="Selecionar para mover em conjunto"
                            />
                            <IconButton
                              iconProps={{ iconName: alertOpen ? 'ChevronDown' : 'ChevronRight' }}
                              title={alertOpen ? 'Fechar configurações' : 'Abrir configurações'}
                              ariaLabel={alertOpen ? 'Fechar configurações' : 'Abrir configurações'}
                              onClick={() => toggleStructureField(fname)}
                              styles={{ root: { width: 30, height: 30 } }}
                            />
                            <Text styles={{ root: { fontWeight: 600, minWidth: 80 } }}>Alerta</Text>
                            <Text variant="small" styles={{ root: { color: '#605e5c', flex: '1 1 200px' } }}>
                              {fname} · {alertOpen ? 'configurações visíveis' : 'clique para configurar'}
                            </Text>
                            <DefaultButton text="Remover" onClick={() => removeField(fname)} />
                          </Stack>
                          {alertOpen ? (
                            <>
                              <Stack horizontal tokens={{ childrenGap: 12 }} wrap styles={{ root: { width: '100%' } }}>
                                <TextField
                                  label="Título"
                                  value={fcRow.alertTitle ?? ''}
                                  onChange={(_, v) => {
                                    const t = v ?? '';
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname ? { ...f, alertTitle: t.trim() || undefined } : f
                                      )
                                    );
                                  }}
                                />
                                <TextField
                                  label="Mensagem"
                                  multiline
                                  rows={3}
                                  styles={{ root: { minWidth: 260, flex: '1 1 360px' } }}
                                  value={fcRow.alertMessage ?? ''}
                                  onChange={(_, v) => {
                                    const t = v ?? '';
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname ? { ...f, alertMessage: t.trim() || undefined } : f
                                      )
                                    );
                                  }}
                                />
                                <Dropdown
                                  label="Campos no alerta"
                                  placeholder="Selecione um ou vários campos"
                                  multiSelect
                                  options={fieldOptions}
                                  selectedKeys={currentAlertFields}
                                  styles={{ root: { minWidth: 240, flex: '1 1 320px' } }}
                                  onChange={(_, o) => {
                                    if (!o) return;
                                    const key = String(o.key);
                                    const next = new Set(currentAlertFields);
                                    if (o.selected === true) next.add(key);
                                    else next.delete(key);
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname ? { ...f, alertFields: Array.from(next) } : f
                                      )
                                    );
                                  }}
                                />
                                <Dropdown
                                  label="Tipo"
                                  options={[
                                    { key: 'info', text: 'Informação' },
                                    { key: 'success', text: 'Sucesso' },
                                    { key: 'warning', text: 'Aviso' },
                                    { key: 'error', text: 'Erro' },
                                  ]}
                                  selectedKey={resolveAlertVariant(fcRow)}
                                  onChange={(_, o) => {
                                    if (!o) return;
                                    const k = String(o.key) as TFormAlertVariant;
                                    setFields((prev) =>
                                      prev.map((f) => (f.internalName === fname ? { ...f, alertVariant: k } : f))
                                    );
                                  }}
                                />
                              </Stack>
                              <Checkbox
                                label="Mostrar só quando a condição abaixo for verdadeira"
                                checked={alertWhenEnabled}
                                onChange={(_, c) => {
                                  if (c) {
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname
                                          ? {
                                              ...f,
                                              alertWhen: whenUiToNode(alertWhenUi),
                                            }
                                          : f
                                      )
                                    );
                                  } else {
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname ? { ...f, alertWhen: undefined } : f
                                      )
                                    );
                                  }
                                }}
                              />
                              {alertWhenEnabled && (
                                <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
                                  <Dropdown
                                    label="Campo"
                                    options={fieldOptions}
                                    selectedKey={alertWhenUi.field || undefined}
                                    onChange={(_, o) =>
                                      o &&
                                      setFields((prev) =>
                                        prev.map((f) =>
                                          f.internalName === fname
                                            ? { ...f, alertWhen: whenUiToNode({ ...alertWhenUi, field: String(o.key) }) }
                                            : f
                                        )
                                      )
                                    }
                                  />
                                  <Dropdown
                                    label="Operador"
                                    options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                                    selectedKey={alertWhenUi.op}
                                    onChange={(_, o) =>
                                      o &&
                                      setFields((prev) =>
                                        prev.map((f) =>
                                          f.internalName === fname
                                            ? { ...f, alertWhen: whenUiToNode({ ...alertWhenUi, op: o.key as TFormConditionOp }) }
                                            : f
                                        )
                                      )
                                    }
                                  />
                                  <Dropdown
                                    label="Comparar com"
                                    options={[
                                      { key: 'literal', text: 'Texto fixo' },
                                      { key: 'field', text: 'Outro campo' },
                                      { key: 'token', text: 'Token' },
                                    ]}
                                    selectedKey={alertWhenUi.compareKind}
                                    onChange={(_, o) =>
                                      o &&
                                      setFields((prev) =>
                                        prev.map((f) =>
                                          f.internalName === fname
                                            ? {
                                                ...f,
                                                alertWhen: whenUiToNode({
                                                  ...alertWhenUi,
                                                  compareKind: o.key as IWhenUi['compareKind'],
                                                }),
                                              }
                                            : f
                                        )
                                      )
                                    }
                                  />
                                  <TextField
                                    label="Valor"
                                    value={alertWhenUi.compareValue}
                                    onChange={(_, v) =>
                                      setFields((prev) =>
                                        prev.map((f) =>
                                          f.internalName === fname
                                            ? {
                                                ...f,
                                                alertWhen: whenUiToNode({
                                                  ...alertWhenUi,
                                                  compareValue: v ?? '',
                                                }),
                                              }
                                            : f
                                        )
                                      )
                                    }
                                    disabled={
                                      alertWhenUi.op === 'isEmpty' ||
                                      alertWhenUi.op === 'isFilled' ||
                                      alertWhenUi.op === 'isTrue' ||
                                      alertWhenUi.op === 'isFalse'
                                    }
                                  />
                                </Stack>
                              )}
                              <Stack horizontal tokens={{ childrenGap: 12 }} wrap styles={{ root: { width: '100%' } }}>
                                <TextField
                                  label="Ícone"
                                  description="Opcional. Nome de ícone Fluent."
                                  styles={{ root: { minWidth: 180, maxWidth: 260 } }}
                                  value={fcRow.alertIconName ?? ''}
                                  onChange={(_, v) => {
                                    const t = v ?? '';
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname ? { ...f, alertIconName: t.trim() || undefined } : f
                                      )
                                    );
                                  }}
                                />
                                <Checkbox
                                  label="Destacar visualmente"
                                  checked={fcRow.alertEmphasized === true}
                                  onChange={(_, c) =>
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname ? { ...f, alertEmphasized: !!c } : f
                                      )
                                    )
                                  }
                                />
                                <Checkbox
                                  label="Fechável"
                                  checked={fcRow.alertDismissible === true}
                                  onChange={(_, c) =>
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname ? { ...f, alertDismissible: !!c } : f
                                      )
                                    )
                                  }
                                />
                                <Dropdown
                                  label="Posição no formulário"
                                  options={BANNER_PLACEMENT_DROPDOWN_OPTIONS}
                                  selectedKey={resolveAlertPlacement(fcRow)}
                                  onChange={(_, o) => {
                                    if (!o) return;
                                    const k = String(o.key) as TFormBannerPlacement;
                                    setFields((prev) =>
                                      prev.map((f) => (f.internalName === fname ? { ...f, alertPlacement: k } : f))
                                    );
                                  }}
                                />
                              </Stack>
                              {resolveAlertPlacement(fcRow) !== 'inStep' && (
                                <Stack horizontal tokens={{ childrenGap: 12 }} wrap styles={{ root: { width: '100%' } }}>
                                  <Dropdown
                                    label="Zona fixa"
                                    options={FIXED_CHROME_PLACEMENT_OPTIONS}
                                    selectedKey={resolveFixedPlacement(fcRow)}
                                    onChange={(_, o) => {
                                      if (!o) return;
                                      const k = String(o.key) as TFixedChromePlacement;
                                      setFields((prev) =>
                                        prev.map((f) =>
                                          f.internalName === fname ? { ...f, fixedPlacement: k } : f
                                        )
                                      );
                                    }}
                                  />
                                  <Dropdown
                                    label="Posicionamento"
                                    options={CHROME_POSITION_MODE_OPTIONS}
                                    selectedKey={resolveChromePositionMode(fcRow)}
                                    onChange={(_, o) => {
                                      if (!o) return;
                                      const k = String(o.key) as TChromePositionMode;
                                      setFields((prev) =>
                                        prev.map((f) =>
                                          f.internalName === fname ? { ...f, chromePositionMode: k } : f
                                        )
                                      );
                                    }}
                                  />
                                </Stack>
                              )}
                            </>
                          ) : null}
                        </Stack>
                      );
                    }
                    if (isBanner && fcRow) {
                      const bannerOpen = structureFieldOpen[fname] === true;
                      return (
                        <Stack
                          key={fname}
                          tokens={{ childrenGap: 8 }}
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
                          <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
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
                            <Checkbox
                              checked={structurePoolSelected.indexOf(fname) !== -1}
                              onChange={(_, c) => toggleStructurePoolSelect(fname, !!c)}
                              styles={{ text: { display: 'none' } }}
                              title="Selecionar para mover em conjunto"
                            />
                            <IconButton
                              iconProps={{ iconName: bannerOpen ? 'ChevronDown' : 'ChevronRight' }}
                              title={bannerOpen ? 'Fechar configurações' : 'Abrir configurações'}
                              ariaLabel={bannerOpen ? 'Fechar configurações' : 'Abrir configurações'}
                              onClick={() => toggleStructureField(fname)}
                              styles={{ root: { width: 30, height: 30 } }}
                            />
                            <Text styles={{ root: { fontWeight: 600, minWidth: 80 } }}>Banner</Text>
                            <Text variant="small" styles={{ root: { color: '#605e5c', flex: '1 1 200px' } }}>
                              {fname} · {bannerOpen ? 'configurações visíveis' : 'clique para configurar'}
                            </Text>
                            <DefaultButton text="Remover" onClick={() => removeField(fname)} />
                          </Stack>
                          {bannerOpen ? (
                            <>
                              <TextField
                                label="URL da imagem"
                                value={fcRow.bannerImageUrl ?? ''}
                                onChange={(_, v) => {
                                  const t = v ?? '';
                                  setFields((prev) =>
                                    prev.map((f) =>
                                      f.internalName === fname ? { ...f, bannerImageUrl: t.trim() || undefined } : f
                                    )
                                  );
                                }}
                              />
                              <Stack horizontal tokens={{ childrenGap: 12 }} wrap styles={{ root: { width: '100%' } }}>
                                <TextField
                                  label="Largura (%)"
                                  description="Largura da imagem em % da área do formulário (1–100)."
                                  styles={{ root: { minWidth: 140, maxWidth: 200 } }}
                                  value={String(resolveBannerWidthPercent(fcRow))}
                                  onChange={(_, v) => {
                                    const t = (v ?? '').trim().replace(',', '.');
                                    if (t === '') {
                                      setFields((prev) =>
                                        prev.map((f) =>
                                          f.internalName === fname ? { ...f, bannerWidthPercent: undefined } : f
                                        )
                                      );
                                      return;
                                    }
                                    const n = Number(t);
                                    if (!isFinite(n)) return;
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname
                                          ? { ...f, bannerWidthPercent: Math.min(100, Math.max(1, n)) }
                                          : f
                                      )
                                    );
                                  }}
                                />
                                <TextField
                                  label="Altura (px)"
                                  description="Opcional. Altura do banner em pixels."
                                  styles={{ root: { minWidth: 140, maxWidth: 200 } }}
                                  value={
                                    fcRow.bannerHeightPx != null && isFinite(fcRow.bannerHeightPx)
                                      ? String(fcRow.bannerHeightPx)
                                      : ''
                                  }
                                  onChange={(_, v) => {
                                    const t = (v ?? '').trim().replace(',', '.');
                                    if (t === '') {
                                      setFields((prev) =>
                                        prev.map((f) =>
                                          f.internalName === fname ? { ...f, bannerHeightPx: undefined } : f
                                        )
                                      );
                                      return;
                                    }
                                    const n = Number(t);
                                    if (!isFinite(n)) return;
                                    setFields((prev) =>
                                      prev.map((f) =>
                                        f.internalName === fname
                                          ? { ...f, bannerHeightPx: Math.min(2000, Math.max(40, Math.floor(n))) }
                                          : f
                                      )
                                    );
                                  }}
                                />
                              </Stack>
                              <Stack tokens={{ childrenGap: 8 }}>
                                {st.id === FORM_FIXOS_STEP_ID ? (
                                  <>
                                    <Dropdown
                                      label="Zona fixa"
                                      options={FIXED_CHROME_PLACEMENT_OPTIONS}
                                      selectedKey={resolveFixedPlacement(fcRow)}
                                      onChange={(_, o) => {
                                        if (!o) return;
                                        const k = String(o.key) as TFixedChromePlacement;
                                        setFields((prev) =>
                                          prev.map((f) =>
                                            f.internalName === fname ? { ...f, fixedPlacement: k } : f
                                          )
                                        );
                                      }}
                                    />
                                    <Dropdown
                                      label="Posicionamento"
                                      options={CHROME_POSITION_MODE_OPTIONS}
                                      selectedKey={resolveChromePositionMode(fcRow)}
                                      onChange={(_, o) => {
                                        if (!o) return;
                                        const k = String(o.key) as TChromePositionMode;
                                        setFields((prev) =>
                                          prev.map((f) =>
                                            f.internalName === fname ? { ...f, chromePositionMode: k } : f
                                          )
                                        );
                                      }}
                                    />
                                  </>
                                ) : (
                                  <>
                                    <Dropdown
                                      label="Posição no formulário"
                                      options={BANNER_PLACEMENT_DROPDOWN_OPTIONS}
                                      selectedKey={resolveBannerPlacement(fcRow)}
                                      onChange={(_, o) => {
                                        if (!o) return;
                                        const k = String(o.key) as TFormBannerPlacement;
                                        setFields((prev) =>
                                          prev.map((f) =>
                                            f.internalName === fname ? { ...f, bannerPlacement: k } : f
                                          )
                                        );
                                      }}
                                    />
                                    {resolveBannerPlacement(fcRow) !== 'inStep' && (
                                      <Dropdown
                                        label="Posicionamento"
                                        options={CHROME_POSITION_MODE_OPTIONS}
                                        selectedKey={resolveChromePositionMode(fcRow)}
                                        onChange={(_, o) => {
                                          if (!o) return;
                                          const k = String(o.key) as TChromePositionMode;
                                          setFields((prev) =>
                                            prev.map((f) =>
                                              f.internalName === fname ? { ...f, chromePositionMode: k } : f
                                            )
                                          );
                                        }}
                                      />
                                    )}
                                  </>
                                )}
                              </Stack>
                            </>
                          ) : null}
                        </Stack>
                      );
                    }
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
                        <Checkbox
                          checked={structurePoolSelected.indexOf(fname) !== -1}
                          onChange={(_, c) => toggleStructurePoolSelect(fname, !!c)}
                          styles={{ text: { display: 'none' } }}
                          title="Selecionar para mover em conjunto"
                        />
                        <Text styles={{ root: { fontWeight: 600, minWidth: 120 } }}>
                          {mm ? mm.Title : fname === FORM_ATTACHMENTS_FIELD_INTERNAL ? 'Anexos ao item' : fname}
                        </Text>
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                          {fname} ·{' '}
                          {fname === FORM_ATTACHMENTS_FIELD_INTERNAL
                            ? 'campo virtual · etapa definida aqui; destino da gravação na aba Anexos'
                            : mm
                              ? mm.MappedType
                              : '—'}
                          {mm?.Required ? ' · obrigatório na lista' : ''}
                          {FORM_SYSTEM_LIST_METADATA_INTERNAL_NAMES.has(fname)
                            ? ' · sistema: só leitura no formulário (aba Regras não aplica)'
                            : ''}
                        </Text>
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
                          <DefaultButton text="Colunas" onClick={() => setColumnSpanModalField(fname)} />
                          <Text variant="small" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
                            {formatFieldColumnSpanConfigSummary(fcRow, fname)}
                          </Text>
                        </Stack>
                        {st.id === FORM_FIXOS_STEP_ID && fcRow && (
                          <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 8 }}>
                            <Dropdown
                              label="Zona fixa"
                              options={FIXED_CHROME_PLACEMENT_OPTIONS}
                              selectedKey={resolveFixedPlacement(fcRow)}
                              onChange={(_, o) => {
                                if (!o) return;
                                const k = String(o.key) as TFixedChromePlacement;
                                setFields((prev) =>
                                  prev.map((f) =>
                                    f.internalName === fname ? { ...f, fixedPlacement: k } : f
                                  )
                                );
                              }}
                            />
                            <Dropdown
                              label="Posicionamento"
                              options={CHROME_POSITION_MODE_OPTIONS}
                              selectedKey={resolveChromePositionMode(fcRow)}
                              onChange={(_, o) => {
                                if (!o) return;
                                const k = String(o.key) as TChromePositionMode;
                                setFields((prev) =>
                                  prev.map((f) =>
                                    f.internalName === fname ? { ...f, chromePositionMode: k } : f
                                  )
                                );
                              }}
                            />
                          </Stack>
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
                  {st.id === FORM_OCULTOS_STEP_ID && (
                    <>
                      <Text variant="medium" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>
                        Campos fora do formulário
                      </Text>
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        Campos ainda fora do formulário: marque os que quiser e use a barra amarela no topo desta aba
                        (ícone a pulsar), ou arraste para a etapa desejada.
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
                                e.dataTransfer.setData(
                                  'text/plain',
                                  dragPayloadPool(FORM_ATTACHMENTS_FIELD_INTERNAL)
                                );
                                e.dataTransfer.effectAllowed = 'move';
                              }}
                              style={{
                                cursor: 'grab',
                                display: 'flex',
                                alignItems: 'center',
                                color: '#605e5c',
                              }}
                            >
                              <Icon iconName="GripperBarVertical" />
                            </span>
                            <Checkbox
                              label="Anexos ao item (controlo de ficheiros)"
                              checked={structurePoolSelected.indexOf(FORM_ATTACHMENTS_FIELD_INTERNAL) !== -1}
                              onChange={(_, c) => toggleStructurePoolSelect(FORM_ATTACHMENTS_FIELD_INTERNAL, !!c)}
                            />
                            <Text variant="small" styles={{ root: { minWidth: 80 } }}>
                              anexos
                            </Text>
                            <Text variant="small" styles={{ root: { color: '#a19f9d', flex: '1 1 240px' } }}>
                              Arraste para a etapa desejada. Onde gravar (lista ou biblioteca): aba Anexos.
                            </Text>
                          </Stack>
                        );
                      })()}
                      <Stack
                        horizontal
                        verticalAlign="center"
                        tokens={{ childrenGap: 8 }}
                        wrap
                        styles={{ root: { marginTop: 4 } }}
                      >
                        <DefaultButton text="Adicionar alerta" onClick={addAlertField} />
                        <DefaultButton text="Adicionar banner" onClick={addBannerField} />
                        <Text variant="small" styles={{ root: { color: '#a19f9d', flex: '1 1 240px' } }}>
                          Alerta condicional ou banner por URL; arraste para a etapa. Não grava na lista.
                        </Text>
                      </Stack>
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
                              style={{
                                cursor: 'grab',
                                display: 'flex',
                                alignItems: 'center',
                                color: '#605e5c',
                              }}
                            >
                              <Icon iconName="GripperBarVertical" />
                            </span>
                            <Checkbox
                              label={`${m.Title} (${m.InternalName})${m.Required ? ' *' : ''}`}
                              checked={structurePoolSelected.indexOf(m.InternalName) !== -1}
                              onChange={(_, c) => toggleStructurePoolSelect(m.InternalName, !!c)}
                            />
                            <Text variant="small" styles={{ root: { minWidth: 80 } }}>
                              {m.MappedType}
                              {m.Required ? ' · obrig. lista' : ''}
                            </Text>
                          </Stack>
                        );
                      })}
                    </>
                  )}
                  {st.id === FORM_FIXOS_STEP_ID && (
                    <>
                      <Text variant="medium" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>
                        Incluir em Fixos
                      </Text>
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        Marque os campos e use a barra amarela no topo da aba Estrutura para colocar em Fixos ou outra
                        etapa; em Fixos defina topo ou rodapé na linha do item.
                      </Text>
                      <Stack
                        horizontal
                        verticalAlign="center"
                        tokens={{ childrenGap: 8 }}
                        wrap
                        styles={{ root: { marginTop: 4 } }}
                      >
                        <DefaultButton text="Adicionar alerta" onClick={addAlertFieldToFixos} />
                        <DefaultButton text="Adicionar banner" onClick={addBannerFieldToFixos} />
                        <Text variant="small" styles={{ root: { color: '#a19f9d', flex: '1 1 240px' } }}>
                          Alerta condicional ou banner por URL; depois escolha topo ou rodapé na linha do item.
                        </Text>
                      </Stack>
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
                            key={`fixos-${m.InternalName}`}
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
                              style={{
                                cursor: 'grab',
                                display: 'flex',
                                alignItems: 'center',
                                color: '#605e5c',
                              }}
                            >
                              <Icon iconName="GripperBarVertical" />
                            </span>
                            <Checkbox
                              label={`${m.Title} (${m.InternalName})${m.Required ? ' *' : ''}`}
                              checked={structurePoolSelected.indexOf(m.InternalName) !== -1}
                              onChange={(_, c) => toggleStructurePoolSelect(m.InternalName, !!c)}
                            />
                            <Text variant="small" styles={{ root: { minWidth: 80 } }}>
                              {m.MappedType}
                              {m.Required ? ' · obrig. lista' : ''}
                            </Text>
                          </Stack>
                        );
                      })}
                    </>
                  )}
                </Stack>
                )}
              </Stack>
              );
            })}
            {linkedChildFormsSortedForStructure.length > 0 && (
              <>
                <Text variant="medium" styles={{ root: { fontWeight: 600, marginTop: 16 } }}>
                  Listas vinculadas (etapa no passador)
                </Text>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Cada bloco aparece só na etapa escolhida do formulário principal.
                </Text>
                {linkedMainStepPlacementOptions.length === 0 ? (
                  <MessageBar messageBarType={MessageBarType.info}>
                    Adicione pelo menos uma etapa (além de Ocultos/Fixos) para posicionar os blocos.
                  </MessageBar>
                ) : (
                  <Stack tokens={{ childrenGap: 10 }}>
                    {linkedChildFormsSortedForStructure.map((cfg) => {
                      const validSel =
                        typeof cfg.mainFormStepId === 'string' &&
                        linkedMainStepPlacementOptions.some((o) => o.key === cfg.mainFormStepId);
                      const selectedKey = validSel ? cfg.mainFormStepId! : linkedMainStepDefaultKey;
                      const label = (cfg.title?.trim() || cfg.listTitle.trim() || cfg.id).slice(0, 120);
                      return (
                        <Stack
                          key={`lcf-step-${cfg.id}`}
                          horizontal
                          verticalAlign="end"
                          tokens={{ childrenGap: 12 }}
                          wrap
                        >
                          <Text styles={{ root: { minWidth: 160, maxWidth: 280 } }} variant="small">
                            {label}
                          </Text>
                          <div style={{ minWidth: 260, flex: '1 1 200px' }}>
                            <Dropdown
                              label="Etapa"
                              options={linkedMainStepPlacementOptions}
                              selectedKey={selectedKey}
                              onChange={(_, o) => {
                                if (!o) return;
                                setLinkedChildForms((prev) =>
                                  patchLinkedChildFormById(prev, cfg.id, {
                                    mainFormStepId: String(o.key),
                                  })
                                );
                              }}
                            />
                          </div>
                        </Stack>
                      );
                    })}
                  </Stack>
                )}
              </>
            )}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Regras dos campos">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            {!fieldsListedForRulesTab.length ? (
              <Text>Adicione campos ao formulário na aba Estrutura.</Text>
            ) : (
              <Stack tokens={{ childrenGap: 8 }}>
                <Stack horizontal wrap verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  {fieldRulesTabSort === 'asc' ? (
                    <PrimaryButton text="Crescente (A–Z)" onClick={() => setFieldRulesTabSort('asc')} />
                  ) : (
                    <DefaultButton text="Crescente (A–Z)" onClick={() => setFieldRulesTabSort('asc')} />
                  )}
                  {fieldRulesTabSort === 'desc' ? (
                    <PrimaryButton text="Decrescente (Z–A)" onClick={() => setFieldRulesTabSort('desc')} />
                  ) : (
                    <DefaultButton text="Decrescente (Z–A)" onClick={() => setFieldRulesTabSort('desc')} />
                  )}
                  {fieldRulesTabSort === 'type' ? (
                    <PrimaryButton
                      text="Tipo de dado (agrupa por tipo)"
                      onClick={() => setFieldRulesTabSort('type')}
                    />
                  ) : (
                    <DefaultButton
                      text="Tipo de dado (agrupa por tipo)"
                      onClick={() => setFieldRulesTabSort('type')}
                    />
                  )}
                </Stack>
                {fieldsForRulesTabDisplay.map((fc) => {
                  const mm = meta.find((m) => m.InternalName === fc.internalName);
                  const title = mm?.Title ?? fc.internalName;
                  const typeAs = (mm?.TypeAsString ?? '').trim() || '—';
                  return (
                    <Stack
                      key={fc.internalName}
                      horizontal
                      verticalAlign="center"
                      tokens={{ childrenGap: 16 }}
                      wrap
                      styles={{
                        root: {
                          padding: '8px 10px',
                          borderRadius: 4,
                          border: '1px solid #edebe9',
                          background: '#faf9f8',
                        },
                      }}
                    >
                      <Text styles={{ root: { fontWeight: 600, minWidth: 140, maxWidth: 220 } }}>{title}</Text>
                      <Text
                        variant="small"
                        styles={{ root: { fontFamily: 'monospace', minWidth: 120, color: '#323130' } }}
                      >
                        {fc.internalName}
                      </Text>
                      <Text variant="small" styles={{ root: { color: '#605e5c', minWidth: 96 } }}>
                        {typeAs}
                      </Text>
                      <DefaultButton text="Regras" onClick={() => setFieldPanelName(fc.internalName)} />
                      <DefaultButton
                        text="Clonar regras"
                        onClick={() => {
                          setCloneRulesModalTarget(fc.internalName);
                          setCloneRulesSourceKey(undefined);
                        }}
                      />
                    </Stack>
                  );
                })}
              </Stack>
            )}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Componentes">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <FormManagerComponentsTabContent
              loading={loading}
              stepLayout={stepLayout}
              onStepLayoutChange={setStepLayout}
              stepAccentPaletteSlot={stepAccentPaletteSlot}
              onStepAccentPaletteSlotChange={(slot) =>
                setStepAccentPaletteSlot(slot === 'themePrimary' ? undefined : slot)
              }
              stepNavButtons={stepNavButtons}
              onStepNavButtonsChange={setStepNavButtons}
              formDataLoadingKind={formDataLoadingKind}
              onFormDataLoadingKindChange={setFormDataLoadingKind}
              defaultSubmitLoadingKind={defaultSubmitLoadingKind}
              onDefaultSubmitLoadingKindChange={setDefaultSubmitLoadingKind}
              historyEnabled={historyEnabled}
              onHistoryEnabledChange={setHistoryEnabled}
              historyPresentationKind={historyPresentationKind}
              onHistoryPresentationKindChange={setHistoryPresentationKind}
              historyButtonKind={historyButtonKind}
              onHistoryButtonKindChange={setHistoryButtonKind}
              historyButtonLabel={historyButtonLabel}
              onHistoryButtonLabelChange={setHistoryButtonLabel}
              historyButtonIcon={historyButtonIcon}
              onHistoryButtonIconChange={setHistoryButtonIcon}
              historyPanelSubtitle={historyPanelSubtitle}
              onHistoryPanelSubtitleChange={setHistoryPanelSubtitle}
              historyGroupTitles={historyGroupTitles}
              onHistoryGroupTitlesChange={setHistoryGroupTitles}
              siteGroups={siteGroups}
              siteGroupsSorted={siteGroupsSorted}
              siteGroupsLoading={siteGroupsLoading}
              siteGroupsErr={siteGroupsErr}
              onRetryLoadSiteGroups={loadSiteGroups}
              historyLayoutKind={historyLayoutKind}
              onHistoryLayoutKindChange={setHistoryLayoutKind}
            />
          </Stack>
        </PivotItem>
        <PivotItem headerText="Anexos">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12, width: '100%', maxWidth: '100%' } }}>
            <FormManagerAttachmentsTabContent
              loading={loading}
              primaryListTitle={listTitle}
              attachmentStorageKind={attachmentStorageKind}
              onAttachmentStorageKindChange={setAttachmentStorageKind}
              attachmentLibraryTitle={attachmentLibLibraryTitle}
              onAttachmentLibraryTitleChange={setAttachmentLibLibraryTitle}
              attachmentLibraryLookupField={attachmentLibLookupField}
              onAttachmentLibraryLookupFieldChange={setAttachmentLibLookupField}
              attachmentLibFolderTree={attachmentLibFolderTree}
              onAttachmentLibFolderTreeChange={setAttachmentLibFolderTree}
              attachmentUploadLayout={attachmentUploadLayout}
              onAttachmentUploadLayoutChange={setAttachmentUploadLayout}
              attachmentFilePreview={attachmentFilePreview}
              onAttachmentFilePreviewChange={setAttachmentFilePreview}
              attachmentAllowedExtensions={attachAllowedExt}
              onAttachmentExtensionToggle={(ext, selected) => {
                const e = ext.trim().replace(/^\./, '').toLowerCase();
                if (!e) return;
                setAttachAllowedExt((prev) => {
                  if (selected) return prev.indexOf(e) === -1 ? prev.concat([e]) : prev;
                  return prev.filter((x) => x !== e);
                });
              }}
              attachmentFolderStepOptions={attachmentFolderStepOptions}
              attachmentFolderVisibilityEditor={attachmentFolderVisibilityEditor}
            />
          </Stack>
        </PivotItem>
        <PivotItem headerText="Botões">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
              Onde mostrar a barra de botões
            </Text>
            <Stack horizontal wrap verticalAlign="end" tokens={{ childrenGap: 16 }}>
              <Dropdown
                label="Vertical"
                options={[
                  { key: 'bottom', text: 'Inferior' },
                  { key: 'top', text: 'Superior' },
                ]}
                selectedKey={customButtonsBarVertical}
                onChange={(_, o) => {
                  if (!o) return;
                  setCustomButtonsBarVertical(String(o.key) as TFormCustomButtonsBarVertical);
                }}
              />
              <Dropdown
                label="Horizontal"
                options={[
                  { key: 'left', text: 'Esquerda' },
                  { key: 'right', text: 'Direita' },
                ]}
                selectedKey={customButtonsBarHorizontal}
                onChange={(_, o) => {
                  if (!o) return;
                  setCustomButtonsBarHorizontal(String(o.key) as TFormCustomButtonsBarHorizontal);
                }}
              />
            </Stack>
            <PrimaryButton text="Adicionar botão" onClick={addCustomButton} />
            {customButtons.map((btn, bi) => {
              const chk = checkboxesFromModes(btn.modes);
              const whenRowsState = parseButtonWhenToRows(btn.when, meta);
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
                    <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
                      <DefaultButton text="Clonar" onClick={() => cloneCustomButton(bi)} />
                      <DefaultButton text="Remover botão" onClick={() => removeCustomButton(bi)} />
                    </Stack>
                  </Stack>
                  {panelOpen && (
                  <>
                  <TextField
                    label="Texto do botão"
                    value={btn.label}
                    onChange={(_, v) => patchCustomButton(bi, { label: v ?? '' })}
                  />
                  {(btn.operation ?? 'legacy') === 'history' && (
                    <TextField
                      label="Descrição curta"
                      description="Texto de ajuda (tooltip no botão no formulário)."
                      multiline
                      rows={2}
                      value={btn.shortDescription ?? ''}
                      onChange={(_, v) => {
                        const t = (v ?? '').trim();
                        patchCustomButton(bi, { shortDescription: t ? t : undefined });
                      }}
                    />
                  )}
                  <Dropdown
                    label="Tipo de operação"
                    options={buttonOperationDropdownOptions}
                    selectedKey={(btn.operation ?? 'legacy') as string}
                    onChange={(_, o) => {
                      if (!o) return;
                      const k = String(o.key) as TFormCustomButtonOperation;
                      patchCustomButton(bi, {
                        operation: k,
                        ...(k === 'redirect'
                          ? { redirectUrlTemplate: btn.redirectUrlTemplate ?? '', actions: [] }
                          : {}),
                        ...(k === 'history'
                          ? {
                              actions: [],
                              behavior: 'actionsOnly',
                              submitLoadingKind: undefined,
                            }
                          : {}),
                        ...(k !== 'history' ? { shortDescription: undefined } : {}),
                      });
                    }}
                  />
                  {(btn.operation ?? 'legacy') === 'history' && (
                    <MessageBar messageBarType={MessageBarType.info}>
                      Preferível o botão integrado: ative-o na aba «Componentes» (secção Histórico de auditoria) e
                      configure a lista de log na aba «Lista de logs». Pode remover este botão legado.
                    </MessageBar>
                  )}
                  {(btn.operation ?? 'legacy') !== 'history' && (
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
                  )}
                  {(btn.operation ?? 'legacy') === 'redirect' && (
                    <Stack tokens={{ childrenGap: 10 }}>
                      <TextField
                        label="URL de destino"
                        description="Escreva o endereço. Use {{}} vazio para escolher um campo na lista abaixo."
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
                    <ThemePaletteSlotDropdown
                      label="Cor do botão (tema do site)"
                      selectedKey={resolveFormCustomButtonPaletteSlot(btn)}
                      onChange={(slot) =>
                        patchCustomButton(bi, {
                          themePaletteSlot: slot,
                          appearance: slot === 'outline' ? 'default' : 'primary',
                        })
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
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Modal de confirmação
                  </Text>
                  <Toggle
                    label="Pedir confirmação antes de executar (primeiro passo; cancelar não executa ações nem o resto do botão)"
                    checked={btn.confirmBeforeRun?.enabled === true}
                    onChange={(_, c) => {
                      if (c) {
                        patchCustomButton(bi, {
                          confirmBeforeRun: {
                            enabled: true,
                            kind: btn.confirmBeforeRun?.kind ?? 'info',
                            message: btn.confirmBeforeRun?.message ?? '',
                            ...(btn.confirmBeforeRun?.promptFieldInternalName
                              ? {
                                  promptFieldInternalName: btn.confirmBeforeRun.promptFieldInternalName,
                                }
                              : {}),
                          },
                        });
                      } else {
                        patchCustomButton(bi, { confirmBeforeRun: undefined });
                      }
                    }}
                  />
                  {btn.confirmBeforeRun?.enabled === true && (
                    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { paddingLeft: 4 } }}>
                      <Dropdown
                        label="Ícone / tipo"
                        options={[
                          { key: 'info', text: 'Informação' },
                          { key: 'success', text: 'Sucesso' },
                          { key: 'warning', text: 'Aviso' },
                          { key: 'error', text: 'Erro / crítico' },
                          { key: 'blocked', text: 'Bloqueado' },
                        ]}
                        selectedKey={btn.confirmBeforeRun?.kind ?? 'info'}
                        onChange={(_, o) => {
                          if (!o) return;
                          patchCustomButton(bi, {
                            confirmBeforeRun: {
                              enabled: true,
                              kind: String(o.key) as TFormCustomButtonConfirmKind,
                              message: btn.confirmBeforeRun?.message ?? '',
                              ...(btn.confirmBeforeRun?.promptFieldInternalName
                                ? {
                                    promptFieldInternalName: btn.confirmBeforeRun.promptFieldInternalName,
                                  }
                                : {}),
                            },
                          });
                        }}
                      />
                      <TextField
                        label="Mensagem"
                        multiline
                        rows={4}
                        value={btn.confirmBeforeRun?.message ?? ''}
                        onChange={(_, v) =>
                          patchCustomButton(bi, {
                            confirmBeforeRun: {
                              enabled: true,
                              kind: btn.confirmBeforeRun?.kind ?? 'info',
                              message: v ?? '',
                              ...(btn.confirmBeforeRun?.promptFieldInternalName
                                ? {
                                    promptFieldInternalName: btn.confirmBeforeRun.promptFieldInternalName,
                                  }
                                : {}),
                            },
                          })
                        }
                        description="Obrigatório salvo se escolher abaixo um campo a preencher (nesse caso a mensagem pode ficar vazia). Com ambos vazios, a confirmação não é gravada."
                      />
                      <Dropdown
                        label="Campo da lista principal a preencher no modal"
                        options={confirmModalPromptFieldOptions}
                        selectedKey={btn.confirmBeforeRun?.promptFieldInternalName ?? ''}
                        onChange={(_, o) => {
                          if (!o) return;
                          const key = String(o.key);
                          patchCustomButton(bi, {
                            confirmBeforeRun: {
                              enabled: true,
                              kind: btn.confirmBeforeRun?.kind ?? 'info',
                              message: btn.confirmBeforeRun?.message ?? '',
                              ...(key ? { promptFieldInternalName: key } : {}),
                            },
                          });
                        }}
                      />
                    </Stack>
                  )}
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Último passo (após tudo concluir com sucesso)
                  </Text>
                  
                  <Dropdown
                    label="Quando o fluxo do botão terminar sem erro"
                    options={BUTTON_FINISH_AFTER_OPTIONS}
                    selectedKey={
                      btn.finishAfterRun?.kind === 'redirect'
                        ? 'redirect'
                        : btn.finishAfterRun?.kind === 'clearForm'
                          ? 'clearForm'
                          : 'none'
                    }
                    onChange={(_, o) => {
                      if (!o) return;
                      const k = String(o.key);
                      if (k === 'none') {
                        patchCustomButton(bi, { finishAfterRun: undefined });
                        return;
                      }
                      if (k === 'clearForm') {
                        patchCustomButton(bi, { finishAfterRun: { kind: 'clearForm' } });
                        return;
                      }
                      const prevTpl =
                        btn.finishAfterRun?.kind === 'redirect'
                          ? btn.finishAfterRun.redirectUrlTemplate
                          : '';
                      patchCustomButton(bi, {
                        finishAfterRun: { kind: 'redirect', redirectUrlTemplate: prevTpl },
                      });
                    }}
                  />
                  {btn.finishAfterRun?.kind === 'redirect' && (
                    <TextField
                      label="URL de redirecionamento"
                      value={btn.finishAfterRun.redirectUrlTemplate}
                      onChange={(_, v) =>
                        patchCustomButton(bi, {
                          finishAfterRun: {
                            kind: 'redirect',
                            redirectUrlTemplate: v ?? '',
                          },
                        })
                      }
                      multiline
                      rows={2}
                    />
                  )}
                  <Checkbox
                    label="Só mostrar se todos os campos obrigatórios estiverem preenchidos"
                    checked={btn.showOnlyWhenAllRequiredFilled === true}
                    onChange={(_, c) =>
                      patchCustomButton(bi, {
                        showOnlyWhenAllRequiredFilled: c ? true : undefined,
                      })
                    }
                  />
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Grupos do SharePoint
                  </Text>
         
                  <TextField
                    placeholder="Filtrar grupos por nome"
                    value={customButtonGroupNameFilter}
                    onChange={(_: unknown, v?: string) => setCustomButtonGroupNameFilter(v ?? '')}
                    styles={{ root: { maxWidth: 420 } }}
                  />
                  {siteGroupsLoading && <Spinner label="A carregar grupos do site…" />}
                  {siteGroupsErr && (
                    <>
                      <MessageBar messageBarType={MessageBarType.warning}>{siteGroupsErr}</MessageBar>
                      <DefaultButton text="Tentar carregar grupos novamente" onClick={() => loadSiteGroups()} />
                    </>
                  )}
                  {!siteGroupsLoading ? (
                    <Stack
                      tokens={{ childrenGap: 6 }}
                      styles={{
                        root: {
                          maxHeight: 240,
                          overflowY: 'auto',
                          border: '1px solid #edebe9',
                          borderRadius: 4,
                          padding: 8,
                        },
                      }}
                    >
                      {(btn.groupTitles ?? [])
                        .filter(
                          (t) =>
                            !siteGroups.some(
                              (g) => normSpGroupTitle(g.Title) === normSpGroupTitle(t)
                            )
                        )
                        .filter((t) => {
                          const q = customButtonGroupNameFilter.trim().toLowerCase();
                          return !q || t.toLowerCase().includes(q);
                        })
                        .map((t, oi) => (
                          <Checkbox
                            key={`orphan-grp-${bi}-${oi}-${t}`}
                            label={`${t} (guardado; não na lista do site)`}
                            checked
                            onChange={(_, c) => {
                              if (c) return;
                              const cur = btn.groupTitles ?? [];
                              const n = normSpGroupTitle(t);
                              const next = cur.filter((x) => normSpGroupTitle(x) !== n);
                              patchCustomButton(bi, { groupTitles: next.length ? next : undefined });
                            }}
                          />
                        ))}
                      {siteGroupsSortedForCustomButtons.map((g) => {
                        const cur = btn.groupTitles ?? [];
                        const n = normSpGroupTitle(g.Title);
                        const checked = cur.some((x) => normSpGroupTitle(x) === n);
                        return (
                          <Checkbox
                            key={g.Id}
                            label={g.Title}
                            title={g.Description || undefined}
                            checked={checked}
                            onChange={(_, c) => {
                              let next: string[];
                              if (c) {
                                next = checked ? cur : cur.concat([g.Title]);
                              } else {
                                next = cur.filter((x) => normSpGroupTitle(x) !== n);
                              }
                              patchCustomButton(bi, { groupTitles: next.length ? next : undefined });
                            }}
                          />
                        );
                      })}
                      {siteGroupsSorted.length > 0 &&
                        !siteGroupsSortedForCustomButtons.length &&
                        customButtonGroupNameFilter.trim() && (
                          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                            Nenhum grupo corresponde ao filtro.
                          </Text>
                        )}
                      {!siteGroupsSorted.length && !(btn.groupTitles ?? []).length && (
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                          Nenhum grupo no site.
                        </Text>
                      )}
                    </Stack>
                  ) : null}
                  <Checkbox
                    label="Mostrar só quando as condições abaixo forem verdadeiras"
                    checked={!!btn.when}
                    onChange={(_, c) => {
                      if (c) patchCustomButton(bi, { when: whenUiToNode(defaultWhenUi(meta)) });
                      else patchCustomButton(bi, { when: undefined });
                    }}
                  />
                  {btn.when && (
                    <Stack tokens={{ childrenGap: 10 }}>
                      <Dropdown
                        label="Lógica entre condições"
                        options={[
                          { key: 'all', text: 'Todas (E)' },
                          { key: 'any', text: 'Pelo menos uma (OU)' },
                        ]}
                        selectedKey={whenRowsState.rows.length <= 1 ? 'all' : whenRowsState.combiner}
                        disabled={whenRowsState.rows.length <= 1}
                        onChange={(_, o) => o && setButtonWhenCombiner(bi, String(o.key) as 'all' | 'any')}
                      />
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        Condições nos dados do formulário
                      </Text>
                      {whenRowsState.rows.map((row, ri) => (
                        <Stack
                          key={`btn-${btn.id}-when-${ri}`}
                          tokens={{ childrenGap: 8 }}
                          styles={{
                            root: {
                              border: '1px solid #edebe9',
                              borderRadius: 4,
                              padding: 10,
                              background: '#faf9f8',
                            },
                          }}
                        >
                          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                              Condição {ri + 1}
                            </Text>
                            <IconButton
                              iconProps={{ iconName: 'Delete' }}
                              title="Remover condição"
                              disabled={whenRowsState.rows.length <= 1}
                              onClick={() => removeButtonWhenRow(bi, ri)}
                            />
                          </Stack>
                          <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                            <Dropdown
                              label="Campo"
                              options={fieldOptions}
                              selectedKey={row.field || undefined}
                              onChange={(_, o) => o && patchButtonWhenRow(bi, ri, { field: String(o.key) })}
                            />
                            <Dropdown
                              label="Operador"
                              options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                              selectedKey={row.op}
                              onChange={(_, o) => o && patchButtonWhenRow(bi, ri, { op: o.key as TFormConditionOp })}
                            />
                            <Dropdown
                              label="Comparar com"
                              options={[
                                { key: 'literal', text: 'Texto fixo' },
                                { key: 'field', text: 'Outro campo' },
                                { key: 'token', text: 'Token' },
                              ]}
                              selectedKey={row.compareKind}
                              onChange={(_, o) =>
                                o && patchButtonWhenRow(bi, ri, { compareKind: o.key as IWhenUi['compareKind'] })
                              }
                            />
                            <TextField
                              label="Valor"
                              value={row.compareValue}
                              onChange={(_, v) => patchButtonWhenRow(bi, ri, { compareValue: v ?? '' })}
                              disabled={
                                row.op === 'isEmpty' ||
                                row.op === 'isFilled' ||
                                row.op === 'isTrue' ||
                                row.op === 'isFalse'
                              }
                            />
                          </Stack>
                        </Stack>
                      ))}
                      <DefaultButton text="Adicionar condição" onClick={() => addButtonWhenRow(bi)} />
                    </Stack>
                  )}
                  {(btn.operation ?? 'legacy') !== 'redirect' && (btn.operation ?? 'legacy') !== 'history' && (
                    <FormManagerChainedActionsBlock
                      actions={btn.actions}
                      patchAction={(ai, next) => patchButtonAction(bi, ai, next)}
                      removeAction={(ai) => removeButtonAction(bi, ai)}
                      addAction={() => addButtonAction(bi)}
                      patchActionCondition={(ai, when) => patchButtonActionCondition(bi, ai, when)}
                      reactKeysPrefix={`btn-${btn.id}`}
                      meta={meta}
                      metaSortedForPool={metaSortedForPool}
                      steps={steps}
                      fieldOptions={fieldOptions}
                      loading={loading}
                      getDefaultWhenUi={() => defaultWhenUi(meta)}
                    />
                  )}
                  </>
                  )}
                </Stack>
              );
            })}
            {!customButtons.length && <Text>Nenhum botão personalizado.</Text>}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Lista de logs">
          <FormManagerActionLogTabContent
            historyEnabled={historyEnabled}
            captureEnabled={actionLogCaptureEnabled}
            onCaptureEnabledChange={setActionLogCaptureEnabled}
            listTitle={actionLogListTitle}
            onListTitleChange={(t) => {
              setActionLogListTitle(t);
              setActionLogFieldInternalName('');
              setActionLogSourceListLookupFieldInternalName('');
              setActionLogCaptureEnabled(false);
            }}
            actionFieldInternalName={actionLogFieldInternalName}
            onActionFieldInternalNameChange={(name) => {
              setActionLogFieldInternalName(name);
              if (!name.trim() && actionLogCaptureEnabled) setActionLogCaptureEnabled(false);
            }}
            primaryListTitle={listTitle.trim()}
            sourceListLookupFieldInternalName={actionLogSourceListLookupFieldInternalName}
            onSourceListLookupFieldInternalNameChange={(name) => {
              setActionLogSourceListLookupFieldInternalName(name);
              if (!name.trim() && actionLogCaptureEnabled) setActionLogCaptureEnabled(false);
            }}
            automaticChangesOnUpdate={actionLogAutomaticChangesOnUpdate}
            onAutomaticChangesOnUpdateChange={setActionLogAutomaticChangesOnUpdate}
            descriptionsHtmlByButtonId={actionLogDescById}
            onDescriptionChange={(buttonId, html) =>
              setActionLogDescById((prev) => ({ ...prev, [buttonId]: html }))
            }
            descriptionPaletteSlotByButtonId={actionLogPaletteSlotById}
            onDescriptionPaletteSlotChange={(buttonId, slot) =>
              setActionLogPaletteSlotById((prev) => ({ ...prev, [buttonId]: slot }))
            }
            customButtons={customButtons}
          />
        </PivotItem>
        <PivotItem headerText="Listas vinculadas">
          <FormManagerLinkedChildFormsTabContent
            primaryListTitle={listTitle.trim()}
            linkedChildForms={linkedChildForms}
            onLinkedChildFormsChange={setLinkedChildForms}
            mainAttachmentStorageKind={attachmentStorageKind}
            mainAttachmentLibraryFromPanel={attachmentLibraryFromPanelState(
              attachmentLibLibraryTitle,
              attachmentLibLookupField,
              attachmentLibFolderTree
            )}
            listWebServerRelativeUrl={lw}
          />
        </PivotItem>
        <PivotItem headerText="Quebra de permissões">
          <FormManagerPermissionBreakTabContent
            primaryListTitle={listTitle.trim()}
            primaryMeta={meta}
            linkedChildForms={linkedChildForms}
            formManagerForResolve={formManagerForPermissionBreakResolve}
            mainAttachmentLibraryEnabled={mainAttachmentLibraryEnabledTab}
            value={permissionBreak}
            onChange={setPermissionBreak}
            siteGroups={siteGroups}
            siteGroupsLoading={siteGroupsLoading}
            siteGroupsErr={siteGroupsErr}
          />
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
      <Modal isOpen={columnSpanModalField !== null} onDismiss={() => setColumnSpanModalField(null)} isBlocking>
        <Stack
          tokens={{ childrenGap: 16 }}
          styles={{
            root: {
              background: '#ffffff',
              padding: 24,
              maxWidth: 440,
              margin: '48px auto',
              borderRadius: 4,
              boxShadow: '0 6px 24px rgba(0,0,0,0.18)',
            },
          }}
        >
          <Text variant="large">Colunas na grelha</Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Campo:{' '}
            <strong>
              {columnSpanModalField
                ? meta.find((m) => m.InternalName === columnSpanModalField)?.Title ?? columnSpanModalField
                : '—'}
            </strong>{' '}
            {columnSpanModalField ? (
              <span style={{ fontFamily: 'monospace' }}>({columnSpanModalField})</span>
            ) : null}
          </Text>
          <Pivot>
            {COLUMN_SPAN_BY_MODE_TABS.map(({ mode, headerText }) => (
              <PivotItem key={mode} headerText={headerText}>
                <Dropdown
                  label="Colunas ocupadas (de 12)"
                  options={FIELD_COLUMN_SPAN_OPTIONS}
                  selectedKey={String(
                    resolveFieldColumnSpan(
                      fields.find((f) => f.internalName === columnSpanModalField) ?? {
                        internalName: columnSpanModalField ?? '',
                      },
                      mode
                    )
                  )}
                  onChange={(_, o) => {
                    if (!o || !columnSpanModalField) return;
                    const span = Number(o.key);
                    if (span !== 3 && span !== 4 && span !== 6 && span !== 8 && span !== 12) return;
                    applyFieldColumnSpanForMode(columnSpanModalField, mode, span as TFormFieldColumnSpan);
                  }}
                />
              </PivotItem>
            ))}
          </Pivot>
          <DefaultButton text="Fechar" onClick={() => setColumnSpanModalField(null)} />
        </Stack>
      </Modal>
      <Modal isOpen={cloneRulesModalTarget !== null} onDismiss={dismissCloneRulesModal} isBlocking>
        <Stack
          tokens={{ childrenGap: 16 }}
          styles={{
            root: {
              background: '#ffffff',
              padding: 24,
              maxWidth: 440,
              margin: '48px auto',
              borderRadius: 4,
              boxShadow: '0 6px 24px rgba(0,0,0,0.18)',
            },
          }}
        >
          <Text variant="large">Clonar regras</Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Campo de destino:{' '}
            <strong>
              {cloneRulesModalTarget
                ? meta.find((m) => m.InternalName === cloneRulesModalTarget)?.Title ?? cloneRulesModalTarget
                : '—'}
            </strong>{' '}
            <span style={{ fontFamily: 'monospace' }}>({cloneRulesModalTarget})</span>
          </Text>
          <Dropdown
            label="Copiar regras do campo"
            options={cloneRulesSourceOptions}
            selectedKey={cloneRulesSourceKey}
            onChange={(_, o) => setCloneRulesSourceKey(o ? String(o.key) : undefined)}
            disabled={cloneRulesSourceOptions.length === 0}
            placeholder={cloneRulesSourceOptions.length ? 'Selecione o campo' : 'Não há outro campo disponível'}
          />
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Confirmar"
              onClick={applyCloneFieldRules}
              disabled={!cloneRulesSourceKey || cloneRulesSourceOptions.length === 0}
            />
            <DefaultButton text="Cancelar" onClick={dismissCloneRulesModal} />
          </Stack>
        </Stack>
      </Modal>
      <Panel
        isOpen={stepVisibilityPanelStepId !== null}
        type={PanelType.medium}
        headerText="Visibilidade da etapa"
        closeButtonAriaLabel="Fechar"
        onDismiss={() => setStepVisibilityPanelStepId(null)}
        isLightDismiss
        styles={{
          scrollableContent: { maxHeight: '100%' },
        }}
      >
        {(() => {
          const vst = stepVisibilityPanelStepId
            ? steps.find((x) => x.id === stepVisibilityPanelStepId)
            : undefined;
          if (!vst) return null;
          const whenRowsState = parseButtonWhenToRows(vst.showStepWhen, meta);
          return (
            <Stack tokens={{ childrenGap: 14 }}>
              <Text>
                <span style={{ fontWeight: 600 }}>{vst.title}</span>{' '}
                <span style={{ color: '#605e5c', fontFamily: 'monospace', fontSize: 12 }}>{vst.id}</span>
              </Text>
              <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                Modos de formulário
              </Text>
              <Stack horizontal wrap tokens={{ childrenGap: 12 }} verticalAlign="center">
                {ALL_FORM_MANAGER_MODES.map((m) => {
                  const sel = vst.showInFormModes;
                  const checked = !sel?.length || sel.indexOf(m) !== -1;
                  const label = m === 'create' ? 'Criar' : m === 'edit' ? 'Editar' : 'Ver';
                  return (
                    <Checkbox
                      key={m}
                      label={label}
                      checked={checked}
                      onChange={(_, c) =>
                        patchStepById(vst.id, {
                          showInFormModes: toggleStepShowInFormMode(vst.showInFormModes, m, !!c),
                        })
                      }
                    />
                  );
                })}
              </Stack>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Todas marcadas ou nenhuma restrição = Criar, Editar e Ver.
              </Text>
              <Checkbox
                label="Só mostrar esta etapa quando as condições abaixo forem verdadeiras"
                checked={!!vst.showStepWhen}
                onChange={(_, c) =>
                  patchStepById(vst.id, {
                    showStepWhen: c ? whenUiToNode(defaultWhenUi(meta)) : undefined,
                  })
                }
              />
              {vst.showStepWhen && (
                <Stack tokens={{ childrenGap: 10 }}>
                  <Dropdown
                    label="Lógica entre condições"
                    options={[
                      { key: 'all', text: 'Todas (E)' },
                      { key: 'any', text: 'Pelo menos uma (OU)' },
                    ]}
                    selectedKey={whenRowsState.rows.length <= 1 ? 'all' : whenRowsState.combiner}
                    disabled={whenRowsState.rows.length <= 1}
                    onChange={(_, o) =>
                      o && setStepVisibilityWhenCombiner(vst.id, String(o.key) as 'all' | 'any')
                    }
                  />
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Condições sobre os dados do formulário (e grupos de utilizador, se aplicável).
                  </Text>
                  {whenRowsState.rows.map((row, ri) => (
                    <Stack
                      key={`step-${vst.id}-when-${ri}`}
                      tokens={{ childrenGap: 8 }}
                      styles={{
                        root: {
                          border: '1px solid #edebe9',
                          borderRadius: 4,
                          padding: 10,
                          background: '#faf9f8',
                        },
                      }}
                    >
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                          Condição {ri + 1}
                        </Text>
                        <IconButton
                          iconProps={{ iconName: 'Delete' }}
                          title="Remover condição"
                          disabled={whenRowsState.rows.length <= 1}
                          onClick={() => removeStepVisibilityWhenRow(vst.id, ri)}
                        />
                      </Stack>
                      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                        <Dropdown
                          label="Campo"
                          options={fieldOptions}
                          selectedKey={row.field || undefined}
                          disabled={
                            row.compareKind === 'spGroupMember' ||
                            row.compareKind === 'spGroupNotMember'
                          }
                          onChange={(_, o) => o && patchStepVisibilityWhenRow(vst.id, ri, { field: String(o.key) })}
                        />
                        <Dropdown
                          label="Operador"
                          options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                          selectedKey={row.op}
                          disabled={
                            row.compareKind === 'spGroupMember' ||
                            row.compareKind === 'spGroupNotMember'
                          }
                          onChange={(_, o) =>
                            o && patchStepVisibilityWhenRow(vst.id, ri, { op: o.key as TFormConditionOp })
                          }
                        />
                        <Dropdown
                          label="Comparar com"
                          options={[
                            { key: 'literal', text: 'Texto fixo' },
                            { key: 'field', text: 'Outro campo' },
                            { key: 'token', text: 'Token' },
                            { key: 'spGroupMember', text: 'Membro do grupo' },
                            { key: 'spGroupNotMember', text: 'Fora do grupo' },
                          ]}
                          selectedKey={row.compareKind}
                          onChange={(_, o) =>
                            o &&
                            patchStepVisibilityWhenRow(vst.id, ri, {
                              compareKind: o.key as IWhenUi['compareKind'],
                            })
                          }
                        />
                        <TextField
                          label={
                            row.compareKind === 'spGroupMember' ||
                            row.compareKind === 'spGroupNotMember'
                              ? 'Grupo (título)'
                              : 'Valor'
                          }
                          value={row.compareValue}
                          onChange={(_, v) =>
                            patchStepVisibilityWhenRow(vst.id, ri, { compareValue: v ?? '' })
                          }
                          disabled={
                            row.op === 'isEmpty' ||
                            row.op === 'isFilled' ||
                            row.op === 'isTrue' ||
                            row.op === 'isFalse'
                          }
                        />
                      </Stack>
                    </Stack>
                  ))}
                  <DefaultButton text="Adicionar condição" onClick={() => addStepVisibilityWhenRow(vst.id)} />
                </Stack>
              )}
              <DefaultButton text="Fechar" onClick={() => setStepVisibilityPanelStepId(null)} />
            </Stack>
          );
        })()}
      </Panel>
      <Panel
        isOpen={jsonOpen}
        type={PanelType.medium}
        headerText="Configuração em JSON"
        onDismiss={() => setJsonOpen(false)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 8 } }}>
          Cole um JSON completo do gestor de formulário e clique em Aplicar para carregar no painel. A gravação final
          continua a ser no botão Gravar do formulário.
        </Text>
        {jsonPanelErr && (
          <MessageBar messageBarType={MessageBarType.error} styles={{ root: { marginBottom: 8 } }}>
            {jsonPanelErr}
          </MessageBar>
        )}
        <TextField
          multiline
          rows={22}
          value={jsonPanelText}
          onChange={(_, v) => setJsonPanelText(v ?? '')}
          styles={{ root: { fontFamily: 'monospace', fontSize: 12 } }}
        />
        <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 12 } }}>
          <PrimaryButton text="Aplicar JSON" onClick={() => applyJsonFromPanel()} />
          <DefaultButton text="Fechar" onClick={() => setJsonOpen(false)} />
        </Stack>
      </Panel>
      {fieldPanelName && fieldPanelConfig && (
        <FormFieldRulesPanel
          isOpen={true}
          internalName={fieldPanelName}
          fieldConfig={fieldPanelConfig}
          meta={fieldPanelMeta}
          rules={rules}
          fieldOptions={fieldOptions}
          attachmentLibraryFolderOptions={attachmentFolderOptionsForFieldRules}
          lookupFieldsWebServerRelativeUrl={lw}
          listFieldMetadata={meta}
          allFieldConfigs={fields}
          onDismiss={() => setFieldPanelName(null)}
          onApply={(nextFc, editor) => {
            setFields((prev) =>
              prev.map((f) =>
                f.internalName === fieldPanelName
                  ? mergeFormFieldConfigFromRulesPanel(f, nextFc)
                  : f
              )
            );
            setRules((r) =>
              mergeFieldRules(
                r,
                fieldPanelName,
                buildFieldUiRules(fieldPanelName, editor, nextFc, {
                  mappedType: fieldPanelMeta?.MappedType ?? 'unknown',
                })
              )
            );
          }}
        />
      )}
    </Panel>
    </>
  );
};
