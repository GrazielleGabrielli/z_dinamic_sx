import type {
  IFormManagerAttachmentLibraryConfig,
  IFormManagerConfig,
  IFormManagerActionLogConfig,
  IFormLinkedChildFormConfig,
  IFormStepNavigationConfig,
  IFormFieldConfig,
  IFormSectionConfig,
  IFormStepConfig,
  IFormCustomButtonConfig,
  IFormCustomButtonConfirmConfig,
  TFormCustomButtonConfirmKind,
  TFormButtonAction,
  TFormCustomButtonOperation,
  TFormManagerFormMode,
  TFormRule,
  TFormConditionNode,
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
  TLinkedChildAttachmentStorageKind,
  TFormBannerPlacement,
  TChromePositionMode,
} from '../config/types/formManager';
import {
  FORM_BANNER_INTERNAL_PREFIX,
  FORM_FIXOS_STEP_ID,
  FORM_OCULTOS_STEP_ID,
} from '../config/types/formManager';
import { migrateFolderPathSegmentsToTree, sanitizeFolderTreeInput } from './attachmentFolderTree';
import { sanitizeConditionNode } from './formConditionSanitize';

const HISTORY_BUTTON_KIND_SET = new Set<string>(['text', 'icon', 'iconAndText']);

function pinOcultosFirstSections(sections: IFormSectionConfig[]): void {
  const oi = sections.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
  if (oi > 0) {
    const [oc] = sections.splice(oi, 1);
    sections.unshift(oc);
  }
}

function pinFixosAfterOcultosSections(sections: IFormSectionConfig[]): void {
  const fi = sections.findIndex((s) => s.id === FORM_FIXOS_STEP_ID);
  if (fi < 0) return;
  const oi = sections.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
  const wantIdx = oi >= 0 ? oi + 1 : 0;
  if (fi === wantIdx) return;
  const [fx] = sections.splice(fi, 1);
  const insertAt = fi < wantIdx ? wantIdx - 1 : wantIdx;
  sections.splice(insertAt, 0, fx);
}

function pinOcultosFirstSteps(steps: IFormStepConfig[]): void {
  const oi = steps.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
  if (oi > 0) {
    const [oc] = steps.splice(oi, 1);
    steps.unshift(oc);
  }
}

function pinFixosAfterOcultosSteps(steps: IFormStepConfig[]): void {
  const fi = steps.findIndex((s) => s.id === FORM_FIXOS_STEP_ID);
  if (fi < 0) return;
  const oi = steps.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
  const wantIdx = oi >= 0 ? oi + 1 : 0;
  if (fi === wantIdx) return;
  const [fx] = steps.splice(fi, 1);
  const insertAt = fi < wantIdx ? wantIdx - 1 : wantIdx;
  steps.splice(insertAt, 0, fx);
}

const STEP_LAYOUT_SET = new Set<string>([
  'rail',
  'segmented',
  'timeline',
  'cards',
  'breadcrumb',
  'underline',
  'outline',
  'compact',
  'steps',
  'minimal',
]);
const STEP_NAV_BUTTONS_SET = new Set<string>([
  'fluent',
  'pills',
  'dots',
  'icons',
  'links',
  'split',
  'stacked',
  'ghost',
  'toolbar',
  'compact',
]);
const BUTTON_OPERATION_SET = new Set<string>(['legacy', 'redirect', 'add', 'update', 'delete', 'history']);
const HISTORY_PRESENTATION_SET = new Set<string>(['panel', 'modal', 'collapse']);
const HISTORY_LAYOUT_SET = new Set<string>(['list', 'timeline', 'cards', 'compact']);
const FORM_DATA_LOADING_SET = new Set<string>(['spinner', 'spinnerLarge', 'shimmer', 'progress', 'cardShimmer']);
const FORM_SUBMIT_LOADING_SET = new Set<string>([
  'overlay',
  'topProgress',
  'formShimmer',
  'belowButtons',
  'infoBar',
]);
const ATTACHMENT_UPLOAD_LAYOUT_SET = new Set<string>([
  'default',
  'dropzone',
  'card',
  'ribbon',
  'compact',
]);
const ATTACHMENT_FILE_PREVIEW_SET = new Set<string>([
  'nameOnly',
  'nameAndSize',
  'iconAndName',
  'thumbnailAndName',
  'thumbnailLarge',
]);

const ACTION_SET = new Set<string>([
  'setVisibility', 'setRequired', 'setDisabled', 'setReadOnly', 'clearFields', 'setDefault',
  'validateValue', 'validateDate', 'atLeastOne', 'multiMinMax', 'showMessage', 'filterLookupOptions',
  'setComputed', 'profileVisibility', 'profileEditable', 'profileRequired', 'authorFieldAccess',
  'attachmentRules', 'asyncUniqueness', 'asyncCountLimit', 'setEffectiveSection',
]);

function sanitizeRule(raw: unknown): TFormRule | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const r = raw as Record<string, unknown>;
  const id = typeof r.id === 'string' ? r.id.trim() : '';
  const action = typeof r.action === 'string' ? r.action : '';
  if (!id || !ACTION_SET.has(action)) return undefined;
  const enabled = r.enabled === false ? false : true;
  const when = sanitizeConditionNode(r.when);
  const modes = Array.isArray(r.modes)
    ? (r.modes as string[]).filter((m) => m === 'create' || m === 'edit' || m === 'view') as ('create' | 'edit' | 'view')[]
    : undefined;
  const groupTitles = Array.isArray(r.groupTitles)
    ? (r.groupTitles as unknown[]).map((x) => String(x).trim()).filter(Boolean)
    : undefined;
  const tags = Array.isArray(r.tags)
    ? (r.tags as unknown[]).map((x) => String(x).trim()).filter(Boolean)
    : undefined;
  const base = { id, enabled, ...(when ? { when } : {}), ...(modes?.length ? { modes } : {}), ...(groupTitles?.length ? { groupTitles } : {}), ...(tags?.length ? { tags } : {}) };

  switch (action) {
    case 'setVisibility': {
      const targetKind = r.targetKind === 'section' ? 'section' : 'field';
      const targetId = typeof r.targetId === 'string' ? r.targetId.trim() : '';
      const visibility = r.visibility === 'hide' ? 'hide' : 'show';
      if (!targetId) return undefined;
      return { ...base, action: 'setVisibility', targetKind, targetId, visibility };
    }
    case 'setRequired': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      if (!field) return undefined;
      return { ...base, action: 'setRequired', field, required: r.required === true };
    }
    case 'setDisabled': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      if (!field) return undefined;
      return { ...base, action: 'setDisabled', field, disabled: r.disabled === true };
    }
    case 'setReadOnly': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      if (!field) return undefined;
      return { ...base, action: 'setReadOnly', field, readOnly: r.readOnly === true };
    }
    case 'clearFields': {
      const fields = Array.isArray(r.fields) ? (r.fields as unknown[]).map((x) => String(x).trim()).filter(Boolean) : [];
      if (!fields.length) return undefined;
      const triggerField = typeof r.triggerField === 'string' && r.triggerField.trim() ? r.triggerField.trim() : undefined;
      return { ...base, action: 'clearFields', fields, ...(triggerField ? { triggerField } : {}) };
    }
    case 'setDefault': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      const value = typeof r.value === 'string' ? r.value : String(r.value ?? '');
      if (!field) return undefined;
      return { ...base, action: 'setDefault', field, value };
    }
    case 'validateValue': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      if (!field) return undefined;
      return {
        ...base,
        action: 'validateValue',
        field,
        ...(typeof r.minNumber === 'number' ? { minNumber: r.minNumber } : {}),
        ...(typeof r.maxNumber === 'number' ? { maxNumber: r.maxNumber } : {}),
        ...(typeof r.minLength === 'number' ? { minLength: r.minLength } : {}),
        ...(typeof r.maxLength === 'number' ? { maxLength: r.maxLength } : {}),
        ...(typeof r.pattern === 'string' && r.pattern.trim() ? { pattern: r.pattern.trim(), patternMessage: typeof r.patternMessage === 'string' ? r.patternMessage : undefined } : {}),
        ...(Array.isArray(r.allowList) ? { allowList: (r.allowList as unknown[]).map((x) => String(x)) } : {}),
        ...(Array.isArray(r.denyList) ? { denyList: (r.denyList as unknown[]).map((x) => String(x)) } : {}),
        ...(typeof r.message === 'string' ? { message: r.message } : {}),
      };
    }
    case 'validateDate': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      if (!field) return undefined;
      return {
        ...base,
        action: 'validateDate',
        field,
        ...(typeof r.minIso === 'string' ? { minIso: r.minIso } : {}),
        ...(typeof r.maxIso === 'string' ? { maxIso: r.maxIso } : {}),
        ...(typeof r.minDaysFromToday === 'number' ? { minDaysFromToday: r.minDaysFromToday } : {}),
        ...(typeof r.maxDaysFromToday === 'number' ? { maxDaysFromToday: r.maxDaysFromToday } : {}),
        ...(r.blockWeekends === true ? { blockWeekends: true } : {}),
        ...(Array.isArray(r.blockedIsoDates) ? { blockedIsoDates: (r.blockedIsoDates as unknown[]).map((x) => String(x)) } : {}),
        ...(typeof r.gteField === 'string' ? { gteField: r.gteField.trim() } : {}),
        ...(typeof r.lteField === 'string' ? { lteField: r.lteField.trim() } : {}),
        ...(typeof r.gtField === 'string' ? { gtField: r.gtField.trim() } : {}),
        ...(typeof r.ltField === 'string' ? { ltField: r.ltField.trim() } : {}),
        ...(typeof r.message === 'string' ? { message: r.message } : {}),
      };
    }
    case 'atLeastOne': {
      const fields = Array.isArray(r.fields) ? (r.fields as unknown[]).map((x) => String(x).trim()).filter(Boolean) : [];
      if (!fields.length) return undefined;
      return { ...base, action: 'atLeastOne', fields, ...(typeof r.message === 'string' ? { message: r.message } : {}) };
    }
    case 'multiMinMax': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      if (!field) return undefined;
      return {
        ...base,
        action: 'multiMinMax',
        field,
        ...(typeof r.min === 'number' ? { min: r.min } : {}),
        ...(typeof r.max === 'number' ? { max: r.max } : {}),
        ...(typeof r.message === 'string' ? { message: r.message } : {}),
      };
    }
    case 'showMessage': {
      const text = typeof r.text === 'string' ? r.text : '';
      if (!text.trim()) return undefined;
      const variant = r.variant === 'warning' || r.variant === 'error' ? r.variant : 'info';
      return { ...base, action: 'showMessage', variant, text };
    }
    case 'filterLookupOptions': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      const parentField = typeof r.parentField === 'string' ? r.parentField.trim() : '';
      const odataFilterTemplate = typeof r.odataFilterTemplate === 'string' ? r.odataFilterTemplate : '';
      if (!field || !parentField || !odataFilterTemplate.trim()) return undefined;
      return { ...base, action: 'filterLookupOptions', field, parentField, odataFilterTemplate };
    }
    case 'setComputed': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      const expression = typeof r.expression === 'string' ? r.expression : '';
      if (!field || !expression.trim()) return undefined;
      return { ...base, action: 'setComputed', field, expression };
    }
    case 'profileVisibility':
    case 'profileEditable':
    case 'profileRequired': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      const gt = Array.isArray(r.groupTitles) ? (r.groupTitles as unknown[]).map((x) => String(x).trim()).filter(Boolean) : [];
      if (!field || !gt.length) return undefined;
      return { ...base, action, field, groupTitles: gt, allow: r.allow === true };
    }
    case 'authorFieldAccess': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      if (!field) return undefined;
      return { ...base, action: 'authorFieldAccess', field };
    }
    case 'attachmentRules': {
      const whenAtt = sanitizeConditionNode(r.requiredWhen);
      const allowedFileExtensions = Array.isArray((r as { allowedFileExtensions?: unknown }).allowedFileExtensions)
        ? (r as { allowedFileExtensions: unknown[] }).allowedFileExtensions
            .map((x) => String(x).trim().replace(/^\./, '').toLowerCase())
            .filter(Boolean)
        : [];
      return {
        ...base,
        action: 'attachmentRules',
        ...(typeof r.minCount === 'number' ? { minCount: r.minCount } : {}),
        ...(typeof r.maxCount === 'number' ? { maxCount: r.maxCount } : {}),
        ...(typeof r.maxBytesPerFile === 'number' ? { maxBytesPerFile: r.maxBytesPerFile } : {}),
        ...(allowedFileExtensions.length ? { allowedFileExtensions } : {}),
        ...(Array.isArray(r.allowedMimeTypes) ? { allowedMimeTypes: (r.allowedMimeTypes as unknown[]).map((x) => String(x)) } : {}),
        ...(whenAtt ? { requiredWhen: whenAtt } : {}),
        ...(typeof r.message === 'string' ? { message: r.message } : {}),
      };
    }
    case 'asyncUniqueness': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      if (!field) return undefined;
      return {
        ...base,
        action: 'asyncUniqueness',
        field,
        ...(typeof r.listTitle === 'string' && r.listTitle.trim() ? { listTitle: r.listTitle.trim() } : {}),
        ...(typeof r.message === 'string' ? { message: r.message } : {}),
      };
    }
    case 'asyncCountLimit': {
      const filterTemplate = typeof r.filterTemplate === 'string' ? r.filterTemplate : '';
      const maxCount = typeof r.maxCount === 'number' ? r.maxCount : 0;
      if (!filterTemplate.trim() || maxCount < 1) return undefined;
      return {
        ...base,
        action: 'asyncCountLimit',
        filterTemplate,
        maxCount,
        ...(typeof r.listTitle === 'string' && r.listTitle.trim() ? { listTitle: r.listTitle.trim() } : {}),
        ...(typeof r.message === 'string' ? { message: r.message } : {}),
      };
    }
    case 'setEffectiveSection': {
      const field = typeof r.field === 'string' ? r.field.trim() : '';
      const sectionId = typeof r.sectionId === 'string' ? r.sectionId.trim() : '';
      if (!field || !sectionId) return undefined;
      return { ...base, action: 'setEffectiveSection', field, sectionId };
    }
    default:
      return undefined;
  }
}

function sanitizeField(raw: unknown): IFormFieldConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const f = raw as Record<string, unknown>;
  const internalName = typeof f.internalName === 'string' ? f.internalName.trim() : '';
  if (!internalName) return undefined;
  const fpRaw = f.fixedPlacement;
  const fixedPl = fpRaw === 'top' || fpRaw === 'bottom' ? fpRaw : undefined;
  const cmRaw = f.chromePositionMode;
  const chromeMode: TChromePositionMode | undefined =
    cmRaw === 'sticky' || cmRaw === 'absolute' || cmRaw === 'flow' ? cmRaw : undefined;
  const isBanner =
    f.fieldKind === 'banner' || internalName.indexOf(FORM_BANNER_INTERNAL_PREFIX) === 0;
  const bannerUrlRaw = typeof f.bannerImageUrl === 'string' ? f.bannerImageUrl.trim() : '';
  const bannerImageUrl = bannerUrlRaw ? bannerUrlRaw.slice(0, 4000) : undefined;
  const common: IFormFieldConfig = {
    internalName,
    ...(fixedPl ? { fixedPlacement: fixedPl } : {}),
    ...(chromeMode ? { chromePositionMode: chromeMode } : {}),
    ...(typeof f.label === 'string' ? { label: f.label } : {}),
    ...(typeof f.helpText === 'string' ? { helpText: f.helpText } : {}),
    ...(typeof f.placeholder === 'string' ? { placeholder: f.placeholder } : {}),
    ...(typeof f.sectionId === 'string' ? { sectionId: f.sectionId.trim() } : {}),
    ...(f.visible === false ? { visible: false } : {}),
    ...(f.required === true ? { required: true } : {}),
    ...(f.disabled === true ? { disabled: true } : {}),
    ...(f.readOnly === true ? { readOnly: true } : {}),
    ...(f.width === 'half' ? { width: 'half' } : {}),
    ...(typeof f.modalGroupId === 'string' ? { modalGroupId: f.modalGroupId.trim() } : {}),
    ...(typeof f.effectiveSectionId === 'string' ? { effectiveSectionId: f.effectiveSectionId.trim() } : {}),
  };
  if (isBanner) {
    const bp = f.bannerPlacement;
    const placement: TFormBannerPlacement | undefined =
      bp === 'topFixed' || bp === 'bottomFixed' || bp === 'inStep' ? bp : undefined;
    const bw = typeof f.bannerWidthPercent === 'number' && isFinite(f.bannerWidthPercent)
      ? Math.min(100, Math.max(1, f.bannerWidthPercent))
      : undefined;
    const bh = typeof f.bannerHeightPercent === 'number' && isFinite(f.bannerHeightPercent)
      ? Math.min(100, Math.max(1, f.bannerHeightPercent))
      : undefined;
    return {
      ...common,
      fieldKind: 'banner',
      ...(bannerImageUrl ? { bannerImageUrl } : {}),
      ...(placement ? { bannerPlacement: placement } : {}),
      ...(bw !== undefined ? { bannerWidthPercent: bw } : {}),
      ...(bh !== undefined ? { bannerHeightPercent: bh } : {}),
    };
  }
  return common;
}

function sanitizeSection(raw: unknown): IFormSectionConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const s = raw as Record<string, unknown>;
  const id = typeof s.id === 'string' ? s.id.trim() : '';
  const title = typeof s.title === 'string' ? s.title.trim() : '';
  if (!id || !title) return undefined;
  return {
    id,
    title,
    ...(s.visible === false ? { visible: false } : {}),
    ...(s.collapsed === true ? { collapsed: true } : {}),
  };
}

function sanitizeButtonAction(raw: unknown): TFormButtonAction | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const a = raw as Record<string, unknown>;
  const kind = a.kind;
  const whenAct = sanitizeConditionNode(a.when);
  if (kind === 'showFields' || kind === 'hideFields') {
    const fields = Array.isArray(a.fields)
      ? (a.fields as unknown[]).map((x) => String(x).trim()).filter(Boolean)
      : [];
    if (!fields.length) return undefined;
    const displayOnStepId =
      kind === 'showFields' && typeof a.displayOnStepId === 'string' && a.displayOnStepId.trim()
        ? a.displayOnStepId.trim()
        : undefined;
    return kind === 'showFields'
      ? {
          kind: 'showFields',
          fields,
          ...(displayOnStepId ? { displayOnStepId } : {}),
          ...(whenAct ? { when: whenAct } : {}),
        }
      : { kind: 'hideFields', fields, ...(whenAct ? { when: whenAct } : {}) };
  }
  if (kind === 'setFieldValue') {
    const field = typeof a.field === 'string' ? a.field.trim() : '';
    const valueTemplate = typeof a.valueTemplate === 'string' ? a.valueTemplate : '';
    if (!field) return undefined;
    return { kind: 'setFieldValue', field, valueTemplate, ...(whenAct ? { when: whenAct } : {}) };
  }
  if (kind === 'joinFields') {
    const targetField = typeof a.targetField === 'string' ? a.targetField.trim() : '';
    const sourceFields = Array.isArray(a.sourceFields)
      ? (a.sourceFields as unknown[]).map((x) => String(x).trim()).filter(Boolean)
      : [];
    const separator = typeof a.separator === 'string' ? a.separator : ' ';
    const valueTemplateRaw = typeof a.valueTemplate === 'string' ? a.valueTemplate : '';
    const hasTemplate = valueTemplateRaw.trim().length > 0;
    if (!targetField) return undefined;
    if (hasTemplate) {
      return {
        kind: 'joinFields',
        targetField,
        valueTemplate: valueTemplateRaw,
        sourceFields,
        separator,
        ...(whenAct ? { when: whenAct } : {}),
      };
    }
    if (!sourceFields.length) return undefined;
    return {
      kind: 'joinFields',
      targetField,
      sourceFields,
      separator,
      ...(whenAct ? { when: whenAct } : {}),
    };
  }
  return undefined;
}

function sanitizeCustomButton(raw: unknown): IFormCustomButtonConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const b = raw as Record<string, unknown>;
  const id = typeof b.id === 'string' ? b.id.trim() : '';
  if (!id) return undefined;
  const labelRaw = typeof b.label === 'string' ? b.label.trim() : '';
  const label = labelRaw || id;
  const appearance = b.appearance === 'primary' ? 'primary' : 'default';
  const behaviorRaw = b.behavior;
  const behavior: IFormCustomButtonConfig['behavior'] =
    behaviorRaw === 'draft'
      ? 'draft'
      : behaviorRaw === 'submit'
        ? 'submit'
        : behaviorRaw === 'close'
          ? 'close'
          : 'actionsOnly';
  const opRaw = b.operation;
  const operation: TFormCustomButtonOperation | undefined =
    typeof opRaw === 'string' && BUTTON_OPERATION_SET.has(opRaw) ? (opRaw as TFormCustomButtonOperation) : undefined;
  const redirectUrlTemplate =
    typeof b.redirectUrlTemplate === 'string' ? b.redirectUrlTemplate : undefined;
  const deleteShowInView = b.deleteShowInView === false ? false : undefined;
  const deleteShowInEdit = b.deleteShowInEdit === false ? false : undefined;
  const modes = Array.isArray(b.modes)
    ? (b.modes as string[]).filter((m) => m === 'create' || m === 'edit' || m === 'view') as TFormManagerFormMode[]
    : undefined;
  const enabled = b.enabled === false ? false : true;
  const when = sanitizeConditionNode(b.when);
  const groupTitles = Array.isArray(b.groupTitles)
    ? (b.groupTitles as unknown[]).map((x) => String(x).trim()).filter(Boolean)
    : undefined;
  const showOnlyWhenAllRequiredFilled = b.showOnlyWhenAllRequiredFilled === true ? true : undefined;
  const shortDescriptionRaw = typeof b.shortDescription === 'string' ? b.shortDescription.trim() : '';
  const shortDescription = shortDescriptionRaw ? shortDescriptionRaw : undefined;
  const slKindRaw = b.submitLoadingKind;
  const submitLoadingKind: TFormSubmitLoadingUiKind | undefined =
    typeof slKindRaw === 'string' && FORM_SUBMIT_LOADING_SET.has(slKindRaw)
      ? (slKindRaw as TFormSubmitLoadingUiKind)
      : undefined;
  const actionsRaw = Array.isArray(b.actions) ? b.actions : [];
  const actionsSan: TFormButtonAction[] = [];
  for (let i = 0; i < actionsRaw.length; i++) {
    const act = sanitizeButtonAction(actionsRaw[i]);
    if (act) actionsSan.push(act);
  }
  const opResolved = operation ?? 'legacy';
  const actions: TFormButtonAction[] =
    opResolved === 'redirect' || opResolved === 'history' ? [] : actionsSan;
  const confirmRaw = b.confirmBeforeRun;
  let confirmBeforeRun: IFormCustomButtonConfirmConfig | undefined;
  if (confirmRaw && typeof confirmRaw === 'object') {
    const c = confirmRaw as Record<string, unknown>;
    const enabled = c.enabled === true;
    const msg = typeof c.message === 'string' ? c.message.trim() : '';
    const kindRaw = c.kind;
    const kind: TFormCustomButtonConfirmKind =
      kindRaw === 'success' ||
      kindRaw === 'warning' ||
      kindRaw === 'error' ||
      kindRaw === 'blocked' ||
      kindRaw === 'info'
        ? (kindRaw as TFormCustomButtonConfirmKind)
        : 'info';
    if (enabled && msg) {
      confirmBeforeRun = { enabled: true, kind, message: msg };
    }
  }
  return {
    id,
    label,
    appearance,
    behavior,
    ...(operation && operation !== 'legacy' ? { operation } : {}),
    ...(shortDescription && opResolved === 'history' ? { shortDescription } : {}),
    ...(redirectUrlTemplate !== undefined && redirectUrlTemplate.trim() ? { redirectUrlTemplate } : {}),
    ...(deleteShowInView === false ? { deleteShowInView: false } : {}),
    ...(deleteShowInEdit === false ? { deleteShowInEdit: false } : {}),
    ...(modes?.length ? { modes } : {}),
    ...(enabled === false ? { enabled: false } : {}),
    ...(when ? { when } : {}),
    ...(groupTitles?.length ? { groupTitles } : {}),
    ...(showOnlyWhenAllRequiredFilled ? { showOnlyWhenAllRequiredFilled: true } : {}),
    ...(submitLoadingKind ? { submitLoadingKind } : {}),
    ...(confirmBeforeRun ? { confirmBeforeRun } : {}),
    actions,
  };
}

function coerceBoolTrue(v: unknown): boolean {
  return v === true || v === 'true' || v === 1 || v === '1';
}

function coerceBoolFalse(v: unknown): boolean {
  return v === false || v === 'false' || v === 0 || v === '0';
}

function sanitizeStepNavigation(raw: unknown): IFormStepNavigationConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const o = raw as Record<string, unknown>;
  const out: IFormStepNavigationConfig = {};
  if (coerceBoolTrue(o.requireFilledRequiredToAdvance)) out.requireFilledRequiredToAdvance = true;
  if (coerceBoolTrue(o.fullValidationOnAdvance)) out.fullValidationOnAdvance = true;
  if (coerceBoolFalse(o.allowBackWithoutValidation)) out.allowBackWithoutValidation = false;
  if (Object.keys(out).length === 0) return undefined;
  return out;
}

const FORM_MANAGER_MODES = ['create', 'edit', 'view'] as const;

function sanitizeShowInFormModes(raw: unknown): TFormManagerFormMode[] | undefined {
  if (!Array.isArray(raw)) return undefined;
  const out: TFormManagerFormMode[] = [];
  for (let i = 0; i < raw.length; i++) {
    const v = String(raw[i]).trim();
    if (FORM_MANAGER_MODES.indexOf(v as (typeof FORM_MANAGER_MODES)[number]) !== -1) {
      out.push(v as TFormManagerFormMode);
    }
  }
  const uniq = out.filter((m, j) => out.indexOf(m) === j);
  if (uniq.length === 0) return undefined;
  if (uniq.length === 3) return undefined;
  return uniq.sort(
    (a, b) =>
      FORM_MANAGER_MODES.indexOf(a as (typeof FORM_MANAGER_MODES)[number]) -
      FORM_MANAGER_MODES.indexOf(b as (typeof FORM_MANAGER_MODES)[number])
  );
}

function sanitizeStep(raw: unknown): IFormStepConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const s = raw as Record<string, unknown>;
  const id = typeof s.id === 'string' ? s.id.trim() : '';
  const title = typeof s.title === 'string' ? s.title.trim() : '';
  const fieldNames = Array.isArray(s.fieldNames) ? (s.fieldNames as unknown[]).map((x) => String(x).trim()).filter(Boolean) : [];
  if (!id || !title) return undefined;
  const showInFormModes = sanitizeShowInFormModes(s.showInFormModes);
  return { id, title, fieldNames, ...(showInFormModes?.length ? { showInFormModes } : {}) };
}

const MAX_ATTACHMENT_FOLDER_SEGMENTS = 10;
const MAX_ATTACHMENT_FOLDER_TEMPLATE_CHARS = 200;
const MAX_LINKED_CHILD_FORMS = 10;

const LINKED_CHILD_ATTACH_KIND_SET = new Set<string>([
  'none',
  'itemAttachments',
  'documentLibraryInheritMain',
  'documentLibraryCustom',
]);

function sanitizeLinkedChildAttachmentStorageKind(raw: unknown): TLinkedChildAttachmentStorageKind | undefined {
  const s = typeof raw === 'string' ? raw.trim() : '';
  if (!s || !LINKED_CHILD_ATTACH_KIND_SET.has(s)) return undefined;
  return s as TLinkedChildAttachmentStorageKind;
}

function sanitizeFormBodySubset(o: Record<string, unknown>): Pick<
  IFormLinkedChildFormConfig,
  'sections' | 'fields' | 'rules' | 'steps'
> {
  const sectionsRaw = Array.isArray(o.sections) ? o.sections : [];
  const fieldsRaw = Array.isArray(o.fields) ? o.fields : [];
  const rulesRaw = Array.isArray(o.rules) ? o.rules : [];
  const stepsRaw = Array.isArray(o.steps) ? o.steps : [];
  const sections: IFormSectionConfig[] = [];
  for (let i = 0; i < sectionsRaw.length; i++) {
    const sec = sanitizeSection(sectionsRaw[i]);
    if (sec) sections.push(sec);
  }
  if (sections.length === 0) sections.push({ id: 'main', title: 'Geral', visible: true });
  if (!sections.some((s) => s.id === FORM_OCULTOS_STEP_ID)) {
    sections.unshift({ id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', visible: true });
  }
  pinOcultosFirstSections(sections);
  if (!sections.some((s) => s.id === FORM_FIXOS_STEP_ID)) {
    const oi = sections.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
    sections.splice(oi >= 0 ? oi + 1 : 0, 0, { id: FORM_FIXOS_STEP_ID, title: 'Fixos', visible: true });
  } else {
    pinFixosAfterOcultosSections(sections);
  }
  const fields: IFormFieldConfig[] = [];
  for (let i = 0; i < fieldsRaw.length; i++) {
    const fc = sanitizeField(fieldsRaw[i]);
    if (fc) fields.push(fc);
  }
  const rules: TFormRule[] = [];
  for (let i = 0; i < rulesRaw.length; i++) {
    const rule = sanitizeRule(rulesRaw[i]);
    if (rule) rules.push(rule);
  }
  const steps: IFormStepConfig[] = [];
  for (let i = 0; i < stepsRaw.length; i++) {
    const st = sanitizeStep(stepsRaw[i]);
    if (st) steps.push(st);
  }
  if (steps.length > 0 && !steps.some((s) => s.id === FORM_OCULTOS_STEP_ID)) {
    steps.unshift({ id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] });
  }
  if (steps.length > 0) {
    pinOcultosFirstSteps(steps);
  }
  if (steps.length > 0 && !steps.some((s) => s.id === FORM_FIXOS_STEP_ID)) {
    const oi = steps.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
    steps.splice(oi >= 0 ? oi + 1 : 0, 0, { id: FORM_FIXOS_STEP_ID, title: 'Fixos', fieldNames: [] });
  } else if (steps.length > 0) {
    pinFixosAfterOcultosSteps(steps);
  }
  return {
    sections,
    fields,
    rules,
    ...(steps.length ? { steps } : {}),
  };
}

function sanitizeLinkedChildId(raw: unknown): string {
  if (typeof raw === 'string') {
    const t = raw.trim();
    if (t) return t.slice(0, 200);
  }
  if (typeof raw === 'number' && isFinite(raw)) {
    const t = String(Math.trunc(raw));
    if (t) return t.slice(0, 200);
  }
  return '';
}

function sanitizeLinkedChildFormConfig(raw: unknown): IFormLinkedChildFormConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const o = raw as Record<string, unknown>;
  const id = sanitizeLinkedChildId(o.id);
  if (!id) return undefined;
  const listTitle =
    typeof o.listTitle === 'string'
      ? o.listTitle.trim()
      : String(o.listTitle ?? '')
          .trim()
          .slice(0, 500);
  const parentLookupFieldInternalName =
    typeof o.parentLookupFieldInternalName === 'string'
      ? o.parentLookupFieldInternalName.trim()
      : String(o.parentLookupFieldInternalName ?? '')
          .trim()
          .slice(0, 500);
  const body = sanitizeFormBodySubset(o);
  const titleRaw =
    typeof o.title === 'string' ? o.title.trim() : String(o.title ?? '').trim();
  const title = titleRaw ? titleRaw.slice(0, 200) : undefined;
  let order: number | undefined;
  if (typeof o.order === 'number' && isFinite(o.order)) order = Math.round(o.order);
  else if (typeof o.order === 'string' && o.order.trim()) {
    const n = Number(o.order);
    if (isFinite(n)) order = Math.round(n);
  }
  let minRows: number | undefined;
  if (typeof o.minRows === 'number' && isFinite(o.minRows)) {
    const m = Math.round(o.minRows);
    if (m >= 0 && m <= 500) minRows = m;
  }
  let maxRows: number | undefined;
  if (typeof o.maxRows === 'number' && isFinite(o.maxRows)) {
    const m = Math.round(o.maxRows);
    if (m >= 0 && m <= 500) maxRows = m;
  }
  if (minRows !== undefined && maxRows !== undefined && maxRows < minRows) maxRows = minRows;
  const collapsedDefault = o.collapsedDefault === true ? true : undefined;
  let mainFormStepId: string | undefined;
  if (typeof o.mainFormStepId === 'string') {
    const t = o.mainFormStepId.trim();
    if (t) mainFormStepId = t.slice(0, 120);
  }
  let childAttachmentStorageKind = sanitizeLinkedChildAttachmentStorageKind(o.childAttachmentStorageKind);
  let childAttachmentLibraryLookupToChildListField: string | undefined;
  if (typeof o.childAttachmentLibraryLookupToChildListField === 'string') {
    const t = o.childAttachmentLibraryLookupToChildListField.trim();
    if (t) childAttachmentLibraryLookupToChildListField = t.slice(0, 255);
  }
  const childLibRaw = sanitizeAttachmentLibrary(o.childAttachmentLibrary);
  if (childAttachmentStorageKind === 'documentLibraryInheritMain' && !childAttachmentLibraryLookupToChildListField) {
    childAttachmentStorageKind = undefined;
  }
  if (childAttachmentStorageKind === 'documentLibraryCustom') {
    if (!childLibRaw?.libraryTitle || !childLibRaw.sourceListLookupFieldInternalName) {
      childAttachmentStorageKind = undefined;
    }
  }
  return {
    id,
    listTitle,
    parentLookupFieldInternalName,
    ...body,
    ...(title ? { title } : {}),
    ...(order !== undefined ? { order } : {}),
    ...(minRows !== undefined ? { minRows } : {}),
    ...(maxRows !== undefined ? { maxRows } : {}),
    ...(collapsedDefault ? { collapsedDefault: true } : {}),
    ...(mainFormStepId ? { mainFormStepId } : {}),
    ...(childAttachmentStorageKind && childAttachmentStorageKind !== 'none'
      ? { childAttachmentStorageKind }
      : {}),
    ...(childAttachmentLibraryLookupToChildListField
      ? { childAttachmentLibraryLookupToChildListField }
      : {}),
    ...(childAttachmentStorageKind === 'documentLibraryCustom' && childLibRaw
      ? { childAttachmentLibrary: childLibRaw }
      : {}),
  };
}

function stripLeadingRedundantItemIdTemplate(segments: string[]): string[] {
  const out = segments.slice();
  while (out.length > 0 && /^\{\{\s*ItemId\s*\}\}$/i.test(out[0].trim())) {
    out.shift();
  }
  return out;
}

function sanitizeAttachmentLibrary(raw: unknown): IFormManagerAttachmentLibraryConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const o = raw as Record<string, unknown>;
  const libraryTitle = typeof o.libraryTitle === 'string' ? o.libraryTitle.trim() : '';
  const sourceListLookupFieldInternalName =
    typeof o.sourceListLookupFieldInternalName === 'string'
      ? o.sourceListLookupFieldInternalName.trim()
      : '';
  let folderPathSegments: string[] | undefined;
  if (Array.isArray(o.folderPathSegments)) {
    const seg: string[] = [];
    for (let i = 0; i < o.folderPathSegments.length && i < MAX_ATTACHMENT_FOLDER_SEGMENTS; i++) {
      const t = typeof o.folderPathSegments[i] === 'string' ? String(o.folderPathSegments[i]).trim() : '';
      if (!t) continue;
      seg.push(t.slice(0, MAX_ATTACHMENT_FOLDER_TEMPLATE_CHARS));
    }
    const stripped = stripLeadingRedundantItemIdTemplate(seg);
    if (stripped.length) folderPathSegments = stripped;
  }
  let folderTree = sanitizeFolderTreeInput(o.folderTree);
  if (!folderTree.length && folderPathSegments?.length) {
    folderTree = sanitizeFolderTreeInput(migrateFolderPathSegmentsToTree(folderPathSegments));
  }
  if (!libraryTitle && !sourceListLookupFieldInternalName && !folderTree.length) return undefined;
  return {
    ...(libraryTitle ? { libraryTitle } : {}),
    ...(sourceListLookupFieldInternalName ? { sourceListLookupFieldInternalName } : {}),
    ...(folderTree.length ? { folderTree } : {}),
  };
}

function sanitizeActionLog(raw: unknown): IFormManagerActionLogConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const o = raw as Record<string, unknown>;
  const listTitle = typeof o.listTitle === 'string' ? o.listTitle.trim() : '';
  const actionFieldInternalName =
    typeof o.actionFieldInternalName === 'string' ? o.actionFieldInternalName.trim() : '';
  const sourceListLookupFieldInternalName =
    typeof o.sourceListLookupFieldInternalName === 'string'
      ? o.sourceListLookupFieldInternalName.trim()
      : '';
  let captureEnabled = o.captureEnabled === true;
  if (captureEnabled && (!listTitle || !actionFieldInternalName || !sourceListLookupFieldInternalName)) {
    captureEnabled = false;
  }
  const descRaw = o.descriptionsHtmlByButtonId;
  const descriptionsHtmlByButtonId: Record<string, string> = {};
  if (descRaw && typeof descRaw === 'object' && !Array.isArray(descRaw)) {
    const entries = Object.entries(descRaw as Record<string, unknown>);
    for (let i = 0; i < entries.length; i++) {
      const k = entries[i][0];
      const v = entries[i][1];
      const id = String(k).trim();
      if (!id) continue;
      const html = typeof v === 'string' ? v : '';
      if (html.trim()) descriptionsHtmlByButtonId[id] = html;
    }
  }
  if (
    !captureEnabled &&
    !listTitle &&
    !actionFieldInternalName &&
    !sourceListLookupFieldInternalName &&
    Object.keys(descriptionsHtmlByButtonId).length === 0
  ) {
    return undefined;
  }
  return {
    ...(captureEnabled ? { captureEnabled: true } : {}),
    ...(listTitle ? { listTitle } : {}),
    ...(actionFieldInternalName ? { actionFieldInternalName } : {}),
    ...(sourceListLookupFieldInternalName ? { sourceListLookupFieldInternalName } : {}),
    ...(Object.keys(descriptionsHtmlByButtonId).length ? { descriptionsHtmlByButtonId } : {}),
  };
}

export function sanitizeFormManagerConfig(raw: unknown): IFormManagerConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const o = raw as Record<string, unknown>;
  const sectionsRaw = Array.isArray(o.sections) ? o.sections : [];
  const fieldsRaw = Array.isArray(o.fields) ? o.fields : [];
  const rulesRaw = Array.isArray(o.rules) ? o.rules : [];
  const stepsRaw = Array.isArray(o.steps) ? o.steps : [];
  const sections: IFormSectionConfig[] = [];
  for (let i = 0; i < sectionsRaw.length; i++) {
    const sec = sanitizeSection(sectionsRaw[i]);
    if (sec) sections.push(sec);
  }
  if (sections.length === 0) sections.push({ id: 'main', title: 'Geral', visible: true });
  if (!sections.some((s) => s.id === FORM_OCULTOS_STEP_ID)) {
    sections.unshift({ id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', visible: true });
  }
  pinOcultosFirstSections(sections);
  if (!sections.some((s) => s.id === FORM_FIXOS_STEP_ID)) {
    const oi = sections.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
    sections.splice(oi >= 0 ? oi + 1 : 0, 0, { id: FORM_FIXOS_STEP_ID, title: 'Fixos', visible: true });
  } else {
    pinFixosAfterOcultosSections(sections);
  }
  const fields: IFormFieldConfig[] = [];
  for (let i = 0; i < fieldsRaw.length; i++) {
    const fc = sanitizeField(fieldsRaw[i]);
    if (fc) fields.push(fc);
  }
  const rules: TFormRule[] = [];
  for (let i = 0; i < rulesRaw.length; i++) {
    const rule = sanitizeRule(rulesRaw[i]);
    if (rule) rules.push(rule);
  }
  const steps: IFormStepConfig[] = [];
  for (let i = 0; i < stepsRaw.length; i++) {
    const st = sanitizeStep(stepsRaw[i]);
    if (st) steps.push(st);
  }
  if (steps.length > 0 && !steps.some((s) => s.id === FORM_OCULTOS_STEP_ID)) {
    steps.unshift({ id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] });
  }
  if (steps.length > 0) {
    pinOcultosFirstSteps(steps);
  }
  if (steps.length > 0 && !steps.some((s) => s.id === FORM_FIXOS_STEP_ID)) {
    const oi = steps.findIndex((s) => s.id === FORM_OCULTOS_STEP_ID);
    steps.splice(oi >= 0 ? oi + 1 : 0, 0, { id: FORM_FIXOS_STEP_ID, title: 'Fixos', fieldNames: [] });
  } else if (steps.length > 0) {
    pinFixosAfterOcultosSteps(steps);
  }
  const managerColumnFields = Array.isArray(o.managerColumnFields)
    ? (o.managerColumnFields as unknown[]).map((x) => String(x).trim()).filter(Boolean)
    : undefined;
  const dynamicHelpRaw = Array.isArray(o.dynamicHelp) ? o.dynamicHelp : [];
  const dynamicHelp: { field: string; when: TFormConditionNode; helpText: string }[] = [];
  for (let i = 0; i < dynamicHelpRaw.length; i++) {
    const dh = dynamicHelpRaw[i];
    if (!dh || typeof dh !== 'object') continue;
    const d = dh as Record<string, unknown>;
    const field = typeof d.field === 'string' ? d.field.trim() : '';
    const helpText = typeof d.helpText === 'string' ? d.helpText : '';
    const when = sanitizeConditionNode(d.when);
    if (field && helpText.trim() && when) dynamicHelp.push({ field, when, helpText });
  }
  const customButtonsRaw = Array.isArray(o.customButtons) ? o.customButtons : [];
  const customButtons: IFormCustomButtonConfig[] = [];
  for (let i = 0; i < customButtonsRaw.length; i++) {
    const btn = sanitizeCustomButton(customButtonsRaw[i]);
    if (btn) customButtons.push(btn);
  }
  const slRaw = o.stepLayout;
  const stepLayout: TFormStepLayoutKind | undefined =
    typeof slRaw === 'string' && STEP_LAYOUT_SET.has(slRaw) ? (slRaw as TFormStepLayoutKind) : undefined;
  const snRaw = o.stepNavButtons;
  const stepNavButtons: TFormStepNavButtonsKind | undefined =
    typeof snRaw === 'string' && STEP_NAV_BUTTONS_SET.has(snRaw) ? (snRaw as TFormStepNavButtonsKind) : undefined;
  const fdlRaw = o.formDataLoadingKind;
  const formDataLoadingKind: TFormDataLoadingUiKind | undefined =
    typeof fdlRaw === 'string' && FORM_DATA_LOADING_SET.has(fdlRaw)
      ? (fdlRaw as TFormDataLoadingUiKind)
      : undefined;
  const dslRaw = o.defaultSubmitLoadingKind;
  const defaultSubmitLoadingKind: TFormSubmitLoadingUiKind | undefined =
    typeof dslRaw === 'string' && FORM_SUBMIT_LOADING_SET.has(dslRaw)
      ? (dslRaw as TFormSubmitLoadingUiKind)
      : undefined;
  const frwmRaw = o.formRootWidthMode;
  const formRootWidthMode: TFormRootWidthMode | undefined =
    frwmRaw === 'full' || frwmRaw === 'percent' ? frwmRaw : undefined;
  const frpRaw = o.formRootWidthPercent;
  const formRootWidthPercent: number | undefined =
    typeof frpRaw === 'number' && isFinite(frpRaw)
      ? Math.min(100, Math.max(1, Math.round(frpRaw)))
      : undefined;
  const frhaRaw = o.formRootHorizontalAlign;
  const formRootHorizontalAlign: TFormRootHorizontalAlign | undefined =
    frhaRaw === 'start' || frhaRaw === 'center' || frhaRaw === 'end' ? frhaRaw : undefined;
  const frppRaw = o.formRootPaddingPx;
  let formRootPaddingPx: number | undefined;
  if (typeof frppRaw === 'number' && isFinite(frppRaw)) {
    const r = Math.round(frppRaw);
    if (r >= 1 && r <= 160) formRootPaddingPx = r;
  }
  const stepNavigation = sanitizeStepNavigation(o.stepNavigation);
  const attLayoutRaw = o.attachmentUploadLayout;
  const attachmentUploadLayout: TFormAttachmentUploadLayoutKind | undefined =
    typeof attLayoutRaw === 'string' && ATTACHMENT_UPLOAD_LAYOUT_SET.has(attLayoutRaw)
      ? (attLayoutRaw as TFormAttachmentUploadLayoutKind)
      : undefined;
  const attPreviewRaw = o.attachmentFilePreview;
  const attachmentFilePreview: TFormAttachmentFilePreviewKind | undefined =
    typeof attPreviewRaw === 'string' && ATTACHMENT_FILE_PREVIEW_SET.has(attPreviewRaw)
      ? (attPreviewRaw as TFormAttachmentFilePreviewKind)
      : undefined;
  const actionLog = sanitizeActionLog(o.actionLog);
  const skRaw = o.attachmentStorageKind;
  let attachmentStorageKind: TFormAttachmentStorageKind | undefined =
    skRaw === 'documentLibrary' ? 'documentLibrary' : undefined;
  const attachmentLibraryRaw = sanitizeAttachmentLibrary(o.attachmentLibrary);
  let attachmentLibrary: IFormManagerAttachmentLibraryConfig | undefined;
  if (attachmentStorageKind === 'documentLibrary') {
    const lt = attachmentLibraryRaw?.libraryTitle?.trim() ?? '';
    const lk = attachmentLibraryRaw?.sourceListLookupFieldInternalName?.trim() ?? '';
    if (!lt || !lk) {
      attachmentStorageKind = undefined;
      attachmentLibrary = attachmentLibraryRaw;
    } else {
      attachmentLibrary = {
        libraryTitle: lt,
        sourceListLookupFieldInternalName: lk,
        ...(attachmentLibraryRaw?.folderTree?.length ? { folderTree: attachmentLibraryRaw.folderTree } : {}),
      };
    }
  } else {
    attachmentStorageKind = undefined;
    attachmentLibrary = attachmentLibraryRaw;
  }
  const historyEnabled = o.historyEnabled === true;
  const hkRaw = o.historyPresentationKind;
  const historyPresentationKind: TFormHistoryPresentationKind | undefined =
    typeof hkRaw === 'string' && HISTORY_PRESENTATION_SET.has(hkRaw)
      ? (hkRaw as TFormHistoryPresentationKind)
      : undefined;
  const hbkRaw = o.historyButtonKind;
  const historyButtonKind: TFormHistoryButtonKind | undefined =
    typeof hbkRaw === 'string' && HISTORY_BUTTON_KIND_SET.has(hbkRaw)
      ? (hbkRaw as TFormHistoryButtonKind)
      : undefined;
  const historyButtonLabel =
    typeof o.historyButtonLabel === 'string' ? o.historyButtonLabel.trim().slice(0, 120) : undefined;
  const historyButtonIcon =
    typeof o.historyButtonIcon === 'string' ? o.historyButtonIcon.trim().slice(0, 80) : undefined;
  const historyPanelSubtitle =
    typeof o.historyPanelSubtitle === 'string' ? o.historyPanelSubtitle.trim() : undefined;
  const historyGroupTitlesRaw = Array.isArray(o.historyGroupTitles)
    ? (o.historyGroupTitles as unknown[]).map((x) => String(x).trim()).filter(Boolean)
    : [];
  const historyGroupTitles =
    historyGroupTitlesRaw.length > 0 ? (historyGroupTitlesRaw as string[]) : undefined;
  const hlRaw = o.historyLayoutKind;
  const historyLayoutKind: TFormHistoryLayoutKind | undefined =
    typeof hlRaw === 'string' && HISTORY_LAYOUT_SET.has(hlRaw)
      ? (hlRaw as TFormHistoryLayoutKind)
      : undefined;
  const customButtonsAdjusted: IFormCustomButtonConfig[] = [];
  for (let i = 0; i < customButtons.length; i++) {
    const btn = customButtons[i];
    if (btn.operation === 'history' && !historyEnabled) {
      const { operation: _op, shortDescription: _sd, ...rest } = btn;
      customButtonsAdjusted.push({ ...rest, actions: rest.actions });
    } else {
      customButtonsAdjusted.push(btn);
    }
  }
  return {
    sections,
    fields,
    rules,
    ...(steps.length ? { steps } : {}),
    ...(managerColumnFields?.length ? { managerColumnFields } : {}),
    ...(dynamicHelp.length ? { dynamicHelp } : {}),
    ...(customButtonsAdjusted.length ? { customButtons: customButtonsAdjusted } : {}),
    ...(stepLayout ? { stepLayout } : {}),
    ...(stepNavButtons && stepNavButtons !== 'fluent' ? { stepNavButtons } : {}),
    ...(formDataLoadingKind && formDataLoadingKind !== 'spinner' ? { formDataLoadingKind } : {}),
    ...(defaultSubmitLoadingKind && defaultSubmitLoadingKind !== 'overlay'
      ? { defaultSubmitLoadingKind }
      : {}),
    ...(formRootWidthMode ? { formRootWidthMode } : {}),
    ...(formRootWidthPercent !== undefined && formRootWidthPercent !== 100
      ? { formRootWidthPercent }
      : {}),
    ...(formRootHorizontalAlign ? { formRootHorizontalAlign } : {}),
    ...(formRootPaddingPx !== undefined ? { formRootPaddingPx } : {}),
    ...(stepNavigation ? { stepNavigation } : {}),
    ...(attachmentUploadLayout && attachmentUploadLayout !== 'default' ? { attachmentUploadLayout } : {}),
    ...(attachmentFilePreview && attachmentFilePreview !== 'nameAndSize' ? { attachmentFilePreview } : {}),
    ...(attachmentStorageKind === 'documentLibrary' && attachmentLibrary
      ? { attachmentStorageKind, attachmentLibrary }
      : attachmentLibrary
      ? { attachmentLibrary }
      : {}),
    ...(actionLog ? { actionLog } : {}),
    ...(historyEnabled ? { historyEnabled: true } : {}),
    ...(historyPresentationKind && historyPresentationKind !== 'panel'
      ? { historyPresentationKind }
      : {}),
    ...(historyEnabled && historyButtonKind && historyButtonKind !== 'text' ? { historyButtonKind } : {}),
    ...(historyEnabled && historyButtonLabel && historyButtonLabel !== 'Histórico'
      ? { historyButtonLabel }
      : {}),
    ...(historyEnabled && historyButtonIcon && historyButtonIcon !== 'History' ? { historyButtonIcon } : {}),
    ...(historyEnabled && historyPanelSubtitle ? { historyPanelSubtitle } : {}),
    ...(historyEnabled && historyGroupTitles?.length ? { historyGroupTitles } : {}),
    ...(historyLayoutKind && historyLayoutKind !== 'list' ? { historyLayoutKind } : {}),
    ...((): { linkedChildForms?: IFormLinkedChildFormConfig[] } => {
      if (!('linkedChildForms' in o)) return {};
      const linkedRaw = Array.isArray(o.linkedChildForms) ? o.linkedChildForms : [];
      const linkedChildForms: IFormLinkedChildFormConfig[] = [];
      for (let i = 0; i < linkedRaw.length && i < MAX_LINKED_CHILD_FORMS; i++) {
        const lc = sanitizeLinkedChildFormConfig(linkedRaw[i]);
        if (lc) linkedChildForms.push(lc);
      }
      return { linkedChildForms };
    })(),
  };
}
