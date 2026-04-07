import type {
  IFormManagerConfig,
  IFormStepNavigationConfig,
  IFormFieldConfig,
  IFormSectionConfig,
  IFormStepConfig,
  IFormCustomButtonConfig,
  TFormButtonAction,
  TFormCustomButtonOperation,
  TFormManagerFormMode,
  TFormRule,
  TFormConditionNode,
  TFormConditionOp,
  TFormStepLayoutKind,
  TFormStepNavButtonsKind,
  TFormDataLoadingUiKind,
  TFormSubmitLoadingUiKind,
  TFormAttachmentUploadLayoutKind,
  TFormAttachmentFilePreviewKind,
  IFormCompareRef,
} from '../config/types/formManager';
import { FORM_OCULTOS_STEP_ID } from '../config/types/formManager';

const STEP_LAYOUT_SET = new Set<string>(['rail', 'segmented', 'timeline', 'cards']);
const STEP_NAV_BUTTONS_SET = new Set<string>(['fluent', 'pills', 'dots', 'icons', 'links']);
const BUTTON_OPERATION_SET = new Set<string>(['legacy', 'redirect', 'add', 'update', 'delete']);
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

function sanitizeCompareRef(raw: unknown): IFormCompareRef | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const r = raw as Record<string, unknown>;
  const kind = r.kind === 'field' || r.kind === 'token' || r.kind === 'literal' ? r.kind : 'literal';
  const value = typeof r.value === 'string' ? r.value : String(r.value ?? '');
  return { kind, value };
}

function sanitizeConditionNode(raw: unknown): TFormConditionNode | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const n = raw as Record<string, unknown>;
  if (n.kind === 'all' || n.kind === 'any') {
    const childrenRaw = Array.isArray(n.children) ? n.children : [];
    const children: TFormConditionNode[] = [];
    for (let i = 0; i < childrenRaw.length; i++) {
      const c = sanitizeConditionNode(childrenRaw[i]);
      if (c) children.push(c);
    }
    if (children.length === 0) return undefined;
    return { kind: n.kind, children };
  }
  const leafLike =
    n.kind === 'leaf' ||
    (typeof n.field === 'string' &&
      n.field.trim() &&
      typeof n.op === 'string' &&
      n.kind !== 'all' &&
      n.kind !== 'any');
  if (leafLike) {
    const field = typeof n.field === 'string' ? n.field.trim() : '';
    const opRaw = typeof n.op === 'string' ? n.op : 'eq';
    const ops = new Set<string>([
      'eq', 'ne', 'gt', 'ge', 'lt', 'le', 'contains', 'startsWith', 'endsWith',
      'isEmpty', 'isFilled', 'isTrue', 'isFalse',
    ]);
    const op: TFormConditionOp = ops.has(opRaw) ? (opRaw as TFormConditionOp) : 'eq';
    if (!field) return undefined;
    const compare = sanitizeCompareRef(n.compare);
    return { kind: 'leaf', field, op: op as never, ...(compare ? { compare } : {}) };
  }
  return undefined;
}

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
      return {
        ...base,
        action: 'attachmentRules',
        ...(typeof r.minCount === 'number' ? { minCount: r.minCount } : {}),
        ...(typeof r.maxCount === 'number' ? { maxCount: r.maxCount } : {}),
        ...(typeof r.maxBytesPerFile === 'number' ? { maxBytesPerFile: r.maxBytesPerFile } : {}),
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
  return {
    internalName,
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
    return kind === 'showFields'
      ? { kind: 'showFields', fields, ...(whenAct ? { when: whenAct } : {}) }
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
    if (!targetField || !sourceFields.length) return undefined;
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
  const actions: TFormButtonAction[] = opResolved === 'redirect' ? [] : actionsSan;
  return {
    id,
    label,
    appearance,
    behavior,
    ...(operation && operation !== 'legacy' ? { operation } : {}),
    ...(redirectUrlTemplate !== undefined && redirectUrlTemplate.trim() ? { redirectUrlTemplate } : {}),
    ...(deleteShowInView === false ? { deleteShowInView: false } : {}),
    ...(deleteShowInEdit === false ? { deleteShowInEdit: false } : {}),
    ...(modes?.length ? { modes } : {}),
    ...(enabled === false ? { enabled: false } : {}),
    ...(when ? { when } : {}),
    ...(groupTitles?.length ? { groupTitles } : {}),
    ...(submitLoadingKind ? { submitLoadingKind } : {}),
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

function sanitizeStep(raw: unknown): IFormStepConfig | undefined {
  if (!raw || typeof raw !== 'object') return undefined;
  const s = raw as Record<string, unknown>;
  const id = typeof s.id === 'string' ? s.id.trim() : '';
  const title = typeof s.title === 'string' ? s.title.trim() : '';
  const fieldNames = Array.isArray(s.fieldNames) ? (s.fieldNames as unknown[]).map((x) => String(x).trim()).filter(Boolean) : [];
  if (!id || !title) return undefined;
  return { id, title, fieldNames };
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
    sections.push({ id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', visible: true });
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
    steps.push({ id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] });
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
  const showDefaultFormButtons = o.showDefaultFormButtons === true;
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
  return {
    sections,
    fields,
    rules,
    ...(steps.length ? { steps } : {}),
    ...(managerColumnFields?.length ? { managerColumnFields } : {}),
    ...(dynamicHelp.length ? { dynamicHelp } : {}),
    ...(customButtons.length ? { customButtons } : {}),
    ...(stepLayout ? { stepLayout } : {}),
    ...(stepNavButtons && stepNavButtons !== 'fluent' ? { stepNavButtons } : {}),
    ...(formDataLoadingKind && formDataLoadingKind !== 'spinner' ? { formDataLoadingKind } : {}),
    ...(defaultSubmitLoadingKind && defaultSubmitLoadingKind !== 'overlay'
      ? { defaultSubmitLoadingKind }
      : {}),
    ...(showDefaultFormButtons ? { showDefaultFormButtons: true } : {}),
    ...(stepNavigation ? { stepNavigation } : {}),
    ...(attachmentUploadLayout && attachmentUploadLayout !== 'default' ? { attachmentUploadLayout } : {}),
    ...(attachmentFilePreview && attachmentFilePreview !== 'nameAndSize' ? { attachmentFilePreview } : {}),
  };
}
