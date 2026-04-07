import * as React from 'react';
import { useState, useEffect, useMemo, useCallback, useRef } from 'react';
import { flushSync } from 'react-dom';
import {
  Stack,
  Text,
  TextField,
  Toggle,
  Dropdown,
  IDropdownOption,
  DatePicker,
  PrimaryButton,
  DefaultButton,
  IconButton,
  MessageBar,
  MessageBarType,
  Label,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type {
  IFormManagerConfig,
  IFormFieldConfig,
  IFormCustomButtonConfig,
  TFormButtonAction,
  TFormCustomButtonOperation,
  TFormManagerFormMode,
  TFormSubmitKind,
  TFormRule,
  TFormSubmitLoadingUiKind,
} from '../../core/config/types/formManager';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_OCULTOS_STEP_ID,
  FORM_BUILTIN_HISTORY_BUTTON_ID,
} from '../../core/config/types/formManager';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import {
  buildFormDerivedState,
  collectFormValidationErrors,
  filterValidationErrorsToStepFields,
  pickRequiredStyleStepErrors,
  evaluateCondition,
  evaluateFormValueExpression,
  getDefaultValuesFromRules,
  shouldShowCustomButton,
  shouldShowBuiltinHistoryButton,
  areAllRequiredFieldsFilled,
  type IFormRuleRuntimeContext,
  type IFormValidationAttachmentContext,
} from '../../core/formManager/formRuleEngine';
import { formValuesToSharePointPayload } from '../../core/formManager/formSharePointValues';
import { FormStepNavigation, FormStepPrevNextNav } from './FormStepLayoutUi';
import { FormAttachmentUploader } from './FormAttachmentUploader';
import { runAsyncFormValidations } from '../../core/formManager/formAsyncValidation';
import { interpolateFormButtonRedirectUrl } from '../../core/formManager/formButtonRedirectUrl';
import { appendFormActionLogEntry } from '../../core/formManager/formActionLog';
import { parseAttachmentUiRule } from '../../core/formManager/formManagerVisualModel';
import { ItemsService } from '../../../../services';
import { getSP } from '../../../../services/core/sp';
import { FormSubmitLoadingChrome, resolveSubmitLoadingKind } from './FormLoadingUi';
import { FormItemHistoryUi } from './FormItemHistoryUi';

export interface IDynamicListFormProps {
  listTitle: string;
  formManager: IFormManagerConfig;
  fieldMetadata: IFieldMetadata[];
  formMode: TFormManagerFormMode;
  initialItem?: Record<string, unknown> | null;
  itemId?: number;
  dynamicContext: IDynamicContext;
  userGroupTitles: string[];
  currentUserId: number;
  onSubmit: (payload: Record<string, unknown>, submitKind: TFormSubmitKind, pendingFiles: File[]) => Promise<void>;
  onDismiss: () => void;
  /** Chamado após botão «Atualizar» personalizado gravar com sucesso (ex.: recarregar item). */
  onAfterItemUpdated?: () => void | Promise<void>;
}

async function uploadListItemAttachments(listTitle: string, itemId: number, files: File[]): Promise<void> {
  if (!files.length) return;
  const sp = getSP();
  const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(listTitle);
  const list = isGuid ? sp.web.lists.getById(listTitle) : sp.web.lists.getByTitle(listTitle);
  const item = list.items.getById(itemId) as unknown as {
    attachmentFiles: { add(name: string, content: ArrayBuffer): Promise<unknown> };
  };
  for (let i = 0; i < files.length; i++) {
    const buf = await files[i].arrayBuffer();
    await item.attachmentFiles.add(files[i].name, buf);
  }
}

function itemToFormValues(
  item: Record<string, unknown> | undefined,
  names: string[]
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (!item) return out;
  for (let i = 0; i < names.length; i++) {
    const n = names[i];
    out[n] = item[n];
  }
  return out;
}

function formatJoinedFieldValue(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v !== null && 'Title' in (v as object)) {
    return String((v as Record<string, unknown>).Title ?? '');
  }
  return String(v);
}

type IFormButtonFieldOverlay = {
  show: Set<string>;
  hide: Set<string>;
  showOnStepId?: Record<string, string>;
};

function reduceCustomButtonActions(
  actions: TFormButtonAction[],
  startValues: Record<string, unknown>,
  dynamicContext: IDynamicContext,
  baseOverlay: IFormButtonFieldOverlay
): { mergedValues: Record<string, unknown>; mergedOverlay: IFormButtonFieldOverlay } {
  let next = { ...startValues };
  const mergedOverlay: IFormButtonFieldOverlay = {
    show: cloneStringSet(baseOverlay.show),
    hide: cloneStringSet(baseOverlay.hide),
    ...(baseOverlay.showOnStepId && Object.keys(baseOverlay.showOnStepId).length > 0
      ? { showOnStepId: { ...baseOverlay.showOnStepId } }
      : {}),
  };
  for (let i = 0; i < actions.length; i++) {
    const a = actions[i];
    if (a.when && !evaluateCondition(a.when, next, dynamicContext)) {
      continue;
    }
    if (a.kind === 'setFieldValue') {
      const tpl = a.valueTemplate;
      const raw =
        tpl.trim().indexOf('str:') === 0 ? evaluateFormValueExpression(tpl, next) : tpl;
      next = { ...next, [a.field]: raw };
    } else if (a.kind === 'joinFields') {
      const tpl = (a.valueTemplate ?? '').trim();
      if (tpl.length > 0) {
        const rawTpl = a.valueTemplate ?? '';
        const interpolated = rawTpl.replace(/\{\{([^}]+)\}\}/g, (_, raw: string) => {
          const name = String(raw).trim();
          return formatJoinedFieldValue(next[name]);
        });
        next = { ...next, [a.targetField]: interpolated };
      } else {
        const parts = a.sourceFields.map((f) => formatJoinedFieldValue(next[f]));
        next = { ...next, [a.targetField]: parts.join(a.separator) };
      }
    } else if (a.kind === 'showFields') {
      const sid = typeof a.displayOnStepId === 'string' ? a.displayOnStepId.trim() : '';
      for (let j = 0; j < a.fields.length; j++) {
        const fn = a.fields[j];
        mergedOverlay.show.add(fn);
        if (sid) {
          if (!mergedOverlay.showOnStepId) mergedOverlay.showOnStepId = {};
          mergedOverlay.showOnStepId[fn] = sid;
        }
      }
    } else if (a.kind === 'hideFields') {
      for (let j = 0; j < a.fields.length; j++) {
        const fn = a.fields[j];
        mergedOverlay.hide.add(fn);
        if (mergedOverlay.showOnStepId && mergedOverlay.showOnStepId[fn]) {
          delete mergedOverlay.showOnStepId[fn];
        }
      }
    }
  }
  return { mergedValues: next, mergedOverlay };
}

function cloneStringSet(s: Set<string>): Set<string> {
  const n = new Set<string>();
  s.forEach((x) => {
    n.add(x);
  });
  return n;
}

const REQ_EMPTY_BORDER = '#a4262c';

function isValueEmptyForRequired(v: unknown, mappedType: string): boolean {
  if (mappedType === 'boolean') {
    return v === undefined || v === null;
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

function stylesTextFieldRequiredEmpty(active: boolean): { fieldGroup?: Record<string, string | number> } | undefined {
  if (!active) return undefined;
  return {
    fieldGroup: {
      borderColor: REQ_EMPTY_BORDER,
      borderWidth: 1,
      borderStyle: 'solid',
      borderRadius: 2,
    },
  };
}

function listRequiredEmptyErrorsInStep(
  stepFieldList: Set<string>,
  values: Record<string, unknown>,
  metaByName: Map<string, IFieldMetadata>,
  fieldVisible: (n: string) => boolean
): Record<string, string> {
  const out: Record<string, string> = {};
  stepFieldList.forEach((name) => {
    if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) return;
    if (!fieldVisible(name)) return;
    const m = metaByName.get(name);
    if (!m?.Required) return;
    if (!isValueEmptyForRequired(values[name], m.MappedType)) return;
    out[name] = 'Obrigatório.';
  });
  return out;
}

function lookupIdFromValue(v: unknown): number | undefined {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (typeof id === 'number') return id;
  }
  return undefined;
}

export const DynamicListForm: React.FC<IDynamicListFormProps> = ({
  listTitle,
  formManager,
  fieldMetadata,
  formMode,
  initialItem,
  itemId,
  dynamicContext,
  userGroupTitles,
  currentUserId,
  onSubmit,
  onDismiss,
  onAfterItemUpdated,
}) => {
  const fieldConfigs: IFormFieldConfig[] =
    formManager.fields.length > 0
      ? formManager.fields
      : fieldMetadata
          .filter((f) => !f.Hidden && !f.ReadOnlyField && f.InternalName !== 'Id')
          .map((f) => ({ internalName: f.InternalName, sectionId: FORM_OCULTOS_STEP_ID }));
  const names = useMemo(
    () => fieldConfigs.map((f) => f.internalName).filter((n) => n !== FORM_ATTACHMENTS_FIELD_INTERNAL),
    [fieldConfigs]
  );
  const ocultosNullFieldNames = useMemo(
    () =>
      fieldConfigs
        .filter(
          (f) => f.sectionId === FORM_OCULTOS_STEP_ID && f.internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL
        )
        .map((f) => f.internalName),
    [fieldConfigs]
  );
  const metaByName = useMemo(() => new Map(fieldMetadata.map((f) => [f.InternalName, f])), [fieldMetadata]);

  const attachmentAllowedExtensions = useMemo(
    () => parseAttachmentUiRule(formManager.rules ?? []).allowedFileExtensions ?? [],
    [formManager.rules]
  );

  const [values, setValues] = useState<Record<string, unknown>>(() =>
    itemToFormValues(initialItem ?? undefined, names)
  );
  const [submitUi, setSubmitUi] = useState<TFormSubmitLoadingUiKind | null>(null);
  const submitting = submitUi !== null;
  const [formError, setFormError] = useState<string | undefined>(undefined);
  const [localErrors, setLocalErrors] = useState<Record<string, string>>({});
  const [lookupOptions, setLookupOptions] = useState<Record<string, IDropdownOption[]>>({});
  const [pendingFiles, setPendingFiles] = useState<File[]>([]);
  const [attachmentCount, setAttachmentCount] = useState(0);
  const prevByTriggerRef = useRef<Record<string, unknown>>({});
  const [buttonOverlay, setButtonOverlay] = useState<IFormButtonFieldOverlay>(() => ({
    show: new Set<string>(),
    hide: new Set<string>(),
  }));

  const authorId = useMemo(() => {
    const a = initialItem?.AuthorId ?? initialItem?.Author;
    if (typeof a === 'number') return a;
    if (a && typeof a === 'object' && 'Id' in (a as object)) return (a as { Id: number }).Id;
    return undefined;
  }, [initialItem]);

  const itemsService = useMemo(() => new ItemsService(), []);

  const builtinHistoryButtonConfig = useMemo((): IFormCustomButtonConfig => {
    const label = (formManager.historyButtonLabel ?? 'Histórico').trim() || 'Histórico';
    const sub = formManager.historyPanelSubtitle?.trim();
    return {
      id: FORM_BUILTIN_HISTORY_BUTTON_ID,
      label,
      shortDescription: sub || undefined,
      appearance: 'default',
      behavior: 'actionsOnly',
      operation: 'history',
      actions: [],
    };
  }, [formManager.historyButtonLabel, formManager.historyPanelSubtitle]);

  useEffect(() => {
    setValues(itemToFormValues(initialItem ?? undefined, names));
    setButtonOverlay({ show: new Set<string>(), hide: new Set<string>() });
  }, [initialItem, names]);

  useEffect(() => {
    if (formMode !== 'create') return;
    setValues((prev) => {
      const merged = getDefaultValuesFromRules(formManager, prev, dynamicContext);
      return merged;
    });
  }, [formManager, formMode, dynamicContext]);

  const runtimeCtx = useCallback(
    (submitKind?: TFormSubmitKind): IFormRuleRuntimeContext => ({
      formMode,
      values,
      submitKind,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
    }),
    [formMode, values, userGroupTitles, currentUserId, authorId, dynamicContext]
  );

  const derived = useMemo(
    () =>
      buildFormDerivedState(formManager, fieldConfigs, runtimeCtx(), {
        show: buttonOverlay.show,
        hide: buttonOverlay.hide,
      }),
    [
      formManager,
      fieldConfigs,
      runtimeCtx,
      values,
      formMode,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      buttonOverlay,
    ]
  );

  const allRequiredFilled = useMemo(
    () =>
      areAllRequiredFieldsFilled(
        formManager,
        fieldConfigs,
        runtimeCtx(),
        metaByName,
        { show: buttonOverlay.show, hide: buttonOverlay.hide },
        {
          attachmentCount,
          pendingFiles: pendingFiles.map((f) => ({
            size: f.size,
            type: f.type || 'application/octet-stream',
            name: f.name,
          })),
        }
      ),
    [
      formManager,
      fieldConfigs,
      runtimeCtx,
      values,
      formMode,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      buttonOverlay,
      attachmentCount,
      pendingFiles,
      metaByName,
    ]
  );

  const clearRules = useMemo(
    () => formManager.rules.filter((r): r is Extract<TFormRule, { action: 'clearFields' }> => r.action === 'clearFields'),
    [formManager.rules]
  );

  useEffect(() => {
    for (let i = 0; i < clearRules.length; i++) {
      const rule = clearRules[i];
      if (!rule.triggerField) continue;
      const cur = values[rule.triggerField];
      const prev = prevByTriggerRef.current[rule.triggerField];
      if (prev !== undefined && prev !== cur) {
        setValues((v) => {
          const next = { ...v };
          for (let j = 0; j < rule.fields.length; j++) next[rule.fields[j]] = null;
          return next;
        });
      }
      prevByTriggerRef.current[rule.triggerField] = cur;
    }
  }, [values, clearRules]);

  const loadLookupOptions = useCallback(
    async (fieldName: string, odataFilter?: string): Promise<void> => {
      const m = metaByName.get(fieldName);
      if (!m?.LookupList) return;
      try {
        const filter = odataFilter?.trim() ? odataFilter : undefined;
        const rows = await itemsService.getItems<Record<string, unknown>>(m.LookupList, {
          select: ['Id', m.LookupField || 'Title'],
          filter,
          top: 200,
        });
        const lf = m.LookupField || 'Title';
        const opts: IDropdownOption[] = [
          { key: '', text: '—' },
          ...rows.map((row) => ({
            key: String(row.Id),
            text: String(row[lf] ?? row.Id),
          })),
        ];
        setLookupOptions((o) => ({ ...o, [fieldName]: opts }));
      } catch {
        setLookupOptions((o) => ({ ...o, [fieldName]: [] }));
      }
    },
    [itemsService, metaByName]
  );

  useEffect(() => {
    void (async (): Promise<void> => {
      for (let i = 0; i < fieldConfigs.length; i++) {
        const fn = fieldConfigs[i].internalName;
        const m = metaByName.get(fn);
        if (m?.MappedType === 'lookup') {
          const lf = derived.lookupFilters[fn];
          let filter: string | undefined;
          if (lf) {
            const pid = lookupIdFromValue(values[lf.parentField]);
            if (pid !== undefined) filter = lf.odataFilterTemplate.split('{parent}').join(String(pid));
          }
          await loadLookupOptions(fn, filter);
        }
      }
    })();
  }, [fieldConfigs, metaByName, derived.lookupFilters, values, loadLookupOptions]);

  useEffect(() => {
    if (formMode === 'create' || !itemId) {
      setAttachmentCount(0);
      return;
    }
    let cancelled = false;
    void (async (): Promise<void> => {
      try {
        const sp = getSP();
        const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(listTitle);
        const list = isGuid ? sp.web.lists.getById(listTitle) : sp.web.lists.getByTitle(listTitle);
        const item = list.items.getById(itemId) as unknown as { attachmentFiles(): Promise<unknown[]> };
        const files = await item.attachmentFiles();
        if (!cancelled) setAttachmentCount(Array.isArray(files) ? files.length : 0);
      } catch {
        if (!cancelled) setAttachmentCount(0);
      }
    })();
    return (): void => {
      cancelled = true;
    };
  }, [listTitle, itemId, formMode]);

  const updateField = (name: string, v: unknown): void => {
    setValues((prev) => ({ ...prev, [name]: v }));
  };

  const validate = async (
    submitKind: TFormSubmitKind,
    opts?: {
      values?: Record<string, unknown>;
      buttonOverlay?: IFormButtonFieldOverlay;
    }
  ): Promise<boolean> => {
    const vals = opts?.values ?? values;
    const ov = opts?.buttonOverlay ?? buttonOverlay;
    const att: IFormValidationAttachmentContext = {
      attachmentCount,
      pendingFiles: pendingFiles.map((f) => ({
        size: f.size,
        type: f.type || 'application/octet-stream',
        name: f.name,
      })),
    };
    const ctx: IFormRuleRuntimeContext = {
      formMode,
      values: vals,
      submitKind,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
    };
    const sync = collectFormValidationErrors(formManager, fieldConfigs, ctx, att, {
      show: ov.show,
      hide: ov.hide,
    });
    setLocalErrors(sync);
    if (Object.keys(sync).length > 0) return false;
    const asyncErr = await runAsyncFormValidations(formManager, vals, itemsService, listTitle, itemId, submitKind);
    if (Object.keys(asyncErr).length > 0) {
      setLocalErrors(asyncErr);
      return false;
    }
    return true;
  };

  const handleSave = async (
    submitKind: TFormSubmitKind,
    opts?: {
      valuesOverride?: Record<string, unknown>;
      buttonOverlayOverride?: IFormButtonFieldOverlay;
      submitLoadingFromButton?: IFormCustomButtonConfig;
    }
  ): Promise<boolean> => {
    const vals = opts?.valuesOverride ?? values;
    const ov = opts?.buttonOverlayOverride ?? buttonOverlay;
    setFormError(undefined);
    const ok = await validate(submitKind, { values: vals, buttonOverlay: ov });
    if (!ok) return false;
    setSubmitUi(resolveSubmitLoadingKind(formManager, opts?.submitLoadingFromButton));
    try {
      const payload = formValuesToSharePointPayload(fieldMetadata, vals, names, {
        nullWhenEmptyFieldNames: ocultosNullFieldNames,
      });
      await onSubmit(payload, submitKind, pendingFiles);
      return true;
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
      return false;
    } finally {
      setSubmitUi(null);
    }
  };

  const stepsAll = formManager.steps?.length ? formManager.steps : null;
  const visibleStepsForUi = useMemo(() => {
    if (!stepsAll) return null;
    return stepsAll.filter((s) => s.id !== FORM_OCULTOS_STEP_ID);
  }, [stepsAll]);
  const [stepIndex, setStepIndex] = useState(0);
  const [historyBtn, setHistoryBtn] = useState<IFormCustomButtonConfig | null>(null);
  useEffect(() => {
    if (!visibleStepsForUi?.length) return;
    setStepIndex((i) => Math.min(i, visibleStepsForUi.length - 1));
  }, [visibleStepsForUi]);

  const runCustomButton = async (btn: IFormCustomButtonConfig): Promise<void> => {
    const op: TFormCustomButtonOperation = btn.operation ?? 'legacy';
    if (op === 'history') {
      if (formManager.historyEnabled !== true) {
        setFormError('Ative o histórico na aba Componentes do gestor de formulário.');
        return;
      }
      if (itemId === undefined || itemId === null || formMode === 'create') {
        setFormError('O histórico só está disponível quando o item já existe na lista.');
        return;
      }
      setFormError(undefined);
      try {
        await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
          sourceListTitle: listTitle,
          sourceItemId: itemId,
          formMode,
        });
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
        return;
      }
      setHistoryBtn(btn);
      return;
    }
    const actions = op === 'redirect' ? [] : btn.actions ?? [];
    const { mergedValues, mergedOverlay } = reduceCustomButtonActions(
      actions,
      values,
      dynamicContext,
      buttonOverlay
    );
    if (op !== 'redirect') {
      flushSync(() => {
        setValues(mergedValues);
        setButtonOverlay(mergedOverlay);
      });
      const behaviorEarly = btn.behavior ?? 'actionsOnly';
      if (
        behaviorEarly === 'actionsOnly' &&
        mergedOverlay.show.size > 0 &&
        visibleStepsForUi &&
        visibleStepsForUi.length > 1
      ) {
        const shownNames = Array.from(mergedOverlay.show);
        for (let si = 0; si < shownNames.length; si++) {
          const name = shownNames[si];
          const stepId = mergedOverlay.showOnStepId?.[name];
          if (!stepId) continue;
          const idx = visibleStepsForUi.findIndex((s) => s.id === stepId);
          if (idx >= 0) {
            setStepIndex(idx);
            break;
          }
        }
      }
    }

    if (op === 'redirect') {
      const tpl = (btn.redirectUrlTemplate ?? '').trim();
      if (!tpl) {
        setFormError('Configure o URL de redirecionamento no gestor de formulário.');
        return;
      }
      const url = interpolateFormButtonRedirectUrl(tpl, mergedValues, { itemId, formMode });
      try {
        await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
          sourceListTitle: listTitle,
          sourceItemId: itemId,
          formMode,
        });
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
        return;
      }
      window.location.assign(url);
      return;
    }

    if (op === 'add') {
      setFormError(undefined);
      const ok = await validate('submit', { values: mergedValues, buttonOverlay: mergedOverlay });
      if (!ok) return;
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      try {
        const payload = formValuesToSharePointPayload(fieldMetadata, mergedValues, names, {
          nullWhenEmptyFieldNames: ocultosNullFieldNames,
        });
        const { id: newId, filesForAttachments } = await itemsService.addItem(
          listTitle,
          payload,
          pendingFiles
        );
        await uploadListItemAttachments(listTitle, newId, filesForAttachments);
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: newId,
            formMode,
          });
        } catch (le) {
          setFormError(
            `Item criado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
          );
        }
        onDismiss();
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      return;
    }

    if (op === 'update') {
      if (!itemId || formMode === 'create') {
        setFormError('Atualizar requer um item aberto (parâmetros Form / FormID na página).');
        return;
      }
      setFormError(undefined);
      const ok = await validate('submit', { values: mergedValues, buttonOverlay: mergedOverlay });
      if (!ok) return;
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      try {
        const payload = formValuesToSharePointPayload(fieldMetadata, mergedValues, names, {
          nullWhenEmptyFieldNames: ocultosNullFieldNames,
        });
        await itemsService.updateItem(listTitle, itemId, payload);
        await uploadListItemAttachments(listTitle, itemId, pendingFiles);
        await onAfterItemUpdated?.();
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: itemId,
            formMode,
          });
        } catch (le) {
          setFormError(
            `Gravado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
          );
        }
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      return;
    }

    if (op === 'delete') {
      if (!itemId || formMode === 'create') {
        setFormError('Eliminar só está disponível ao editar ou ver um item existente.');
        return;
      }
      if (!window.confirm('Eliminar este item permanentemente?')) return;
      setFormError(undefined);
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      try {
        await itemsService.deleteItem(listTitle, itemId);
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: itemId,
            formMode,
          });
        } catch (le) {
          setFormError(
            `Eliminado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
          );
        }
        onDismiss();
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      return;
    }

    const behavior = btn.behavior ?? 'actionsOnly';
    if (behavior === 'close') {
      try {
        await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
          sourceListTitle: listTitle,
          sourceItemId: itemId,
          formMode,
        });
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
        return;
      }
      onDismiss();
      return;
    }
    if (behavior === 'draft') {
      const saved = await handleSave('draft', {
        valuesOverride: mergedValues,
        buttonOverlayOverride: mergedOverlay,
        submitLoadingFromButton: btn,
      });
      if (saved) {
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: itemId,
            formMode,
          });
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
        }
      }
    } else if (behavior === 'submit') {
      const saved = await handleSave('submit', {
        valuesOverride: mergedValues,
        buttonOverlayOverride: mergedOverlay,
        submitLoadingFromButton: btn,
      });
      if (saved) {
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: itemId,
            formMode,
          });
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
        }
      }
    } else if (behavior === 'actionsOnly') {
      try {
        await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
          sourceListTitle: listTitle,
          sourceItemId: itemId,
          formMode,
        });
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
      }
    }
  };

  const currentStepFieldSet = useMemo(() => {
    if (!visibleStepsForUi?.length) return null;
    const s = visibleStepsForUi[stepIndex];
    if (!s) return null;
    return new Set(s.fieldNames);
  }, [visibleStepsForUi, stepIndex]);

  const tryGoToStep = useCallback(
    (targetIndex: number) => {
      if (!visibleStepsForUi?.length) return;
      const max = visibleStepsForUi.length - 1;
      const t = Math.max(0, Math.min(max, targetIndex));
      if (t === stepIndex) return;
      if (formMode === 'view') {
        setStepIndex(t);
        return;
      }
      const sn = formManager.stepNavigation;
      const requireBlock =
        sn?.requireFilledRequiredToAdvance === true ||
        (sn as { requireFilledRequiredToAdvance?: unknown } | undefined)?.requireFilledRequiredToAdvance ===
          'true';
      const fullVal = sn?.fullValidationOnAdvance === true;
      const allowBackFree = sn?.allowBackWithoutValidation !== false;
      const attCtx: IFormValidationAttachmentContext = {
        attachmentCount,
        pendingFiles: pendingFiles.map((f) => ({
          size: f.size,
          type: f.type || 'application/octet-stream',
          name: f.name,
        })),
      };
      const ctx: IFormRuleRuntimeContext = {
        formMode,
        values,
        submitKind: 'submit',
        userGroupTitles,
        currentUserId,
        authorId,
        dynamicContext,
      };
      const overlay = { show: buttonOverlay.show, hide: buttonOverlay.hide };
      const syncErrorsForStep = (stepFieldList: Set<string>): Record<string, string> | null => {
        const derivedStep = buildFormDerivedState(formManager, fieldConfigs, ctx, overlay);
        const fv = (n: string): boolean => derivedStep.fieldVisible[n] !== false;
        const sync = collectFormValidationErrors(formManager, fieldConfigs, ctx, attCtx, overlay);
        let rel = filterValidationErrorsToStepFields(sync, stepFieldList);
        if (!fullVal) rel = pickRequiredStyleStepErrors(rel);
        const listReq = listRequiredEmptyErrorsInStep(stepFieldList, values, metaByName, fv);
        const blocks = Object.keys(rel).length > 0 || Object.keys(listReq).length > 0;
        if (!blocks) return null;
        return { ...sync, ...listReq };
      };
      if (!requireBlock) {
        setFormError(undefined);
        setStepIndex(t);
        return;
      }
      if (t < stepIndex) {
        if (allowBackFree) {
          setFormError(undefined);
          setStepIndex(t);
          return;
        }
        const cur = visibleStepsForUi[stepIndex];
        const curSet = new Set(cur?.fieldNames ?? []);
        const bad = syncErrorsForStep(curSet);
        if (bad) {
          setLocalErrors(bad);
          setFormError('Complete esta etapa antes de mudar.');
          return;
        }
        setFormError(undefined);
        setStepIndex(t);
        return;
      }
      for (let k = stepIndex; k < t; k++) {
        const st = visibleStepsForUi[k];
        const fieldSet = new Set(st?.fieldNames ?? []);
        const bad = syncErrorsForStep(fieldSet);
        if (bad) {
          setStepIndex(k);
          setLocalErrors(bad);
          setFormError('Complete esta etapa antes de continuar.');
          return;
        }
      }
      setFormError(undefined);
      setStepIndex(t);
    },
    [
      visibleStepsForUi,
      stepIndex,
      formMode,
      formManager,
      fieldConfigs,
      values,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      attachmentCount,
      pendingFiles,
      buttonOverlay.show,
      buttonOverlay.hide,
      buttonOverlay.showOnStepId,
      metaByName,
    ]
  );

  const [modalOpen, setModalOpen] = useState(false);
  const modalGroupIds = useMemo(() => {
    const seen: Record<string, boolean> = {};
    const ids: string[] = [];
    for (let i = 0; i < fieldConfigs.length; i++) {
      const gid = fieldConfigs[i].modalGroupId;
      if (gid && !seen[gid]) {
        seen[gid] = true;
        ids.push(gid);
      }
    }
    return ids;
  }, [fieldConfigs]);

  const renderFieldControl = (fc: IFormFieldConfig): React.ReactNode => {
    const name = fc.internalName;
    if (derived.fieldVisible[name] === false) return null;
    if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) {
      const readOnly = formMode === 'view' || derived.fieldDisabled[name] === true;
      const attErr = localErrors._attachments;
      const attReq = derived.fieldRequired[name] === true;
      const attachmentSatisfied =
        pendingFiles.length > 0 || (formMode !== 'create' && attachmentCount > 0);
      const attReqEmpty = attReq && !readOnly && !attachmentSatisfied;
      if (formMode === 'view') {
        return (
          <Stack key={name} tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 12 } }}>
            <Label required={attReq}>{fc.label ?? 'Anexos ao item'}</Label>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {attachmentCount} anexo(s) no item. Não é possível adicionar novos em modo ver.
            </Text>
            {fc.helpText && (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {fc.helpText}
              </Text>
            )}
          </Stack>
        );
      }
      return (
        <FormAttachmentUploader
          key={name}
          files={pendingFiles}
          onFilesChange={setPendingFiles}
          disabled={readOnly}
          label={fc.label ?? 'Anexos ao item'}
          description={fc.helpText}
          errorMessage={attErr}
          required={attReq}
          requiredEmptyHighlight={attReqEmpty}
          layout={formManager.attachmentUploadLayout ?? 'default'}
          filePreview={formManager.attachmentFilePreview ?? 'nameAndSize'}
          allowedFileExtensions={
            attachmentAllowedExtensions.length > 0 ? attachmentAllowedExtensions : undefined
          }
        />
      );
    }
    const m = metaByName.get(name);
    if (!m) return null;
    const disabled = formMode === 'view' || derived.fieldDisabled[name] === true;
    const readOnly = derived.fieldReadOnly[name] === true || disabled;
    const err = localErrors[name];
    const label = fc.label ?? m.Title;
    const help = derived.dynamicHelpByField[name] ?? fc.helpText;
    const isRequired = derived.fieldRequired[name] === true || m.Required === true;
    const canFill = formMode !== 'view' && !readOnly;
    const showReqEmpty = isRequired && canFill && isValueEmptyForRequired(values[name], m.MappedType);
    const comp = derived.computedDisplay[name];
    if (comp !== undefined) {
      return (
        <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 12 } }}>
          <Label required={isRequired}>{label}</Label>
          <Text styles={{ root: { color: '#323130' } }}>{String(comp)}</Text>
          {help && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
        </Stack>
      );
    }

    const common = { disabled: readOnly, errorMessage: err };

    switch (m.MappedType) {
      case 'boolean':
        return (
          <Stack
            key={name}
            tokens={{ childrenGap: 6 }}
            styles={{
              root: {
                marginBottom: 12,
                ...(showReqEmpty
                  ? { borderLeft: `3px solid ${REQ_EMPTY_BORDER}`, paddingLeft: 8, paddingTop: 2, paddingBottom: 2 }
                  : {}),
              },
            }}
          >
            <Label required={isRequired}>{label}</Label>
            <Toggle
              ariaLabel={label}
              onText="Sim"
              offText="Não"
              checked={values[name] === true || values[name] === 1}
              onChange={(_, c) => updateField(name, !!c)}
              disabled={readOnly}
            />
          </Stack>
        );
      case 'number':
      case 'currency':
        return (
          <TextField
            key={name}
            label={label}
            type="number"
            placeholder={fc.placeholder}
            value={values[name] !== null && values[name] !== undefined ? String(values[name]) : ''}
            onChange={(_, v) => updateField(name, v === '' ? null : Number(v))}
            required={isRequired}
            {...common}
            description={help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
          />
        );
      case 'datetime':
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 8 } }}>
            <Label required={isRequired}>{label}</Label>
            <DatePicker
              value={values[name] ? new Date(String(values[name])) : undefined}
              onSelectDate={(d) => updateField(name, d ? d.toISOString() : null)}
              disabled={readOnly}
              textField={{
                errorMessage: err,
                styles: stylesTextFieldRequiredEmpty(showReqEmpty),
              }}
            />
            {help && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
          </Stack>
        );
      case 'choice': {
        const opts: IDropdownOption[] = (m.Choices ?? []).map((c) => ({ key: c, text: c }));
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            options={opts}
            selectedKey={values[name] !== undefined && values[name] !== null ? String(values[name]) : ''}
            onChange={(_, o) => o && updateField(name, o.key)}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={
              showReqEmpty
                ? {
                    dropdown: {
                      borderColor: REQ_EMPTY_BORDER,
                      borderWidth: 1,
                      borderStyle: 'solid',
                      borderRadius: 2,
                    },
                  }
                : undefined
            }
          />
        );
      }
      case 'multichoice': {
        const selected: string[] = Array.isArray(values[name])
          ? (values[name] as string[])
          : typeof values[name] === 'string'
            ? String(values[name]).split(';').map((s) => s.trim()).filter(Boolean)
            : [];
        const opts: IDropdownOption[] = (m.Choices ?? []).map((c) => ({
          key: c,
          text: c,
          selected: selected.indexOf(c) !== -1,
        }));
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            multiSelect
            options={opts}
            selectedKeys={selected}
            onChange={(_, o) => {
              if (!o) return;
              const k = String(o.key);
              const next = selected.indexOf(k) !== -1 ? selected.filter((x) => x !== k) : [...selected, k];
              updateField(name, next);
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={
              showReqEmpty
                ? {
                    dropdown: {
                      borderColor: REQ_EMPTY_BORDER,
                      borderWidth: 1,
                      borderStyle: 'solid',
                      borderRadius: 2,
                    },
                  }
                : undefined
            }
          />
        );
      }
      case 'lookup': {
        const opts = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const id = lookupIdFromValue(values[name]);
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            options={opts}
            selectedKey={id !== undefined ? String(id) : ''}
            onChange={(_, o) => {
              if (!o || o.key === '') updateField(name, null);
              else updateField(name, { Id: Number(o.key), Title: o.text });
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={
              showReqEmpty
                ? {
                    dropdown: {
                      borderColor: REQ_EMPTY_BORDER,
                      borderWidth: 1,
                      borderStyle: 'solid',
                      borderRadius: 2,
                    },
                  }
                : undefined
            }
          />
        );
      }
      case 'user': {
        const id = lookupIdFromValue(values[name]);
        return (
          <TextField
            key={name}
            label={`${label} (Id)`}
            placeholder={fc.placeholder}
            value={id !== undefined ? String(id) : ''}
            onChange={(_, v) => updateField(name, v === '' ? null : { Id: Number(v), Title: '' })}
            required={isRequired}
            {...common}
            description={help ?? 'Informe o ID numérico do usuário no site.'}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
          />
        );
      }
      case 'multiline':
        return (
          <TextField
            key={name}
            label={label}
            multiline
            rows={4}
            placeholder={fc.placeholder}
            value={values[name] !== null && values[name] !== undefined ? String(values[name]) : ''}
            onChange={(_, v) => updateField(name, v ?? '')}
            required={isRequired}
            {...common}
            description={help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
          />
        );
      default:
        return (
          <TextField
            key={name}
            label={label}
            placeholder={fc.placeholder}
            value={values[name] !== null && values[name] !== undefined ? String(values[name]) : ''}
            onChange={(_, v) => updateField(name, v ?? '')}
            required={isRequired}
            {...common}
            description={help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
          />
        );
    }
  };

  const renderFields = (scope: 'main' | 'modal'): React.ReactNode => {
    const bySection = new Map<string, IFormFieldConfig[]>();
    for (let i = 0; i < fieldConfigs.length; i++) {
      const fc = fieldConfigs[i];
      const inModal = !!fc.modalGroupId;
      if (scope === 'modal' && !inModal) continue;
      if (scope === 'main' && inModal) continue;
      if (scope === 'main' && currentStepFieldSet) {
        const name = fc.internalName;
        if (!currentStepFieldSet.has(name)) {
          if (!buttonOverlay.show.has(name)) continue;
          const vis = visibleStepsForUi;
          if (!vis?.length) continue;
          const sid = buttonOverlay.showOnStepId?.[name];
          const fallbackSingle = vis.length === 1 ? vis[0].id : undefined;
          const target = sid || fallbackSingle;
          const curId = vis[stepIndex]?.id;
          if (!target || !curId || target !== curId) continue;
        }
      }
      let sid =
        derived.effectiveSectionByField[fc.internalName] ?? fc.sectionId ?? formManager.sections[0]?.id ?? 'main';
      if (
        scope === 'main' &&
        fc.sectionId === FORM_OCULTOS_STEP_ID &&
        buttonOverlay.show.has(fc.internalName) &&
        visibleStepsForUi?.length
      ) {
        const mapSid = buttonOverlay.showOnStepId?.[fc.internalName];
        const fallbackSingle = visibleStepsForUi.length === 1 ? visibleStepsForUi[0].id : undefined;
        const stepTarget = mapSid || fallbackSingle;
        if (stepTarget && formManager.sections.some((x) => x.id === stepTarget)) {
          sid = stepTarget;
        }
      }
      const arr = bySection.get(sid) ?? [];
      arr.push(fc);
      bySection.set(sid, arr);
    }
    const out: React.ReactNode[] = [];
    for (let s = 0; s < formManager.sections.length; s++) {
      const sec = formManager.sections[s];
      if (sec.id === FORM_OCULTOS_STEP_ID) continue;
      if (derived.sectionVisible[sec.id] === false) continue;
      const fields = bySection.get(sec.id);
      if (!fields?.length) continue;
      out.push(
        <Stack key={sec.id} tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 16 } }}>
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>{sec.title}</Text>
          {fields.map((fc) => renderFieldControl(fc))}
        </Stack>
      );
    }
    return <>{out}</>;
  };

  const submitMsg = 'A gravar…';

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 8 } }}>
      <FormSubmitLoadingChrome kind="infoBar" active={submitUi === 'infoBar'} message={submitMsg} />
      <FormSubmitLoadingChrome kind="topProgress" active={submitUi === 'topProgress'} message={submitMsg} />
      <Stack
        styles={{
          root: {
            position: 'relative',
            minHeight: submitUi === 'overlay' || submitUi === 'formShimmer' ? 160 : undefined,
          },
        }}
        tokens={{ childrenGap: 16 }}
      >
        {formError && <MessageBar messageBarType={MessageBarType.error}>{formError}</MessageBar>}
        {localErrors._attachments && (
          <MessageBar messageBarType={MessageBarType.error}>{localErrors._attachments}</MessageBar>
        )}
        {localErrors._async && <MessageBar messageBarType={MessageBarType.error}>{localErrors._async}</MessageBar>}
        {derived.messages.map((m) => (
          <MessageBar
            key={m.ruleId}
            messageBarType={
              m.variant === 'error'
                ? MessageBarType.error
                : m.variant === 'warning'
                  ? MessageBarType.warning
                  : MessageBarType.info
            }
          >
            {m.text}
          </MessageBar>
        ))}
        {visibleStepsForUi && visibleStepsForUi.length > 1 && (
          <FormStepNavigation
            steps={visibleStepsForUi}
            activeIndex={stepIndex}
            onStepSelect={(i) => tryGoToStep(i)}
            layout={formManager.stepLayout ?? 'segmented'}
          />
        )}
        {modalGroupIds.length > 0 && formMode !== 'view' && (
          <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
            {modalGroupIds.map((gid: string) => (
              <DefaultButton key={gid} text={`Editar ${gid}`} onClick={() => setModalOpen(true)} />
            ))}
          </Stack>
        )}
        {renderFields('main')}
        {visibleStepsForUi && visibleStepsForUi.length > 1 && (
          <Stack
            horizontal
            verticalAlign="center"
            tokens={{ childrenGap: 8 }}
            styles={{ root: { width: '100%' } }}
            wrap
          >
            <FormStepPrevNextNav
              variant={formManager.stepNavButtons ?? 'fluent'}
              stepIndex={stepIndex}
              stepCount={visibleStepsForUi.length}
              onPrev={() => tryGoToStep(stepIndex - 1)}
              onNext={() => tryGoToStep(stepIndex + 1)}
              disabled={submitting}
            />
          </Stack>
        )}
        <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
          {formManager.historyEnabled === true &&
            shouldShowBuiltinHistoryButton({
              historyEnabledInConfig: true,
              historyItemId: itemId,
              historyGroupTitles: formManager.historyGroupTitles,
              userGroupTitles,
            }) &&
            (() => {
              const hk = formManager.historyButtonKind ?? 'text';
              const hLabel = (formManager.historyButtonLabel ?? 'Histórico').trim() || 'Histórico';
              const hIcon = (formManager.historyButtonIcon ?? 'History').trim() || 'History';
              const hSub = formManager.historyPanelSubtitle?.trim();
              const onH = (): void => {
                void runCustomButton(builtinHistoryButtonConfig);
              };
              if (hk === 'icon') {
                return (
                  <IconButton
                    key="__builtin_history"
                    iconProps={{ iconName: hIcon }}
                    title={hSub || hLabel}
                    ariaLabel={hLabel}
                    onClick={onH}
                    disabled={submitting}
                  />
                );
              }
              if (hk === 'iconAndText') {
                return (
                  <DefaultButton
                    key="__builtin_history"
                    text={hLabel}
                    iconProps={{ iconName: hIcon }}
                    title={hSub}
                    onClick={onH}
                    disabled={submitting}
                  />
                );
              }
              return (
                <DefaultButton
                  key="__builtin_history"
                  text={hLabel}
                  title={hSub}
                  onClick={onH}
                  disabled={submitting}
                />
              );
            })()}
          {(formManager.customButtons ?? [])
            .filter((b) =>
              formManager.historyEnabled === true ? (b.operation ?? 'legacy') !== 'history' : true
            )
            .filter((b) =>
              shouldShowCustomButton(b, runtimeCtx(), {
                allRequiredFilled: allRequiredFilled,
                historyEnabledInConfig: formManager.historyEnabled === true,
                historyItemId: itemId,
              })
            )
            .map((b) =>
              b.appearance === 'primary' ? (
                <PrimaryButton
                  key={b.id}
                  text={b.label}
                  title={b.shortDescription || undefined}
                  onClick={() => void runCustomButton(b)}
                  disabled={submitting}
                />
              ) : (
                <DefaultButton
                  key={b.id}
                  text={b.label}
                  title={b.shortDescription || undefined}
                  onClick={() => void runCustomButton(b)}
                  disabled={submitting}
                />
              )
            )}
          {formManager.showDefaultFormButtons === true && formMode !== 'view' && (
            <>
              <PrimaryButton text="Enviar" onClick={() => handleSave('submit')} disabled={submitting} />
              <DefaultButton text="Rascunho" onClick={() => handleSave('draft')} disabled={submitting} />
            </>
          )}
          {formManager.showDefaultFormButtons === true && (
            <DefaultButton text="Fechar" onClick={onDismiss} disabled={submitting} />
          )}
        </Stack>
        {historyBtn &&
          itemId !== undefined &&
          itemId !== null &&
          formManager.historyEnabled === true && (
            <FormItemHistoryUi
              listTitle={listTitle}
              itemId={itemId}
              presentationKind={formManager.historyPresentationKind ?? 'panel'}
              isOpen={true}
              onDismiss={() => setHistoryBtn(null)}
              title={historyBtn.label}
              subtitle={historyBtn.shortDescription}
            />
          )}
        <FormSubmitLoadingChrome kind="belowButtons" active={submitUi === 'belowButtons'} message={submitMsg} />
        <FormSubmitLoadingChrome kind="overlay" active={submitUi === 'overlay'} message={submitMsg} />
        <FormSubmitLoadingChrome kind="formShimmer" active={submitUi === 'formShimmer'} message={submitMsg} />
        {modalOpen && (
          <Stack styles={{ root: { borderTop: '1px solid #edebe9', paddingTop: 16 } }} tokens={{ childrenGap: 12 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Campos adicionais</Text>
            {renderFields('modal')}
            <DefaultButton text="Fechar modal" onClick={() => setModalOpen(false)} />
          </Stack>
        )}
      </Stack>
    </Stack>
  );
};
