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
  TFormManagerFormMode,
  TFormSubmitKind,
  TFormRule,
} from '../../core/config/types/formManager';
import { FORM_ATTACHMENTS_FIELD_INTERNAL } from '../../core/config/types/formManager';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import {
  buildFormDerivedState,
  collectFormValidationErrors,
  evaluateFormValueExpression,
  getDefaultValuesFromRules,
  type IFormRuleRuntimeContext,
  type IFormValidationAttachmentContext,
} from '../../core/formManager/formRuleEngine';
import { formValuesToSharePointPayload } from '../../core/formManager/formSharePointValues';
import { FormStepNavigation, FormStepPrevNextNav } from './FormStepLayoutUi';
import { FormAttachmentUploader } from './FormAttachmentUploader';
import { runAsyncFormValidations } from '../../core/formManager/formAsyncValidation';
import { ItemsService } from '../../../../services';
import { getSP } from '../../../../services/core/sp';

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

function applyButtonActionsToValues(
  actions: TFormButtonAction[],
  prev: Record<string, unknown>
): Record<string, unknown> {
  let next = { ...prev };
  for (let i = 0; i < actions.length; i++) {
    const a = actions[i];
    if (a.kind === 'setFieldValue') {
      const tpl = a.valueTemplate;
      const raw =
        tpl.trim().indexOf('str:') === 0 ? evaluateFormValueExpression(tpl, next) : tpl;
      next = { ...next, [a.field]: raw };
    } else if (a.kind === 'joinFields') {
      const parts = a.sourceFields.map((f) => formatJoinedFieldValue(next[f]));
      next = { ...next, [a.targetField]: parts.join(a.separator) };
    }
  }
  return next;
}

function collectButtonOverlayDeltas(actions: TFormButtonAction[]): { show: string[]; hide: string[] } {
  const show: string[] = [];
  const hide: string[] = [];
  for (let i = 0; i < actions.length; i++) {
    const a = actions[i];
    if (a.kind === 'showFields') show.push(...a.fields);
    else if (a.kind === 'hideFields') hide.push(...a.fields);
  }
  return { show, hide };
}

function cloneStringSet(s: Set<string>): Set<string> {
  const n = new Set<string>();
  s.forEach((x) => {
    n.add(x);
  });
  return n;
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
}) => {
  const fieldConfigs: IFormFieldConfig[] =
    formManager.fields.length > 0
      ? formManager.fields
      : fieldMetadata
          .filter((f) => !f.Hidden && !f.ReadOnlyField && f.InternalName !== 'Id')
          .map((f) => ({ internalName: f.InternalName, sectionId: formManager.sections[0]?.id ?? 'main' }));
  const names = useMemo(
    () => fieldConfigs.map((f) => f.internalName).filter((n) => n !== FORM_ATTACHMENTS_FIELD_INTERNAL),
    [fieldConfigs]
  );
  const metaByName = useMemo(() => new Map(fieldMetadata.map((f) => [f.InternalName, f])), [fieldMetadata]);

  const [values, setValues] = useState<Record<string, unknown>>(() =>
    itemToFormValues(initialItem ?? undefined, names)
  );
  const [submitting, setSubmitting] = useState(false);
  const [formError, setFormError] = useState<string | undefined>(undefined);
  const [localErrors, setLocalErrors] = useState<Record<string, string>>({});
  const [lookupOptions, setLookupOptions] = useState<Record<string, IDropdownOption[]>>({});
  const [pendingFiles, setPendingFiles] = useState<File[]>([]);
  const [attachmentCount, setAttachmentCount] = useState(0);
  const prevByTriggerRef = useRef<Record<string, unknown>>({});
  const [buttonOverlay, setButtonOverlay] = useState<{ show: Set<string>; hide: Set<string> }>(() => ({
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
      buttonOverlay?: { show: Set<string>; hide: Set<string> };
    }
  ): Promise<boolean> => {
    const vals = opts?.values ?? values;
    const ov = opts?.buttonOverlay ?? buttonOverlay;
    const att: IFormValidationAttachmentContext = {
      attachmentCount,
      pendingFiles: pendingFiles.map((f) => ({ size: f.size, type: f.type || 'application/octet-stream' })),
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
      buttonOverlayOverride?: { show: Set<string>; hide: Set<string> };
    }
  ): Promise<void> => {
    const vals = opts?.valuesOverride ?? values;
    const ov = opts?.buttonOverlayOverride ?? buttonOverlay;
    setFormError(undefined);
    const ok = await validate(submitKind, { values: vals, buttonOverlay: ov });
    if (!ok) return;
    setSubmitting(true);
    try {
      const payload = formValuesToSharePointPayload(fieldMetadata, vals, names);
      await onSubmit(payload, submitKind, pendingFiles);
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSubmitting(false);
    }
  };

  const runCustomButton = async (btn: IFormCustomButtonConfig): Promise<void> => {
    const actions = btn.actions ?? [];
    const deltas = collectButtonOverlayDeltas(actions);
    const mergedValues = applyButtonActionsToValues(actions, values);
    const mergedOverlay = {
      show: cloneStringSet(buttonOverlay.show),
      hide: cloneStringSet(buttonOverlay.hide),
    };
    for (let i = 0; i < deltas.show.length; i++) mergedOverlay.show.add(deltas.show[i]);
    for (let i = 0; i < deltas.hide.length; i++) mergedOverlay.hide.add(deltas.hide[i]);
    flushSync(() => {
      setValues(mergedValues);
      setButtonOverlay(mergedOverlay);
    });
    const behavior = btn.behavior ?? 'actionsOnly';
    if (behavior === 'close') {
      onDismiss();
      return;
    }
    if (behavior === 'draft') {
      await handleSave('draft', { valuesOverride: mergedValues, buttonOverlayOverride: mergedOverlay });
    } else if (behavior === 'submit') {
      await handleSave('submit', { valuesOverride: mergedValues, buttonOverlayOverride: mergedOverlay });
    }
  };

  const steps = formManager.steps?.length ? formManager.steps : null;
  const [stepIndex, setStepIndex] = useState(0);
  const currentStepFieldSet = useMemo(() => {
    if (!steps) return null;
    const s = steps[stepIndex];
    return new Set(s.fieldNames);
  }, [steps, stepIndex]);

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
      if (formMode === 'view') {
        return (
          <Stack key={name} tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 12 } }}>
            <Label>{fc.label ?? 'Anexos ao item'}</Label>
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
          required={derived.fieldRequired[name] === true}
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
    const req = derived.fieldRequired[name] === true;
    const comp = derived.computedDisplay[name];
    if (comp !== undefined) {
      return (
        <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 12 } }}>
          <Label required={req}>{label}</Label>
          <Text styles={{ root: { color: '#323130' } }}>{String(comp)}</Text>
          {help && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
        </Stack>
      );
    }

    const common = { disabled: readOnly, errorMessage: err };

    switch (m.MappedType) {
      case 'boolean':
        return (
          <Toggle
            key={name}
            label={label}
            onText="Sim"
            offText="Não"
            checked={values[name] === true || values[name] === 1}
            onChange={(_, c) => updateField(name, !!c)}
            disabled={readOnly}
          />
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
            required={req}
            {...common}
            description={help}
          />
        );
      case 'datetime':
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 8 } }}>
            <Label required={req}>{label}</Label>
            <DatePicker
              value={values[name] ? new Date(String(values[name])) : undefined}
              onSelectDate={(d) => updateField(name, d ? d.toISOString() : null)}
              disabled={readOnly}
              textField={{ errorMessage: err }}
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
            required={req}
            errorMessage={err}
            disabled={readOnly}
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
            required={req}
            errorMessage={err}
            disabled={readOnly}
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
            required={req}
            errorMessage={err}
            disabled={readOnly}
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
            required={req}
            {...common}
            description={help ?? 'Informe o ID numérico do usuário no site.'}
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
            required={req}
            {...common}
            description={help}
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
            required={req}
            {...common}
            description={help}
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
      if (scope === 'main' && currentStepFieldSet && !currentStepFieldSet.has(fc.internalName)) continue;
      const sid = derived.effectiveSectionByField[fc.internalName] ?? fc.sectionId ?? formManager.sections[0]?.id ?? 'main';
      const arr = bySection.get(sid) ?? [];
      arr.push(fc);
      bySection.set(sid, arr);
    }
    const out: React.ReactNode[] = [];
    for (let s = 0; s < formManager.sections.length; s++) {
      const sec = formManager.sections[s];
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

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 8 } }}>
      {formError && <MessageBar messageBarType={MessageBarType.error}>{formError}</MessageBar>}
      {localErrors._attachments && (
        <MessageBar messageBarType={MessageBarType.error}>{localErrors._attachments}</MessageBar>
      )}
      {localErrors._async && <MessageBar messageBarType={MessageBarType.error}>{localErrors._async}</MessageBar>}
      {derived.messages.map((m) => (
        <MessageBar
          key={m.ruleId}
          messageBarType={m.variant === 'error' ? MessageBarType.error : m.variant === 'warning' ? MessageBarType.warning : MessageBarType.info}
        >
          {m.text}
        </MessageBar>
      ))}
      {steps && steps.length > 1 && (
        <FormStepNavigation
          steps={steps}
          activeIndex={stepIndex}
          onStepSelect={setStepIndex}
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
      <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
        {(formManager.customButtons ?? [])
          .filter((b) => {
            if (b.modes !== undefined && b.modes.length === 0) return false;
            if (!b.modes?.length) return true;
            return b.modes.indexOf(formMode) !== -1;
          })
          .map((b) =>
            b.appearance === 'primary' ? (
              <PrimaryButton
                key={b.id}
                text={b.label}
                onClick={() => void runCustomButton(b)}
                disabled={submitting}
              />
            ) : (
              <DefaultButton
                key={b.id}
                text={b.label}
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
        {steps && (
          <FormStepPrevNextNav
            variant={formManager.stepNavButtons ?? 'fluent'}
            stepIndex={stepIndex}
            stepCount={steps.length}
            onPrev={() => setStepIndex((i) => Math.max(0, i - 1))}
            onNext={() => setStepIndex((i) => Math.min(steps.length - 1, i + 1))}
            disabled={submitting}
          />
        )}
      </Stack>
      {modalOpen && (
        <Stack styles={{ root: { borderTop: '1px solid #edebe9', paddingTop: 16 } }} tokens={{ childrenGap: 12 }}>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Campos adicionais</Text>
          {renderFields('modal')}
          <DefaultButton text="Fechar modal" onClick={() => setModalOpen(false)} />
        </Stack>
      )}
    </Stack>
  );
};
