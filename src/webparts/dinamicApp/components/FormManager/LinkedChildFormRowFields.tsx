import * as React from 'react';
import { useMemo, useState, useEffect, useCallback, useRef } from 'react';
import {
  Stack,
  Text,
  TextField,
  Toggle,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Label,
  useTheme,
} from '@fluentui/react';
import type { IStyle } from '@fluentui/react/lib/Styling';
import type { IFieldMetadata } from '../../../../services';
import type { IFormFieldConfig, IFormLinkedChildFormConfig } from '../../core/config/types/formManager';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_BANNER_INTERNAL_PREFIX,
  isFormAlertFieldConfig,
  isFormBannerFieldConfig,
  resolveAlertPlacement,
  resolveTextareaRows,
} from '../../core/config/types/formManager';
import { FormManagerAlertBlock } from './FormManagerAlertBlock';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import {
  buildFormDerivedState,
  evaluateValidateDateRulesForField,
  findEnabledSetComputedRule,
  getMergedValidateValueNumberBounds,
  clampNumberToOptionalBounds,
  withRuleRuntimeDynamicContext,
  type IFormRuleRuntimeContext,
} from '../../core/formManager/formRuleEngine';
import { buildValidateDateCalendarProps } from '../../core/formManager/validateDateCalendarProps';
import { linkedChildFormAsManagerConfig } from '../../core/formManager/formLinkedChildSync';
import { FLUENT_DATE_PICKER_PT_BR } from '../../core/formManager/fluentDatePickerPtBr';
import { applyTextTransformsToRecordValues } from '../../core/formManager/formTextValueTransform';
import {
  buildLookupDropdownSelectRaw,
  buildLookupODataFilter,
  hasConfiguredLookupFilter,
  isParentValueReadyForLookupFilter,
  lookupRowToOptionText,
  resolveLookupFormLabelInternalName,
} from '../../core/formManager/lookupFormDropdownHelpers';
import { shouldRenderMultilineNoteAsHtml } from '../../core/formManager/sharePointNoteHtml';
import { MultilineReadonlyHtml } from './MultilineReadonlyHtml';
import { multiSelectDropdownStyles, renderMultiSelectDropdownTitle } from './formMultiSelectDropdownUi';
import { ItemsService, UsersService, FieldsService } from '../../../../services';
import { IMaskInput } from 'react-imask';
import { resolveTextInputMaskOptions } from '../../core/formManager/formTextInputMasks';
import { parseUrlFieldValue } from '../../core/formManager/formUrlUtils';

const REQ_EMPTY_BORDER = '#a4262c';

const FORM_FIELD_CURSOR_DISABLED = 'not-allowed';

function lookupIdFromValue(v: unknown): number | undefined {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (typeof id === 'number') return id;
  }
  return undefined;
}

function normalizeIdTitleArray(v: unknown): { Id: number; Title: string }[] {
  if (v === null || v === undefined) return [];
  if (Array.isArray(v)) {
    const out: { Id: number; Title: string }[] = [];
    for (let i = 0; i < v.length; i++) {
      const x = v[i];
      const id = lookupIdFromValue(x);
      if (id === undefined) continue;
      let title = '';
      if (typeof x === 'object' && x !== null && 'Title' in x) {
        title = String((x as Record<string, unknown>).Title ?? '');
      }
      out.push({ Id: id, Title: title });
    }
    return out;
  }
  const id = lookupIdFromValue(v);
  return id !== undefined ? [{ Id: id, Title: '' }] : [];
}

function mergeOptionsForIds(
  opts: IDropdownOption[],
  entries: { id: number; label: string }[]
): IDropdownOption[] {
  const keys = new Set(opts.map((o) => String(o.key)));
  let out = opts;
  for (let i = 0; i < entries.length; i++) {
    const k = String(entries[i].id);
    if (keys.has(k)) continue;
    keys.add(k);
    out = [...out, { key: k, text: entries[i].label.trim() || `#${entries[i].id}` }];
  }
  return out;
}

function userTitleFromValue(v: unknown): string {
  if (typeof v === 'object' && v !== null && 'Title' in v) {
    return String((v as Record<string, unknown>).Title ?? '');
  }
  return '';
}

function dropdownReqStyles(
  showReq: boolean | undefined,
  disabled?: boolean
): {
  dropdown?: IStyle;
  title?: IStyle;
  caretDown?: IStyle;
  caretDownWrapper?: IStyle;
} | undefined {
  const out: {
    dropdown?: IStyle;
    title?: IStyle;
    caretDown?: IStyle;
    caretDownWrapper?: IStyle;
  } = {};
  const dropdown: Record<string, string | number> = {};
  if (showReq) {
    Object.assign(dropdown, {
      borderColor: REQ_EMPTY_BORDER,
      borderWidth: 1,
      borderStyle: 'solid',
      borderRadius: 2,
    });
  }
  if (disabled) {
    const text = '#201f1e';
    out.title = {
      color: text,
      opacity: 1,
      WebkitTextFillColor: text,
      cursor: FORM_FIELD_CURSOR_DISABLED,
    };
    out.caretDownWrapper = { cursor: FORM_FIELD_CURSOR_DISABLED };
    out.caretDown = { color: '#605e5c', opacity: 1, cursor: FORM_FIELD_CURSOR_DISABLED };
    Object.assign(dropdown, { color: text, opacity: 1, cursor: FORM_FIELD_CURSOR_DISABLED });
  }
  if (Object.keys(dropdown).length) {
    out.dropdown = dropdown as IStyle;
  }
  return Object.keys(out).length ? out : undefined;
}

function stylesTextFieldRequiredEmpty(
  active: boolean,
  disabled?: boolean
): { root?: IStyle; fieldGroup?: IStyle; field?: IStyle; icon?: IStyle } | undefined {
  const fieldGroupMerge: Record<string, string | number> = {};
  if (active) {
    Object.assign(fieldGroupMerge, {
      borderColor: REQ_EMPTY_BORDER,
      borderWidth: 1,
      borderStyle: 'solid',
      borderRadius: 2,
    });
  }
  if (disabled) {
    Object.assign(fieldGroupMerge, { cursor: FORM_FIELD_CURSOR_DISABLED });
  }
  const out: { root?: IStyle; fieldGroup?: IStyle; field?: IStyle; icon?: IStyle } = {};
  if (Object.keys(fieldGroupMerge).length) {
    out.fieldGroup = fieldGroupMerge as IStyle;
  }
  if (disabled) {
    out.root = { cursor: FORM_FIELD_CURSOR_DISABLED };
    out.icon = { cursor: FORM_FIELD_CURSOR_DISABLED };
    out.field = {
      color: '#201f1e',
      WebkitTextFillColor: '#201f1e',
      opacity: 1,
      cursor: FORM_FIELD_CURSOR_DISABLED,
      selectors: {
        '::placeholder': {
          color: '#605e5c',
          opacity: 1,
        },
      },
    };
  }
  return Object.keys(out).length ? out : undefined;
}

function isValueEmptyForRequired(v: unknown, mappedType: string): boolean {
  if (mappedType === 'boolean') {
    return v === undefined || v === null;
  }
  if (mappedType === 'url') {
    if (v === null || v === undefined) return true;
    if (typeof v === 'object' && v !== null && 'Url' in v) {
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

export type TLinkedChildFormRowFieldLayout = 'stack' | 'compact' | 'tableCells';

export interface ILinkedChildFormRowFieldsProps {
  childForm: IFormLinkedChildFormConfig;
  fieldMetadata: IFieldMetadata[];
  values: Record<string, unknown>;
  onChange: (next: Record<string, unknown>) => void;
  formMode: 'create' | 'edit' | 'view';
  userGroupTitles: string[];
  currentUserId: number;
  authorId: number | undefined;
  dynamicContext: IDynamicContext;
  localErrors?: Record<string, string>;
  fieldLayout?: TLinkedChildFormRowFieldLayout;
  /** Linha já gravada na lista vinculada (mostrar valor gravado em vez da expressão calculada). */
  rowPersisted?: boolean;
}

export const LinkedChildFormRowFields: React.FC<ILinkedChildFormRowFieldsProps> = ({
  childForm,
  fieldMetadata,
  values,
  onChange,
  formMode,
  userGroupTitles,
  currentUserId,
  authorId,
  dynamicContext,
  localErrors = {},
  fieldLayout = 'stack',
  rowPersisted = false,
}) => {
  const theme = useTheme();
  const itemsService = useMemo(() => new ItemsService(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);
  const lookupDestMetaCacheRef = useRef<Record<string, IFieldMetadata[]>>({});
  const usersService = useMemo(() => new UsersService(), []);
  const [lookupOptions, setLookupOptions] = useState<Record<string, IDropdownOption[]>>({});
  const [siteUserOptions, setSiteUserOptions] = useState<IDropdownOption[]>([{ key: '', text: '—' }]);
  const [datePickBlockErr, setDatePickBlockErr] = useState<Record<string, string>>({});

  const shell = useMemo(() => linkedChildFormAsManagerConfig(childForm), [childForm]);
  const fieldConfigs = childForm.fields;
  const metaByName = useMemo(
    () => new Map(fieldMetadata.map((f) => [f.InternalName, f])),
    [fieldMetadata]
  );

  const mainNames = useMemo(() => {
    const st = childForm.steps?.find((s) => s.id === 'main');
    return st?.fieldNames ?? [];
  }, [childForm.steps]);

  const orderedFieldConfigs = useMemo(() => {
    const out: IFormFieldConfig[] = [];
    for (let i = 0; i < mainNames.length; i++) {
      const n = mainNames[i];
      if (n === FORM_ATTACHMENTS_FIELD_INTERNAL) continue;
      if (n.indexOf(FORM_BANNER_INTERNAL_PREFIX) === 0) continue;
      if (n === childForm.parentLookupFieldInternalName.trim()) continue;
      const fc = fieldConfigs.find((f) => f.internalName === n);
      if (fc) out.push(fc);
    }
    return out;
  }, [mainNames, fieldConfigs, childForm.parentLookupFieldInternalName]);

  const datetimeFieldInternalNames = useMemo(
    () =>
      orderedFieldConfigs
        .map((fc) => fc.internalName)
        .filter((n) => metaByName.get(n)?.MappedType === 'datetime'),
    [orderedFieldConfigs, metaByName]
  );

  const runtimeCtx: IFormRuleRuntimeContext = useMemo(
    () => ({
      formMode,
      values,
      submitKind: 'submit',
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext: withRuleRuntimeDynamicContext(dynamicContext, currentUserId),
    }),
    [formMode, values, userGroupTitles, currentUserId, authorId, dynamicContext]
  );

  const derived = useMemo(
    () => buildFormDerivedState(shell, fieldConfigs, runtimeCtx, undefined, metaByName),
    [shell, fieldConfigs, runtimeCtx, metaByName]
  );

  const validateValueNumberMergedByField = useMemo(() => {
    const rules = shell.rules ?? [];
    const vis = derived.fieldVisible;
    const out: Record<string, { minNumber?: number; maxNumber?: number }> = {};
    const ctxSlice = {
      formMode,
      values,
      userGroupTitles,
      dynamicContext: withRuleRuntimeDynamicContext(dynamicContext, currentUserId),
    };
    for (let i = 0; i < orderedFieldConfigs.length; i++) {
      const n = orderedFieldConfigs[i].internalName;
      if (vis[n] === false) continue;
      const b = getMergedValidateValueNumberBounds(rules, n, ctxSlice, vis);
      if (b && (b.minNumber !== undefined || b.maxNumber !== undefined)) {
        out[n] = { minNumber: b.minNumber, maxNumber: b.maxNumber };
      }
    }
    return out;
  }, [
    shell.rules,
    orderedFieldConfigs,
    formMode,
    values,
    userGroupTitles,
    dynamicContext,
    currentUserId,
    derived.fieldVisible,
  ]);

  const validateDateCalendarPropsByField = useMemo(() => {
    const rules = shell.rules ?? [];
    const vis = derived.fieldVisible;
    const out: Record<string, ReturnType<typeof buildValidateDateCalendarProps>> = {};
    const dyn = withRuleRuntimeDynamicContext(dynamicContext, currentUserId);
    const paramsBase = {
      formMode,
      submitKind: 'submit' as const,
      userGroupTitles,
      dynamicContext: dyn,
      fieldVisible: (fn: string) => vis[fn] !== false,
      now: new Date(),
    };
    for (let i = 0; i < datetimeFieldInternalNames.length; i++) {
      const n = datetimeFieldInternalNames[i];
      if (vis[n] === false) continue;
      const p = buildValidateDateCalendarProps(rules, n, values, paramsBase);
      if (p.minDate || p.maxDate || (p.restrictedDates && p.restrictedDates.length > 0)) {
        out[n] = p;
      }
    }
    return out;
  }, [
    shell.rules,
    datetimeFieldInternalNames,
    values,
    formMode,
    userGroupTitles,
    dynamicContext,
    currentUserId,
    derived.fieldVisible,
  ]);

  const fieldConfigByInternalName = useMemo(
    () => new Map(fieldConfigs.map((fc) => [fc.internalName, fc])),
    [fieldConfigs]
  );

  useEffect(() => {
    const next = applyTextTransformsToRecordValues(values, fieldConfigs, metaByName);
    if (next !== values) onChange(next);
  }, [values, fieldConfigs, metaByName, onChange]);

  const updateField = useCallback(
    (name: string, v: unknown): void => {
      onChange({ ...values, [name]: v });
    },
    [onChange, values]
  );

  const applyLinkedDateSelect = useCallback(
    (name: string, d: Date | null | undefined) => {
      if (d === null || d === undefined) {
        updateField(name, null);
        setDatePickBlockErr((prev) => {
          if (!prev[name]) return prev;
          const { [name]: _, ...rest } = prev;
          return rest;
        });
        return;
      }
      const iso = d.toISOString();
      const nextValues = { ...values, [name]: iso };
      const msg = evaluateValidateDateRulesForField(shell.rules ?? [], name, nextValues, {
        formMode,
        submitKind: 'submit',
        userGroupTitles,
        dynamicContext: withRuleRuntimeDynamicContext(dynamicContext, currentUserId),
        fieldVisible: (fn) => derived.fieldVisible[fn] !== false,
        now: new Date(),
      });
      if (msg) {
        updateField(name, null);
        setDatePickBlockErr((prev) => ({ ...prev, [name]: msg }));
        return;
      }
      updateField(name, iso);
      setDatePickBlockErr((prev) => {
        if (!prev[name]) return prev;
        const { [name]: _, ...rest } = prev;
        return rest;
      });
    },
    [updateField, values, shell.rules, formMode, userGroupTitles, dynamicContext, currentUserId, derived]
  );

  useEffect(() => {
    if (formMode === 'view') return;
    const patches: Record<string, null> = {};
    const errPatches: Record<string, string> = {};
    const dyn = withRuleRuntimeDynamicContext(dynamicContext, currentUserId);
    const fv = (fn: string): boolean => derived.fieldVisible[fn] !== false;
    for (let i = 0; i < datetimeFieldInternalNames.length; i++) {
      const name = datetimeFieldInternalNames[i];
      if (!fv(name)) continue;
      const raw = values[name];
      if (raw === null || raw === undefined || raw === '') continue;
      const msg = evaluateValidateDateRulesForField(shell.rules ?? [], name, values, {
        formMode,
        submitKind: 'submit',
        userGroupTitles,
        dynamicContext: dyn,
        fieldVisible: fv,
        now: new Date(),
      });
      if (msg) {
        patches[name] = null;
        errPatches[name] = msg;
      }
    }
    const pk = Object.keys(patches);
    if (pk.length === 0) return;
    const nextVals = { ...values };
    let vChanged = false;
    for (let j = 0; j < pk.length; j++) {
      const k = pk[j];
      if (nextVals[k] !== null && nextVals[k] !== undefined) {
        nextVals[k] = null;
        vChanged = true;
      }
    }
    if (vChanged) onChange(nextVals);
    setDatePickBlockErr((prev) => {
      let changed = false;
      const next = { ...prev };
      for (let j = 0; j < pk.length; j++) {
        const k = pk[j];
        const e = errPatches[k];
        if (e !== undefined && next[k] !== e) {
          next[k] = e;
          changed = true;
        }
      }
      return changed ? next : prev;
    });
  }, [
    formMode,
    datetimeFieldInternalNames,
    values,
    derived.fieldVisible,
    shell.rules,
    userGroupTitles,
    dynamicContext,
    currentUserId,
    onChange,
  ]);

  useEffect(() => {
    let cancelled = false;
    void usersService
      .getSiteUsers()
      .then((users) => {
        if (cancelled) return;
        const sorted = [...users].sort((a, b) =>
          (a.Title || a.Email || '').localeCompare(b.Title || b.Email || '', undefined, { sensitivity: 'base' })
        );
        setSiteUserOptions([
          { key: '', text: '—' },
          ...sorted.map((u) => ({
            key: String(u.Id),
            text: (u.Title || u.Email || u.LoginName || '').trim() || `#${u.Id}`,
          })),
        ]);
      })
      .catch(() => {
        if (!cancelled) setSiteUserOptions([{ key: '', text: '—' }]);
      });
    return (): void => {
      cancelled = true;
    };
  }, [usersService]);

  const loadLookupOptions = useCallback(
    async (fieldName: string, lf?: { parentField: string; childField?: string; filterOperator?: string; odataFilterTemplate?: string }): Promise<void> => {
      const m = metaByName.get(fieldName);
      if (!m?.LookupList) return;
      const fc = fieldConfigByInternalName.get(fieldName);
      try {
        let filter: string | undefined;

        const listGuid = m.LookupList;
        let fieldMetaList: IFieldMetadata[] | undefined = lookupDestMetaCacheRef.current[listGuid];
        if (!fieldMetaList?.length) {
          try {
            const fetched = await fieldsService.getFields(listGuid);
            lookupDestMetaCacheRef.current = {
              ...lookupDestMetaCacheRef.current,
              [listGuid]: fetched,
            };
            fieldMetaList = fetched;
          } catch {
            fieldMetaList = undefined;
          }
        }

        if (lf) {
          if (lf.childField && lf.filterOperator) {
            const parentVal = values[lf.parentField];
            const parentMeta = metaByName.get(lf.parentField);
            const childFieldMeta = fieldMetaList?.find((x) => x.InternalName === lf.childField);
            filter = buildLookupODataFilter(lf.childField, lf.filterOperator, parentVal, parentMeta, childFieldMeta);
          } else if (lf.odataFilterTemplate) {
            const parentVal = values[lf.parentField];
            const pid = typeof parentVal === 'number' && isFinite(parentVal) ? parentVal :
              (typeof parentVal === 'object' && parentVal !== null && 'Id' in parentVal &&
               typeof (parentVal as Record<string, unknown>).Id === 'number'
               ? (parentVal as Record<string, unknown>).Id as number : undefined);
            if (pid !== undefined) filter = lf.odataFilterTemplate.split('{parent}').join(String(pid));
          }
        }

        const selectRaw = buildLookupDropdownSelectRaw(m, fc ?? {});
        const labelFieldName = resolveLookupFormLabelInternalName(m, fc ?? {});
        const labelMeta = fieldMetaList?.find((x) => x.InternalName === labelFieldName);

        const rows = await itemsService.getItems<Record<string, unknown>>(m.LookupList, {
          select: selectRaw,
          filter,
          top: 200,
          ...(fieldMetaList?.length ? { fieldMetadata: fieldMetaList } : {}),
        });

        const opts: IDropdownOption[] = [
          { key: '', text: '—' },
          ...rows.map((row) => ({
            key: String(row.Id),
            text: lookupRowToOptionText(row, labelFieldName, labelMeta, fc?.lookupOptionLabelSubProp),
            data: row,
          })),
        ];
        setLookupOptions((o) => ({ ...o, [fieldName]: opts }));
      } catch {
        setLookupOptions((o) => ({ ...o, [fieldName]: [] }));
      }
    },
    [itemsService, metaByName, fieldsService, fieldConfigByInternalName, values]
  );

  const lookupFetchKey = useMemo(() => {
    const parts: string[] = [];
    for (let i = 0; i < orderedFieldConfigs.length; i++) {
      const fn = orderedFieldConfigs[i].internalName;
      const m = metaByName.get(fn);
      if (m?.MappedType !== 'lookup' && m?.MappedType !== 'lookupmulti') continue;
      const listId = String(m.LookupList ?? '');
      const fc = orderedFieldConfigs[i];
      const labelDisp = resolveLookupFormLabelInternalName(m, fc ?? {});
      const extrasSig = JSON.stringify(fc?.lookupOptionExtraSelectFields ?? []);
      const subPropSig = fc?.lookupOptionLabelSubProp ?? '';
      const detailSig = JSON.stringify(fc?.lookupOptionDetailBelowFields ?? []);
      const lf = derived.lookupFilters[fn];
      if (lf) {
        const parentVal = values[lf.parentField];
        const parentId = typeof parentVal === 'number' && isFinite(parentVal) ? parentVal :
          (typeof parentVal === 'object' && parentVal !== null && 'Id' in parentVal &&
           typeof (parentVal as Record<string, unknown>).Id === 'number'
           ? (parentVal as Record<string, unknown>).Id as number : undefined);
        const parentSig = parentId !== undefined ? String(parentId) :
          typeof parentVal === 'string' ? parentVal :
          typeof parentVal === 'number' ? String(parentVal) : '';
        parts.push(
          `${fn}\t${listId}\t${labelDisp}\t${extrasSig}\t${subPropSig}\t${detailSig}\t${lf.parentField}\t${lf.childField ?? ''}\t${lf.filterOperator ?? ''}\t${parentSig}`
        );
      } else {
        parts.push(`${fn}\t${listId}\t${labelDisp}\t${extrasSig}\t${subPropSig}\t${detailSig}\t`);
      }
    }
    parts.sort();
    return parts.join('\n');
  }, [orderedFieldConfigs, metaByName, derived.lookupFilters, values]);

  const lookupDetailSnapshot = useMemo(() => {
    const out: Record<string, Record<string, unknown> | Record<string, unknown>[] | undefined> = {};
    for (let i = 0; i < orderedFieldConfigs.length; i++) {
      const fc = orderedFieldConfigs[i];
      const detailFns = fc.lookupOptionDetailBelowFields ?? [];
      if (!detailFns.length) continue;
      const m = metaByName.get(fc.internalName);
      if (!m || (m.MappedType !== 'lookup' && m.MappedType !== 'lookupmulti')) continue;
      const opts = lookupOptions[fc.internalName] ?? [];
      if (m.MappedType === 'lookup') {
        const id = lookupIdFromValue(values[fc.internalName]);
        if (!id) {
          out[fc.internalName] = undefined;
          continue;
        }
        const opt = opts.find((o) => String(o.key) === String(id));
        const data =
          opt && typeof opt === 'object' && 'data' in opt ? (opt as { data?: Record<string, unknown> }).data : undefined;
        out[fc.internalName] = data;
      } else {
        const sel = normalizeIdTitleArray(values[fc.internalName]);
        const many: Record<string, unknown>[] = [];
        for (let s = 0; s < sel.length; s++) {
          const opt = opts.find((o) => String(o.key) === String(sel[s].Id));
          const row =
            opt && typeof opt === 'object' && 'data' in opt ? (opt as { data?: Record<string, unknown> }).data : undefined;
          if (row) many.push(row);
        }
        out[fc.internalName] = many.length ? many : undefined;
      }
    }
    return out;
  }, [orderedFieldConfigs, metaByName, lookupOptions, values]);

  useEffect(() => {
    let cancelled = false;
    void (async (): Promise<void> => {
      for (let i = 0; i < orderedFieldConfigs.length; i++) {
        if (cancelled) return;
        const fn = orderedFieldConfigs[i].internalName;
        const m = metaByName.get(fn);
        if (m?.MappedType === 'lookup' || m?.MappedType === 'lookupmulti') {
          const lf = derived.lookupFilters[fn];
          await loadLookupOptions(fn, lf);
        }
      }
    })();
    return (): void => {
      cancelled = true;
    };
  }, [lookupFetchKey, loadLookupOptions, orderedFieldConfigs, metaByName, derived.lookupFilters, values]);

  type TRenderMode = 'default' | 'compact' | 'cell';

  const renderFieldControl = (fc: IFormFieldConfig, mode: TRenderMode = 'default'): React.ReactNode => {
    const name = fc.internalName;
    if (derived.fieldVisible[name] === false) return null;
    if (name === FORM_ATTACHMENTS_FIELD_INTERNAL || isFormBannerFieldConfig(fc)) return null;
    if (isFormAlertFieldConfig(fc)) {
      if (fc.sectionId === undefined) return null;
      if (fc.sectionId === 'fixos' || resolveAlertPlacement(fc) === 'inStep') {
        return (
          <Stack key={name} tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 8 } }}>
            <FormManagerAlertBlock
              alert={fc}
              values={values}
              dynamicContext={dynamicContext}
              userGroupTitles={userGroupTitles}
            />
          </Stack>
        );
      }
      return null;
    }

    const setComputedRule = findEnabledSetComputedRule(shell.rules, name, formMode);
    const compRaw = derived.computedDisplay[name];
    const comp =
      setComputedRule?.alwaysLiveComputed === true || !rowPersisted ? compRaw : undefined;
    if (comp !== undefined && (formMode === 'view' || fc.readOnly === true)) {
      const mComp = metaByName.get(name);
      const label = fc.label ?? mComp?.Title ?? name;
      const help = derived.dynamicHelpByField[name] ?? fc.helpText;
      const isRequired = derived.fieldRequired[name] === true || mComp?.Required === true;
      const mb = mode === 'compact' ? 4 : 8;
      const cell = mode === 'cell';
      const compShown =
        mComp?.MappedType === 'datetime'
          ? ((): string => {
              const s = String(comp);
              const ms = Date.parse(s);
              return !isNaN(ms) ? new Date(ms).toLocaleDateString('pt-BR') : s;
            })()
          : String(comp);
      return (
        <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: mb } }}>
          {!cell && <Label required={isRequired}>{label}</Label>}
          <Text styles={{ root: { color: '#323130' } }} title={cell ? label : undefined}>
            {compShown}
          </Text>
          {help && !cell && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
        </Stack>
      );
    }

    const m = metaByName.get(name);
    if (!m) return null;
    const mb = mode === 'compact' ? 4 : 8;
    const cell = mode === 'cell';
    const disabled = formMode === 'view' || derived.fieldDisabled[name] === true;
    const readOnly = derived.fieldReadOnly[name] === true || disabled;
    const err = localErrors[name] || datePickBlockErr[name];
    const label = fc.label ?? m.Title;
    const help = derived.dynamicHelpByField[name] ?? fc.helpText;
    const isRequired = derived.fieldRequired[name] === true || m.Required === true;
    const canFill = formMode !== 'view' && !readOnly;
    const mergedFieldValue = ((): unknown => {
      if (setComputedRule?.alwaysLiveComputed === true && comp !== undefined) return comp;
      if (comp === undefined) return values[name];
      const v = values[name];
      if (v === undefined || v === null) return comp;
      if (typeof v === 'string' && v.trim() === '') return comp;
      return v;
    })();
    const showReqEmpty = isRequired && canFill && isValueEmptyForRequired(mergedFieldValue, m.MappedType);

    const common = { disabled: readOnly, errorMessage: err };

    const renderLookupDetailsBelow = (fieldName: string, meta: IFieldMetadata): React.ReactNode => {
      if (cell) return null;
      const dfn = fc.lookupOptionDetailBelowFields ?? [];
      if (!dfn.length) return null;
      const snap = lookupDetailSnapshot[fieldName];
      if (snap === undefined) return null;
      const listGuid = String(meta.LookupList ?? '');
      const linkedMeta = lookupDestMetaCacheRef.current[listGuid] ?? [];
      const rowsArr = Array.isArray(snap) ? snap : [snap];
      if (rowsArr.length === 0) return null;
      return (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 2, width: '100%' } }}>
          {dfn.map((fn) => {
            const fm = linkedMeta.find((x) => x.InternalName === fn);
            const fieldLabel = fm ? fm.Title : fn;
            const parts = rowsArr.map((row) =>
              lookupRowToOptionText(row as Record<string, unknown>, fn, fm)
            );
            const val = parts.join('; ');
            return (
              <TextField
                key={fn}
                label={fieldLabel}
                value={val}
                readOnly
                disabled
                multiline={val.length > 100}
                rows={val.length > 100 ? 2 : 1}
                styles={{ root: { cursor: FORM_FIELD_CURSOR_DISABLED } }}
              />
            );
          })}
        </Stack>
      );
    };

    switch (m.MappedType) {
      case 'boolean':
        return (
          <Stack key={name} tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: mb } }}>
            {!cell && <Label required={isRequired}>{label}</Label>}
            <Toggle
              ariaLabel={label}
              onText="Sim"
              offText="Não"
              checked={mergedFieldValue === true || mergedFieldValue === 1}
              onChange={(_, c) => updateField(name, !!c)}
              disabled={readOnly}
              styles={readOnly ? { root: { cursor: FORM_FIELD_CURSOR_DISABLED } } : undefined}
            />
          </Stack>
        );
      case 'number':
      case 'currency': {
        const numBounds = validateValueNumberMergedByField[name];
        return (
          <TextField
            key={name}
            {...(cell ? { ariaLabel: label } : { label })}
            type="number"
            placeholder={fc.placeholder}
            value={mergedFieldValue !== null && mergedFieldValue !== undefined ? String(mergedFieldValue) : ''}
            onChange={(_, v) => {
              if (v === '') {
                updateField(name, null);
                return;
              }
              const parsed = Number(String(v).replace(',', '.'));
              if (!isFinite(parsed)) return;
              const maxN = numBounds?.maxNumber;
              const next = maxN !== undefined && parsed > maxN ? maxN : parsed;
              updateField(name, next);
            }}
            onBlur={() => {
              if (readOnly) return;
              if (mergedFieldValue === null || mergedFieldValue === undefined) return;
              const t = String(mergedFieldValue).trim();
              if (t === '') return;
              const parsed = Number(t.replace(',', '.'));
              if (!isFinite(parsed)) return;
              const c = clampNumberToOptionalBounds(parsed, numBounds);
              if (c !== parsed) updateField(name, c);
            }}
            required={isRequired}
            {...common}
            description={cell ? undefined : help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty, readOnly)}
            min={numBounds?.minNumber}
            max={numBounds?.maxNumber}
          />
        );
      }
      case 'datetime':
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: mb } }}>
            {!cell && <Label required={isRequired}>{label}</Label>}
            <DatePicker
              {...FLUENT_DATE_PICKER_PT_BR}
              minDate={validateDateCalendarPropsByField[name]?.minDate}
              maxDate={validateDateCalendarPropsByField[name]?.maxDate}
              calendarProps={validateDateCalendarPropsByField[name]}
              value={mergedFieldValue ? new Date(String(mergedFieldValue)) : undefined}
              onSelectDate={(d) => applyLinkedDateSelect(name, d ?? null)}
              disabled={readOnly}
              textField={{
                ...(cell ? { ariaLabel: label } : {}),
                disabled: readOnly,
                errorMessage: err,
                styles: stylesTextFieldRequiredEmpty(showReqEmpty, readOnly),
              }}
            />
            {help && !cell && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
          </Stack>
        );
      case 'choice': {
        const raw = (m.Choices ?? []).map((c) => ({ key: c, text: c }));
        const opts: IDropdownOption[] = !isRequired ? [{ key: '', text: '—' }, ...raw] : raw;
        return (
          <Dropdown
            key={name}
            {...(cell ? { ariaLabel: label } : { label })}
            placeholder={fc.placeholder}
            options={opts}
            selectedKey={
              mergedFieldValue !== undefined && mergedFieldValue !== null && String(mergedFieldValue) !== ''
                ? String(mergedFieldValue)
                : ''
            }
            onChange={(_, o) => o && updateField(name, o.key === '' ? null : o.key)}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty, readOnly)}
          />
        );
      }
      case 'multichoice': {
        const selected: string[] = Array.isArray(mergedFieldValue)
          ? (mergedFieldValue as string[])
          : typeof mergedFieldValue === 'string'
            ? String(mergedFieldValue).split(';').map((s) => s.trim()).filter(Boolean)
            : [];
        const opts: IDropdownOption[] = (m.Choices ?? []).map((c) => ({
          key: c,
          text: c,
          selected: selected.indexOf(c) !== -1,
        }));
        return (
          <Dropdown
            key={name}
            {...(cell ? { ariaLabel: label } : { label })}
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
            onRenderTitle={(opts) => renderMultiSelectDropdownTitle(theme, opts, readOnly)}
            styles={multiSelectDropdownStyles(showReqEmpty, readOnly)}
          />
        );
      }
      case 'lookup': {
        const lfDrop = derived.lookupFilters[name];
        const filterActive = hasConfiguredLookupFilter(lfDrop);
        const parentMetaDrop = lfDrop ? metaByName.get(lfDrop.parentField) : undefined;
        const parentReady =
          !filterActive ||
          isParentValueReadyForLookupFilter(
            lfDrop ? values[lfDrop.parentField] : undefined,
            parentMetaDrop
          );
        const lookupBlockedByParent = filterActive && !parentReady;
        const id = lookupIdFromValue(mergedFieldValue);
        const baseOpts = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const opts =
          id !== undefined && id > 0
            ? mergeOptionsForIds(baseOpts, [{ id, label: userTitleFromValue(mergedFieldValue) }])
            : baseOpts;
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: mb } }}>
            <Dropdown
              {...(cell ? { ariaLabel: label } : { label })}
              placeholder={fc.placeholder}
              options={opts}
              selectedKey={id !== undefined ? String(id) : ''}
              onChange={(_, o) => {
                if (!o || o.key === '') updateField(name, null);
                else {
                  const raw =
                    o && typeof o === 'object' && 'data' in o ? (o as { data?: Record<string, unknown> }).data : undefined;
                  if (raw && typeof raw === 'object' && typeof raw.Id === 'number') {
                    updateField(name, raw);
                  } else {
                    updateField(name, { Id: Number(o.key), Title: String(o.text ?? '') });
                  }
                }
              }}
              required={isRequired}
              errorMessage={err}
              disabled={readOnly || lookupBlockedByParent}
              styles={dropdownReqStyles(showReqEmpty, readOnly || lookupBlockedByParent)}
            />
            {help && !cell ? (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {help}
              </Text>
            ) : null}
            {renderLookupDetailsBelow(name, m)}
          </Stack>
        );
      }
      case 'lookupmulti': {
        const lfMulti = derived.lookupFilters[name];
        const filterActiveMulti = hasConfiguredLookupFilter(lfMulti);
        const parentMetaMulti = lfMulti ? metaByName.get(lfMulti.parentField) : undefined;
        const parentReadyMulti =
          !filterActiveMulti ||
          isParentValueReadyForLookupFilter(
            lfMulti ? values[lfMulti.parentField] : undefined,
            parentMetaMulti
          );
        const lookupBlockedByParentMulti = filterActiveMulti && !parentReadyMulti;
        const selected = normalizeIdTitleArray(mergedFieldValue);
        const baseRaw = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const baseOpts = baseRaw.filter((o) => o.key !== '');
        const extra = selected.map((x) => ({ id: x.Id, label: x.Title }));
        const opts = mergeOptionsForIds(baseOpts, extra);
        const keys = selected.map((x) => String(x.Id));
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: mb } }}>
            <Dropdown
              {...(cell ? { ariaLabel: label } : { label })}
              placeholder={fc.placeholder}
              multiSelect
              options={opts}
              selectedKeys={keys}
              onChange={(_, o) => {
                if (!o || o.key === '') return;
                const k = String(o.key);
                const hit = selected.findIndex((x) => String(x.Id) === k);
                const raw =
                  o && typeof o === 'object' && 'data' in o ? (o as { data?: Record<string, unknown> }).data : undefined;
                const nextItem =
                  raw && typeof raw === 'object' && typeof raw.Id === 'number'
                    ? raw
                    : { Id: Number(o.key), Title: String(o.text ?? '') };
                const next =
                  hit === -1
                    ? [...selected, nextItem]
                    : selected.filter((_, i) => i !== hit);
                updateField(name, next);
              }}
              required={isRequired}
              errorMessage={err}
              disabled={readOnly || lookupBlockedByParentMulti}
              onRenderTitle={(opts) =>
                renderMultiSelectDropdownTitle(theme, opts, readOnly || lookupBlockedByParentMulti)
              }
              styles={multiSelectDropdownStyles(showReqEmpty, readOnly || lookupBlockedByParentMulti)}
            />
            {help && !cell ? (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {help}
              </Text>
            ) : null}
            {renderLookupDetailsBelow(name, m)}
          </Stack>
        );
      }
      case 'user': {
        const id = lookupIdFromValue(mergedFieldValue);
        const baseOpts = !isRequired ? siteUserOptions : siteUserOptions.filter((o) => o.key !== '');
        const opts =
          id !== undefined && id > 0
            ? mergeOptionsForIds(baseOpts, [{ id, label: userTitleFromValue(mergedFieldValue) }])
            : baseOpts;
        return (
          <Dropdown
            key={name}
            {...(cell ? { ariaLabel: label } : { label })}
            placeholder={fc.placeholder}
            options={opts}
            selectedKey={id !== undefined ? String(id) : ''}
            onChange={(_, o) => {
              if (!o || o.key === '') updateField(name, null);
              else updateField(name, { Id: Number(o.key), Title: String(o.text ?? '') });
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty, readOnly)}
          />
        );
      }
      case 'usermulti': {
        const selected = normalizeIdTitleArray(mergedFieldValue);
        const baseOpts = siteUserOptions.filter((o) => o.key !== '');
        const extra = selected.map((x) => ({ id: x.Id, label: x.Title }));
        const opts = mergeOptionsForIds(baseOpts, extra);
        const keys = selected.map((x) => String(x.Id));
        return (
          <Dropdown
            key={name}
            {...(cell ? { ariaLabel: label } : { label })}
            placeholder={fc.placeholder}
            multiSelect
            options={opts}
            selectedKeys={keys}
            onChange={(_, o) => {
              if (!o || o.key === '') return;
              const k = String(o.key);
              const hit = selected.findIndex((x) => String(x.Id) === k);
              const next =
                hit === -1
                  ? [...selected, { Id: Number(o.key), Title: String(o.text ?? '') }]
                  : selected.filter((_, i) => i !== hit);
              updateField(name, next);
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            onRenderTitle={(opts) => renderMultiSelectDropdownTitle(theme, opts, readOnly)}
            styles={multiSelectDropdownStyles(showReqEmpty, readOnly)}
          />
        );
      }
      case 'url': {
        const uv = parseUrlFieldValue(mergedFieldValue);
        return (
          <Stack key={name} tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: mb } }}>
            {!cell && <Label required={isRequired}>{label}</Label>}
            <TextField
              label={cell ? undefined : 'Endereço web'}
              ariaLabel={cell ? `${label} · Endereço web` : undefined}
              placeholder="https://"
              value={uv.Url}
              onChange={(_, v) => updateField(name, { Url: v ?? '', Description: uv.Description })}
              disabled={readOnly}
              errorMessage={err}
              styles={stylesTextFieldRequiredEmpty(showReqEmpty, readOnly)}
            />
            <TextField
              label={cell ? undefined : 'Descrição a apresentar'}
              ariaLabel={cell ? `${label} · Descrição` : undefined}
              value={uv.Description}
              onChange={(_, v) => updateField(name, { Url: uv.Url, Description: v ?? '' })}
              disabled={readOnly}
              styles={stylesTextFieldRequiredEmpty(showReqEmpty, readOnly)}
            />
            {help && !cell && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
          </Stack>
        );
      }
      case 'text': {
        const rawSingle =
          mergedFieldValue !== null && mergedFieldValue !== undefined ? String(mergedFieldValue) : '';
        const maskOpts = resolveTextInputMaskOptions(fc.textInputMaskKind, fc.textInputMaskCustomPattern);
        const inputBorder = err ? theme.semanticColors.errorText : theme.semanticColors.inputBorder;
        const maskInputStyle: React.CSSProperties = {
          width: '100%',
          boxSizing: 'border-box',
          minHeight: 32,
          padding: '0 8px',
          font: 'inherit',
          color: theme.semanticColors.inputText,
          backgroundColor: theme.semanticColors.inputBackground,
          borderWidth: 1,
          borderStyle: 'solid',
          borderColor: inputBorder,
          borderRadius: 2,
          outline: 'none',
          ...(readOnly ? { cursor: FORM_FIELD_CURSOR_DISABLED } : {}),
        };
        if (maskOpts) {
          return (
            <Stack
              key={name}
              tokens={{ childrenGap: 4 }}
              styles={{
                root: {
                  marginBottom: mb,
                  ...(showReqEmpty && !cell
                    ? {
                        borderLeft: `3px solid ${REQ_EMPTY_BORDER}`,
                        paddingLeft: 8,
                        paddingTop: 2,
                        paddingBottom: 2,
                      }
                    : {}),
                },
              }}
            >
              {!cell && <Label required={isRequired}>{label}</Label>}
              <IMaskInput
                {...maskOpts}
                value={rawSingle}
                disabled={readOnly}
                placeholder={fc.placeholder ?? undefined}
                onAccept={(val) => updateField(name, String(val ?? ''))}
                style={maskInputStyle}
                aria-invalid={err ? true : undefined}
                aria-required={isRequired ? true : undefined}
                aria-label={cell ? label : undefined}
              />
              {err ? (
                <Text variant="small" styles={{ root: { color: theme.semanticColors.errorText } }}>{err}</Text>
              ) : null}
              {help && !cell ? <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text> : null}
            </Stack>
          );
        }
        return (
          <TextField
            key={name}
            {...(cell ? { ariaLabel: label } : { label })}
            placeholder={fc.placeholder}
            value={rawSingle}
            onChange={(_, v) => updateField(name, v ?? '')}
            required={isRequired}
            {...common}
            description={cell ? undefined : help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty, readOnly)}
          />
        );
      }
      case 'multiline': {
        const raw =
          mergedFieldValue !== null && mergedFieldValue !== undefined ? String(mergedFieldValue) : '';
        if (readOnly && shouldRenderMultilineNoteAsHtml(m, raw)) {
          return (
            <MultilineReadonlyHtml
              key={name}
              label={label}
              required={isRequired}
              html={raw}
              help={cell ? undefined : help}
              showReqEmpty={showReqEmpty}
              showLabel={!cell}
            />
          );
        }
        return (
          <TextField
            key={name}
            {...(cell ? { ariaLabel: label } : { label })}
            multiline
            rows={resolveTextareaRows(fc, cell ? 2 : 3)}
            placeholder={fc.placeholder}
            value={raw}
            onChange={(_, v) => updateField(name, v ?? '')}
            required={isRequired}
            {...common}
            description={cell ? undefined : help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty, readOnly)}
          />
        );
      }
      default:
        return (
          <TextField
            key={name}
            {...(cell ? { ariaLabel: label } : { label })}
            placeholder={fc.placeholder}
            value={mergedFieldValue !== null && mergedFieldValue !== undefined ? String(mergedFieldValue) : ''}
            onChange={(_, v) => updateField(name, v ?? '')}
            required={isRequired}
            {...common}
            description={cell ? undefined : help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty, readOnly)}
          />
        );
    }
  };

  if (fieldLayout === 'tableCells') {
    return (
      <>
        {orderedFieldConfigs.map((fc) => (
          <td
            key={fc.internalName}
            style={{
              verticalAlign: 'top',
              padding: '8px 10px',
              borderBottom: '1px solid #edebe9',
              borderRight: '1px solid #edebe9',
              maxWidth: 300,
              minWidth: 80,
            }}
          >
            {renderFieldControl(fc, 'cell')}
          </td>
        ))}
      </>
    );
  }

  const stackMode: TRenderMode = fieldLayout === 'compact' ? 'compact' : 'default';

  return (
    <Stack tokens={{ childrenGap: fieldLayout === 'compact' ? 6 : 8 }}>
      {orderedFieldConfigs.map((fc) => renderFieldControl(fc, stackMode))}
    </Stack>
  );
};
