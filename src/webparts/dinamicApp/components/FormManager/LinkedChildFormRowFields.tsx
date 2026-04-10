import * as React from 'react';
import { useMemo, useState, useEffect, useCallback } from 'react';
import {
  Stack,
  Text,
  TextField,
  Toggle,
  Dropdown,
  IDropdownOption,
  DatePicker,
  Label,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type { IFormFieldConfig, IFormLinkedChildFormConfig } from '../../core/config/types/formManager';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_BANNER_INTERNAL_PREFIX,
  isFormBannerFieldConfig,
} from '../../core/config/types/formManager';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import {
  buildFormDerivedState,
  type IFormRuleRuntimeContext,
} from '../../core/formManager/formRuleEngine';
import { linkedChildFormAsManagerConfig } from '../../core/formManager/formLinkedChildSync';
import { ItemsService, UsersService } from '../../../../services';

const REQ_EMPTY_BORDER = '#a4262c';

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

function parseUrlFieldValue(v: unknown): { Url: string; Description: string } {
  if (v === null || v === undefined) return { Url: '', Description: '' };
  if (typeof v === 'object' && v !== null && 'Url' in v) {
    const o = v as Record<string, unknown>;
    return { Url: String(o.Url ?? ''), Description: String(o.Description ?? '') };
  }
  const s = String(v);
  const comma = s.indexOf(',');
  if (comma !== -1) {
    return { Url: s.slice(0, comma).trim(), Description: s.slice(comma + 1).trim() };
  }
  return { Url: s, Description: '' };
}

function dropdownReqStyles(showReq: boolean | undefined) {
  return showReq
    ? {
        dropdown: {
          borderColor: REQ_EMPTY_BORDER,
          borderWidth: 1,
          borderStyle: 'solid',
          borderRadius: 2,
        },
      }
    : undefined;
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
}) => {
  const itemsService = useMemo(() => new ItemsService(), []);
  const usersService = useMemo(() => new UsersService(), []);
  const [lookupOptions, setLookupOptions] = useState<Record<string, IDropdownOption[]>>({});
  const [siteUserOptions, setSiteUserOptions] = useState<IDropdownOption[]>([{ key: '', text: '—' }]);

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

  const runtimeCtx: IFormRuleRuntimeContext = useMemo(
    () => ({
      formMode,
      values,
      submitKind: 'submit',
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
    }),
    [formMode, values, userGroupTitles, currentUserId, authorId, dynamicContext]
  );

  const derived = useMemo(
    () => buildFormDerivedState(shell, fieldConfigs, runtimeCtx, undefined),
    [shell, fieldConfigs, runtimeCtx]
  );

  const updateField = useCallback(
    (name: string, v: unknown): void => {
      onChange({ ...values, [name]: v });
    },
    [onChange, values]
  );

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
    async (fieldName: string, odataFilter?: string): Promise<void> => {
      const m = metaByName.get(fieldName);
      if (!m?.LookupList) return;
      try {
        const filter = odataFilter?.trim() ? odataFilter : undefined;
        const lf = m.LookupField || 'Title';
        const rows = await itemsService.getItems<Record<string, unknown>>(m.LookupList, {
          select: ['Id', lf],
          filter,
          top: 200,
        });
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

  const lookupFetchKey = useMemo(() => {
    const parts: string[] = [];
    for (let i = 0; i < orderedFieldConfigs.length; i++) {
      const fn = orderedFieldConfigs[i].internalName;
      const m = metaByName.get(fn);
      if (m?.MappedType !== 'lookup' && m?.MappedType !== 'lookupmulti') continue;
      const listId = String(m.LookupList ?? '');
      const disp = String(m.LookupField || 'Title');
      const lf = derived.lookupFilters[fn];
      if (lf) {
        const pid = lookupIdFromValue(values[lf.parentField]);
        parts.push(
          `${fn}\t${listId}\t${disp}\t${lf.parentField}\t${lf.odataFilterTemplate}\t${pid === undefined ? '' : String(pid)}`
        );
      } else {
        parts.push(`${fn}\t${listId}\t${disp}\t`);
      }
    }
    parts.sort();
    return parts.join('\n');
  }, [orderedFieldConfigs, metaByName, derived.lookupFilters, values]);

  useEffect(() => {
    let cancelled = false;
    void (async (): Promise<void> => {
      for (let i = 0; i < orderedFieldConfigs.length; i++) {
        if (cancelled) return;
        const fn = orderedFieldConfigs[i].internalName;
        const m = metaByName.get(fn);
        if (m?.MappedType === 'lookup' || m?.MappedType === 'lookupmulti') {
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
    return (): void => {
      cancelled = true;
    };
  }, [lookupFetchKey, loadLookupOptions, orderedFieldConfigs, metaByName, derived.lookupFilters, values]);

  const renderFieldControl = (fc: IFormFieldConfig): React.ReactNode => {
    const name = fc.internalName;
    if (derived.fieldVisible[name] === false) return null;
    if (name === FORM_ATTACHMENTS_FIELD_INTERNAL || isFormBannerFieldConfig(fc)) return null;

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
        <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 8 } }}>
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
          <Stack key={name} tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 8 } }}>
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
        const raw = (m.Choices ?? []).map((c) => ({ key: c, text: c }));
        const opts: IDropdownOption[] = !isRequired ? [{ key: '', text: '—' }, ...raw] : raw;
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            options={opts}
            selectedKey={
              values[name] !== undefined && values[name] !== null && String(values[name]) !== ''
                ? String(values[name])
                : ''
            }
            onChange={(_, o) => o && updateField(name, o.key === '' ? null : o.key)}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty)}
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
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'lookup': {
        const id = lookupIdFromValue(values[name]);
        const baseOpts = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const opts =
          id !== undefined && id > 0
            ? mergeOptionsForIds(baseOpts, [{ id, label: userTitleFromValue(values[name]) }])
            : baseOpts;
        return (
          <Dropdown
            key={name}
            label={label}
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
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'lookupmulti': {
        const selected = normalizeIdTitleArray(values[name]);
        const baseRaw = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const baseOpts = baseRaw.filter((o) => o.key !== '');
        const extra = selected.map((x) => ({ id: x.Id, label: x.Title }));
        const opts = mergeOptionsForIds(baseOpts, extra);
        const keys = selected.map((x) => String(x.Id));
        return (
          <Dropdown
            key={name}
            label={label}
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
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'user': {
        const id = lookupIdFromValue(values[name]);
        const baseOpts = !isRequired ? siteUserOptions : siteUserOptions.filter((o) => o.key !== '');
        const opts =
          id !== undefined && id > 0
            ? mergeOptionsForIds(baseOpts, [{ id, label: userTitleFromValue(values[name]) }])
            : baseOpts;
        return (
          <Dropdown
            key={name}
            label={label}
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
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'usermulti': {
        const selected = normalizeIdTitleArray(values[name]);
        const baseOpts = siteUserOptions.filter((o) => o.key !== '');
        const extra = selected.map((x) => ({ id: x.Id, label: x.Title }));
        const opts = mergeOptionsForIds(baseOpts, extra);
        const keys = selected.map((x) => String(x.Id));
        return (
          <Dropdown
            key={name}
            label={label}
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
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'url': {
        const uv = parseUrlFieldValue(values[name]);
        return (
          <Stack key={name} tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 8 } }}>
            <Label required={isRequired}>{label}</Label>
            <TextField
              label="Endereço web"
              placeholder="https://"
              value={uv.Url}
              onChange={(_, v) => updateField(name, { Url: v ?? '', Description: uv.Description })}
              disabled={readOnly}
              errorMessage={err}
              styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
            />
            <TextField
              label="Descrição a apresentar"
              value={uv.Description}
              onChange={(_, v) => updateField(name, { Url: uv.Url, Description: v ?? '' })}
              disabled={readOnly}
            />
            {help && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
          </Stack>
        );
      }
      case 'multiline':
        return (
          <TextField
            key={name}
            label={label}
            multiline
            rows={3}
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

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {orderedFieldConfigs.map((fc) => renderFieldControl(fc))}
    </Stack>
  );
};
