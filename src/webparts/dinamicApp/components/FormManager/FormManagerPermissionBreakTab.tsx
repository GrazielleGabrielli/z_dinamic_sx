import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import {
  Stack,
  Text,
  Toggle,
  Checkbox,
  Dropdown,
  IDropdownOption,
  TextField,
  Spinner,
  MessageBar,
  MessageBarType,
  DefaultButton,
  IconButton,
} from '@fluentui/react';
import { FieldsService, UsersService, filterSiteGroupsByNameQuery } from '../../../../services';
import type { IFieldMetadata, IGroupDetails, IPeoplePickerResult } from '../../../../services';
import type {
  IFormLinkedChildFormConfig,
  IFormManagerConfig,
  IFormManagerPermissionBreakConfig,
  IFormPermissionBreakAssignment,
} from '../../core/config/types/formManager';
import { resolveLinkedChildAttachmentRuntime } from '../../core/formManager/linkedChildAttachmentRuntime';

const ROLE_OPTIONS: IDropdownOption[] = [
  { key: 'Leitura', text: 'Leitura' },
  { key: 'Contribuir', text: 'Contribuir' },
  { key: 'Editar', text: 'Editar' },
  { key: 'Controlo total', text: 'Controlo total' },
];

const KIND_OPTIONS: IDropdownOption[] = [
  { key: 'siteGroup', text: 'Grupo do site' },
  { key: 'user', text: 'Pessoa do Site' },
  { key: 'field', text: 'Campo (Pessoa / Vários)' },
];

function newAssignmentId(): string {
  return `pb_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

function userFieldsFromMeta(meta: IFieldMetadata[]): IDropdownOption[] {
  return meta
    .filter((m) => m.MappedType === 'user' || m.MappedType === 'usermulti')
    .map((m) => ({ key: m.InternalName, text: `${m.Title} (${m.InternalName})` }));
}

export interface IFormManagerPermissionBreakTabProps {
  primaryListTitle: string;
  primaryMeta: IFieldMetadata[];
  linkedChildForms: IFormLinkedChildFormConfig[];
  /** Snapshot mínimo para resolver anexos em biblioteca por lista vinculada. */
  formManagerForResolve: Pick<IFormManagerConfig, 'attachmentStorageKind' | 'attachmentLibrary' | 'linkedChildForms'>;
  mainAttachmentLibraryEnabled: boolean;
  value: IFormManagerPermissionBreakConfig;
  onChange: (next: IFormManagerPermissionBreakConfig) => void;
  siteGroups: IGroupDetails[];
  siteGroupsLoading: boolean;
  siteGroupsErr?: string;
}

const UserPickerInline: React.FC<{
  displayText?: string;
  pickerKey?: string;
  usersService: UsersService;
  onPick: (display: string, key: string) => void;
}> = ({ displayText, pickerKey, usersService, onPick }) => {
  const [q, setQ] = useState('');
  const [results, setResults] = useState<IPeoplePickerResult[]>([]);
  const [searching, setSearching] = useState(false);

  useEffect(() => {
    if (q.trim().length < 2) {
      setResults([]);
      return;
    }
    let cancelled = false;
    const t = window.setTimeout(() => {
      setSearching(true);
      usersService
        .searchUsers(q.trim(), 8)
        .then((r) => {
          if (!cancelled) setResults(r);
        })
        .catch(() => {
          if (!cancelled) setResults([]);
        })
        .finally(() => {
          if (!cancelled) setSearching(false);
        });
    }, 300);
    return () => {
      cancelled = true;
      window.clearTimeout(t);
    };
  }, [q, usersService]);

  return (
    <Stack tokens={{ childrenGap: 4 }} styles={{ root: { minWidth: 280 } }}>
      <TextField
        label="Pessoa do Site (pesquisar)"
        value={q}
        onChange={(_, v) => setQ(v ?? '')}
        description={pickerKey ? `Selecionado: ${displayText ?? pickerKey}` : undefined}
      />
      {searching && <Spinner />}
      {results.map((r, i) => (
        <DefaultButton
          key={`${r.Key}_${i}`}
          text={r.DisplayText}
          onClick={() => {
            onPick(r.DisplayText, r.Key);
            setQ('');
            setResults([]);
          }}
          styles={{ root: { justifyContent: 'flex-start', textAlign: 'left' } }}
        />
      ))}
    </Stack>
  );
};

export function FormManagerPermissionBreakTabContent(props: IFormManagerPermissionBreakTabProps): JSX.Element {
  const {
    primaryListTitle,
    primaryMeta,
    linkedChildForms,
    formManagerForResolve,
    mainAttachmentLibraryEnabled,
    value,
    onChange,
    siteGroups,
    siteGroupsLoading,
    siteGroupsErr,
  } = props;

  const fieldsService = useMemo(() => new FieldsService(), []);
  const usersService = useMemo(() => new UsersService(), []);
  const [linkedMetaById, setLinkedMetaById] = useState<Record<string, IFieldMetadata[]>>({});
  const [linkedMetaLoading, setLinkedMetaLoading] = useState(false);
  const [siteGroupPickFilter, setSiteGroupPickFilter] = useState('');

  useEffect(() => {
    let cancelled = false;
    const load = async (): Promise<void> => {
      setLinkedMetaLoading(true);
      const next: Record<string, IFieldMetadata[]> = {};
      for (let i = 0; i < linkedChildForms.length; i++) {
        const c = linkedChildForms[i];
        const t = c.listTitle.trim();
        if (!t) continue;
        try {
          const f = await fieldsService.getVisibleFields(t);
          if (!cancelled) next[c.id] = f;
        } catch {
          if (!cancelled) next[c.id] = [];
        }
      }
      if (!cancelled) {
        setLinkedMetaById(next);
        setLinkedMetaLoading(false);
      }
    };
    void load();
    return () => {
      cancelled = true;
    };
  }, [linkedChildForms, fieldsService]);

  const siteGroupOpts: IDropdownOption[] = useMemo(() => {
    let visible = filterSiteGroupsByNameQuery(siteGroups, siteGroupPickFilter);
    const need = new Set<number>();
    for (const a of value.assignments ?? []) {
      if (a.kind === 'siteGroup' && typeof a.siteGroupId === 'number' && isFinite(a.siteGroupId)) {
        need.add(a.siteGroupId);
      }
    }
    Array.from(need).forEach((id) => {
      if (!visible.some((g) => g.Id === id)) {
        const g = siteGroups.find((x) => x.Id === id);
        if (g) visible = visible.concat([g]);
      }
    });
    return visible.map((g) => ({ key: String(g.Id), text: g.Title }));
  }, [siteGroups, siteGroupPickFilter, value.assignments]);

  const mainUserFields = useMemo(() => userFieldsFromMeta(primaryMeta), [primaryMeta]);

  const scopeOptionsForAssignment = useCallback(
    (a: IFormPermissionBreakAssignment): IDropdownOption[] => {
      const opts: IDropdownOption[] = [{ key: 'main', text: 'Lista principal' }];
      for (let i = 0; i < linkedChildForms.length; i++) {
        const c = linkedChildForms[i];
        const title = (c.title ?? c.listTitle).trim() || c.listTitle;
        opts.push({ key: `linked:${c.id}`, text: title });
      }
      if (a.kind === 'field' && a.fieldScope === 'linked' && a.linkedFormId) {
        const exists = opts.some((o) => o.key === `linked:${a.linkedFormId}`);
        if (!exists) {
          opts.push({ key: `linked:${a.linkedFormId}`, text: a.linkedFormId });
        }
      }
      return opts;
    },
    [linkedChildForms]
  );

  const fieldOptionsForAssignment = (a: IFormPermissionBreakAssignment): IDropdownOption[] => {
    if (a.kind !== 'field') return [];
    if (a.fieldScope === 'linked') {
      const lid = a.linkedFormId ?? '';
      const meta = linkedMetaById[lid] ?? [];
      return userFieldsFromMeta(meta);
    }
    return mainUserFields;
  };

  const update = useCallback(
    (patch: Partial<IFormManagerPermissionBreakConfig>): void => {
      onChange({ ...value, ...patch });
    },
    [onChange, value]
  );

  const updateTargets = useCallback(
    (patch: Partial<NonNullable<IFormManagerPermissionBreakConfig['targets']>>): void => {
      onChange({
        ...value,
        targets: { ...(value.targets ?? {}), ...patch },
      });
    },
    [onChange, value]
  );

  const setAssignment = useCallback(
    (id: string, patch: Partial<IFormPermissionBreakAssignment>): void => {
      const list = (value.assignments ?? []).map((x) => (x.id === id ? { ...x, ...patch } : x));
      onChange({ ...value, assignments: list });
    },
    [onChange, value]
  );

  const addAssignment = useCallback((): void => {
    const list = (value.assignments ?? []).concat([
      {
        id: newAssignmentId(),
        kind: 'siteGroup',
        roleDefinitionName: 'Leitura',
        siteGroupId: siteGroups[0]?.Id,
        siteGroupTitle: siteGroups[0]?.Title,
      },
    ]);
    onChange({ ...value, assignments: list });
  }, [onChange, siteGroups, value]);

  const removeAssignment = useCallback(
    (id: string): void => {
      onChange({ ...value, assignments: (value.assignments ?? []).filter((x) => x.id !== id) });
    },
    [onChange, value]
  );

  return (
    <Stack tokens={{ childrenGap: 14 }} styles={{ root: { maxWidth: 920 } }}>
      <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
        Quebra de permissões
      </Text>
   
      <Toggle
        label="Ativar"
        checked={value.enabled === true}
        onChange={(_, c) => update({ enabled: c === true })}
      />
      {!value.enabled ? (
        <MessageBar messageBarType={MessageBarType.info}>Ative para configurar alvos e principais.</MessageBar>
      ) : (
        <>
          <Stack horizontal tokens={{ childrenGap: 20 }} wrap>
            <Toggle
              label="Copiar permissões herdadas ao quebrar (primeira vez)"
              checked={value.copyInheritedAssignments === true}
              onChange={(_, c) => update({ copyInheritedAssignments: c === true })}
            />
            <Toggle
              label="Manter autor (Created By) com nível abaixo"
              checked={value.retainAuthor !== false}
              onChange={(_, c) => update({ retainAuthor: c === true })}
            />
            <Dropdown
              label="Nível do autor"
              selectedKey={value.authorRoleDefinitionName ?? 'Contribuir'}
              options={ROLE_OPTIONS}
              onChange={(_, o) =>
                update({ authorRoleDefinitionName: o ? String(o.key) : 'Contribuir' })
              }
              styles={{ root: { width: 200 } }}
            />
          </Stack>

          <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
            Alvos
          </Text>
          <Checkbox
            label={`Item na lista principal (${primaryListTitle || '—'})`}
            checked={value.targets?.mainListItem !== false}
            onChange={(_, c) => updateTargets({ mainListItem: c === true })}
          />
          {linkedChildForms.length > 0 ? (
            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="small">Listas vinculadas (itens filhos)</Text>
              {linkedChildForms.map((c) => {
                const sel = value.targets?.linkedChildFormIds;
                const all = !sel || sel.length === 0;
                const checked = all || (sel && sel.indexOf(c.id) >= 0);
                return (
                  <Checkbox
                    key={c.id}
                    label={(c.title ?? c.listTitle).trim() || c.listTitle}
                    checked={checked}
                    onChange={(_, ch) => {
                      const cur = value.targets?.linkedChildFormIds;
                      if (ch === true) {
                        if (!cur || cur.length === 0) {
                          updateTargets({ linkedChildFormIds: undefined });
                          return;
                        }
                        const next = cur.concat(c.id).filter((x, i, a) => a.indexOf(x) === i);
                        updateTargets({ linkedChildFormIds: next });
                      } else {
                        const base =
                          !cur || cur.length === 0 ? linkedChildForms.map((x) => x.id) : cur.slice();
                        const next = base.filter((x) => x !== c.id);
                        updateTargets({ linkedChildFormIds: next.length ? next : [] });
                      }
                    }}
                  />
                );
              })}
            </Stack>
          ) : null}
          {mainAttachmentLibraryEnabled ? (
            <Checkbox
              label="Ficheiros na biblioteca de anexos (lookup ao item principal)"
              checked={value.targets?.mainAttachmentLibraryFiles === true}
              onChange={(_, c) => updateTargets({ mainAttachmentLibraryFiles: c === true })}
            />
          ) : null}
          {linkedChildForms.some(
            (c) => resolveLinkedChildAttachmentRuntime(c, formManagerForResolve as IFormManagerConfig).kind === 'documentLibrary'
          ) ? (
            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="small">Ficheiros na biblioteca por lista vinculada (lookup à linha filha)</Text>
              {linkedChildForms.map((c) => {
                const r = resolveLinkedChildAttachmentRuntime(c, formManagerForResolve as IFormManagerConfig);
                if (r.kind !== 'documentLibrary') return null;
                const ids = value.targets?.linkedAttachmentLibraryFilesByFormId ?? [];
                const checked = ids.indexOf(c.id) >= 0;
                return (
                  <Checkbox
                    key={`lib_${c.id}`}
                    label={(c.title ?? c.listTitle).trim() || c.listTitle}
                    checked={checked}
                    onChange={(_, ch) => {
                      const prev = value.targets?.linkedAttachmentLibraryFilesByFormId ?? [];
                      const next = ch
                        ? prev.concat(c.id).filter((x, i, a) => a.indexOf(x) === i)
                        : prev.filter((x) => x !== c.id);
                      updateTargets({ linkedAttachmentLibraryFilesByFormId: next.length ? next : undefined });
                    }}
                  />
                );
              })}
            </Stack>
          ) : null}

          <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
            Principais e níveis
          </Text>
          <TextField
            placeholder="Filtrar grupos por nome (dropdown «Grupo»)"
            value={siteGroupPickFilter}
            onChange={(_: unknown, v?: string) => setSiteGroupPickFilter(v ?? '')}
            styles={{ root: { maxWidth: 420 } }}
          />
          {linkedMetaLoading && <Spinner label="A carregar campos das listas vinculadas…" />}
          {(value.assignments ?? []).map((a) => (
            <Stack
              key={a.id}
              horizontal
              verticalAlign="end"
              tokens={{ childrenGap: 8 }}
              wrap
              styles={{ root: { borderLeft: '3px solid #edebe9', paddingLeft: 10 } }}
            >
              <Dropdown
                label="Tipo"
                selectedKey={a.kind}
                options={KIND_OPTIONS}
                onChange={(_, o) => {
                  const k = (o?.key as IFormPermissionBreakAssignment['kind']) ?? 'siteGroup';
                  if (k === 'siteGroup') {
                    const g = siteGroups[0];
                    setAssignment(a.id, {
                      kind: 'siteGroup',
                      siteGroupId: g?.Id,
                      siteGroupTitle: g?.Title,
                      userPickerKey: undefined,
                      userDisplayText: undefined,
                      fieldInternalName: undefined,
                      fieldScope: undefined,
                      linkedFormId: undefined,
                    });
                  } else if (k === 'user') {
                    setAssignment(a.id, {
                      kind: 'user',
                      siteGroupId: undefined,
                      siteGroupTitle: undefined,
                      fieldInternalName: undefined,
                      fieldScope: undefined,
                      linkedFormId: undefined,
                    });
                  } else {
                    setAssignment(a.id, {
                      kind: 'field',
                      fieldScope: 'main',
                      siteGroupId: undefined,
                      siteGroupTitle: undefined,
                      userPickerKey: undefined,
                      userDisplayText: undefined,
                    });
                  }
                }}
                styles={{ root: { width: 160 } }}
              />
              <Dropdown
                label="Nível"
                selectedKey={a.roleDefinitionName}
                options={ROLE_OPTIONS}
                onChange={(_, o) => setAssignment(a.id, { roleDefinitionName: o ? String(o.key) : 'Leitura' })}
                styles={{ root: { width: 160 } }}
              />
              {a.kind === 'siteGroup' ? (
                <Dropdown
                  label="Grupo"
                  selectedKey={a.siteGroupId !== undefined ? String(a.siteGroupId) : undefined}
                  options={siteGroupOpts}
                  onChange={(_, o) => {
                    const id = o ? parseInt(String(o.key), 10) : undefined;
                    const g = siteGroups.find((x) => x.Id === id);
                    setAssignment(a.id, {
                      siteGroupId: id,
                      siteGroupTitle: g?.Title,
                    });
                  }}
                  disabled={siteGroupsLoading || !siteGroupOpts.length}
                  styles={{ root: { minWidth: 220 } }}
                />
              ) : null}
              {a.kind === 'user' ? (
                <UserPickerInline
                  displayText={a.userDisplayText}
                  pickerKey={a.userPickerKey}
                  usersService={usersService}
                  onPick={(display, key) => setAssignment(a.id, { userDisplayText: display, userPickerKey: key })}
                />
              ) : null}
              {a.kind === 'field' ? (
                <>
                  <Dropdown
                    label="Lista"
                    selectedKey={
                      a.fieldScope === 'linked' && a.linkedFormId
                        ? `linked:${a.linkedFormId}`
                        : 'main'
                    }
                    options={scopeOptionsForAssignment(a)}
                    onChange={(_, o) => {
                      const k = String(o?.key ?? 'main');
                      if (k === 'main') {
                        setAssignment(a.id, {
                          fieldScope: 'main',
                          linkedFormId: undefined,
                          fieldInternalName: undefined,
                        });
                      } else if (k.indexOf('linked:') === 0) {
                        const lid = k.slice('linked:'.length);
                        setAssignment(a.id, {
                          fieldScope: 'linked',
                          linkedFormId: lid,
                          fieldInternalName: undefined,
                        });
                      }
                    }}
                    styles={{ root: { minWidth: 200 } }}
                  />
                  <Dropdown
                    label="Campo"
                    selectedKey={a.fieldInternalName}
                    options={fieldOptionsForAssignment(a)}
                    onChange={(_, o) =>
                      setAssignment(a.id, { fieldInternalName: o ? String(o.key) : undefined })
                    }
                    styles={{ root: { minWidth: 220 } }}
                  />
                </>
              ) : null}
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Remover"
                ariaLabel="Remover"
                onClick={() => removeAssignment(a.id)}
              />
            </Stack>
          ))}
          <DefaultButton text="Adicionar principal" onClick={addAssignment} />
          {siteGroupsErr ? (
            <MessageBar messageBarType={MessageBarType.warning}>{siteGroupsErr}</MessageBar>
          ) : null}
        </>
      )}
    </Stack>
  );
}
