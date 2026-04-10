import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Toggle,
  Spinner,
  MessageBar,
  MessageBarType,
  IconButton,
} from '@fluentui/react';
import { FieldsService, ListsService } from '../../../../services';
import type { IFieldMetadata, IListSummary } from '../../../../services';
import type { IFormFieldConfig, IFormLinkedChildFormConfig } from '../../core/config/types/formManager';
import { FORM_ATTACHMENTS_FIELD_INTERNAL } from '../../core/config/types/formManager';
import { newLinkedChildFormConfig } from '../../core/config/utils';

const MAX_LINKED = 10;

const EXCLUDE_INTERNALS = new Set(['Attachments', 'ContentType', 'ContentTypeId']);

const ORDER_DROPDOWN: IDropdownOption[] = [
  { key: '__auto', text: 'Automático (0)' },
  ...Array.from({ length: 31 }, (_, i) => ({
    key: String(i),
    text: String(i),
  })),
];

const MIN_ROWS_DROPDOWN: IDropdownOption[] = Array.from({ length: 21 }, (_, i) => ({
  key: String(i),
  text: i === 0 ? '0 (sem mínimo obrigatório)' : String(i),
}));

const MAX_ROWS_DROPDOWN: IDropdownOption[] = [
  { key: '__none', text: 'Sem limite' },
  ...Array.from({ length: 50 }, (_, i) => ({
    key: String(i + 1),
    text: String(i + 1),
  })),
];

function normListGuid(g: string | undefined): string {
  if (!g) return '';
  return g.replace(/[{}]/g, '').toLowerCase();
}

function isSelectableChildField(m: IFieldMetadata): boolean {
  return !EXCLUDE_INTERNALS.has(m.InternalName) && m.InternalName !== FORM_ATTACHMENTS_FIELD_INTERNAL;
}

function newId(): string {
  return `lcf_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
}

function patchChildById(
  arr: IFormLinkedChildFormConfig[],
  id: string,
  patch: Partial<IFormLinkedChildFormConfig>
): IFormLinkedChildFormConfig[] {
  return arr.map((c) => (c.id === id ? { ...c, ...patch } : c));
}

function setMainStepFieldNames(
  cfg: IFormLinkedChildFormConfig,
  fieldNames: string[]
): IFormLinkedChildFormConfig {
  const steps = (cfg.steps ?? []).map((s) =>
    s.id === 'main' ? { ...s, fieldNames: fieldNames.slice() } : { ...s, fieldNames: s.fieldNames.slice() }
  );
  return { ...cfg, steps };
}

function ensureFieldConfigsForNames(
  cfg: IFormLinkedChildFormConfig,
  names: string[]
): IFormLinkedChildFormConfig {
  const existing = new Map(cfg.fields.map((f) => [f.internalName, f]));
  const fields: IFormFieldConfig[] = [];
  for (let i = 0; i < names.length; i++) {
    const n = names[i];
    fields.push(existing.get(n) ?? { internalName: n });
  }
  return { ...cfg, fields };
}

function listDropdownOptionsWithLegacy(
  base: IDropdownOption[],
  currentTitle: string
): IDropdownOption[] {
  const t = currentTitle.trim();
  if (!t) return base;
  const has = base.some((o) => String(o.key) === t);
  if (has) return base;
  return base.concat([{ key: t, text: `${t} (não encontrada — reverifique)` }]);
}

export interface IFormManagerLinkedChildFormsTabProps {
  primaryListTitle: string;
  linkedChildForms: IFormLinkedChildFormConfig[];
  onLinkedChildFormsChange: (next: IFormLinkedChildFormConfig[]) => void;
}

export function FormManagerLinkedChildFormsTabContent(props: IFormManagerLinkedChildFormsTabProps): JSX.Element {
  const { primaryListTitle, linkedChildForms, onLinkedChildFormsChange } = props;
  const fieldsService = useMemo(() => new FieldsService(), []);
  const listsService = useMemo(() => new ListsService(), []);
  const [siteLists, setSiteLists] = useState<IListSummary[]>([]);
  const [listsLoading, setListsLoading] = useState(false);
  const [childMetaById, setChildMetaById] = useState<Record<string, IFieldMetadata[]>>({});
  const [childMetaLoading, setChildMetaLoading] = useState<Record<string, boolean>>({});

  useEffect(() => {
    setListsLoading(true);
    listsService
      .getLists(false)
      .then((lists) => setSiteLists(lists))
      .catch(() => setSiteLists([]))
      .finally(() => setListsLoading(false));
  }, [listsService]);

  const childListsForDropdown = useMemo(() => {
    const lists = siteLists
      .filter((l) => !l.IsLibrary)
      .slice()
      .sort((a, b) => a.Title.localeCompare(b.Title, 'pt', { sensitivity: 'base' }));
    return lists;
  }, [siteLists]);

  const childListDropdownOptions = useMemo((): IDropdownOption[] => {
    const opts: IDropdownOption[] = [{ key: '', text: '— Selecione uma lista —' }];
    for (let i = 0; i < childListsForDropdown.length; i++) {
      const l = childListsForDropdown[i];
      opts.push({
        key: l.Title,
        text: `${l.Title} (${l.ItemCount} itens)`,
      });
    }
    return opts;
  }, [childListsForDropdown]);

  const primaryListId = useMemo(() => {
    const t = primaryListTitle.trim();
    if (!t || !siteLists.length) return '';
    const hit = siteLists.find((l) => l.Title.trim().toLowerCase() === t.toLowerCase());
    return hit?.Id ? normListGuid(hit.Id) : '';
  }, [primaryListTitle, siteLists]);

  const loadChildMeta = useCallback(
    async (configId: string, listTitle: string): Promise<void> => {
      const lt = listTitle.trim();
      if (!lt) {
        setChildMetaById((prev) => {
          const next = { ...prev };
          delete next[configId];
          return next;
        });
        return;
      }
      setChildMetaLoading((prev) => ({ ...prev, [configId]: true }));
      try {
        const f = await fieldsService.getVisibleFields(lt);
        setChildMetaById((prev) => ({ ...prev, [configId]: f }));
      } catch {
        setChildMetaById((prev) => {
          const next = { ...prev };
          delete next[configId];
          return next;
        });
      } finally {
        setChildMetaLoading((prev) => ({ ...prev, [configId]: false }));
      }
    },
    [fieldsService]
  );

  const addLinked = (): void => {
    if (linkedChildForms.length >= MAX_LINKED) return;
    onLinkedChildFormsChange([...linkedChildForms, newLinkedChildFormConfig(newId())]);
  };

  const removeLinked = (id: string): void => {
    onLinkedChildFormsChange(linkedChildForms.filter((c) => c.id !== id));
    setChildMetaById((prev) => {
      const next = { ...prev };
      delete next[id];
      return next;
    });
  };

  const moveLinked = (from: number, to: number): void => {
    if (to < 0 || to >= linkedChildForms.length) return;
    const next = linkedChildForms.slice();
    const [m] = next.splice(from, 1);
    next.splice(to, 0, m);
    onLinkedChildFormsChange(next);
  };

  return (
    <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Listas secundárias com Lookup para a lista principal. A ordem na etapa «Geral» define cada mini-formulário.
        Em «Estrutura», escolha em que etapa do passador cada bloco aparece.
      </Text>
      {!primaryListTitle.trim() && (
        <MessageBar messageBarType={MessageBarType.warning}>
          Defina primeiro a lista principal na configuração da vista (origem de dados).
        </MessageBar>
      )}
      {listsLoading && <Spinner label="A carregar listas do site…" />}
      {primaryListTitle.trim() && !listsLoading && !primaryListId && (
        <MessageBar messageBarType={MessageBarType.severeWarning}>
          Não foi encontrada uma lista com o título «{primaryListTitle.trim()}» no site. Os Lookups filtrados
          podem ficar incompletos.
        </MessageBar>
      )}
      <PrimaryButton text="Adicionar lista vinculada" onClick={addLinked} disabled={linkedChildForms.length >= MAX_LINKED} />
      {linkedChildForms.map((cfg, index) => {
        const meta = childMetaById[cfg.id] ?? [];
        const loading = childMetaLoading[cfg.id] === true;
        const listOpts = listDropdownOptionsWithLegacy(childListDropdownOptions, cfg.listTitle);
        const mainNames =
          cfg.steps?.find((s) => s.id === 'main')?.fieldNames?.slice() ?? [];
        const lookupOpts: IDropdownOption[] = meta
          .filter((m) => m.MappedType === 'lookup' && m.LookupList)
          .filter((m) => {
            if (!primaryListId) return true;
            return normListGuid(m.LookupList) === primaryListId;
          })
          .map((m) => ({ key: m.InternalName, text: `${m.Title} (${m.InternalName})` }));
        const lookupDropdownOptions: IDropdownOption[] = lookupOpts.length
          ? [{ key: '', text: '— Selecione o campo Lookup —' }, ...lookupOpts]
          : [
              {
                key: '',
                text: cfg.listTitle.trim()
                  ? loading
                    ? 'A carregar campos…'
                    : '— Nenhum campo Lookup para a lista principal —'
                  : '— Escolha primeiro a lista filha —',
              },
            ];
        const availableToAdd = meta.filter(
          (m) =>
            isSelectableChildField(m) &&
            m.InternalName !== cfg.parentLookupFieldInternalName.trim() &&
            mainNames.indexOf(m.InternalName) === -1
        );
        const addOptions: IDropdownOption[] = availableToAdd.map((m) => ({
          key: m.InternalName,
          text: `${m.Title} (${m.InternalName})`,
        }));

        return (
          <Stack
            key={cfg.id}
            tokens={{ childrenGap: 10 }}
            styles={{
              root: {
                border: '1px solid #edebe9',
                borderRadius: 4,
                padding: 12,
                background: '#ffffff',
              },
            }}
          >
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
              Lista vinculada {index + 1}: {cfg.listTitle.trim() || '(sem título)'}
            </Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <IconButton
                iconProps={{ iconName: 'Up' }}
                title="Mover bloco para cima"
                disabled={index === 0}
                onClick={() => moveLinked(index, index - 1)}
              />
              <IconButton
                iconProps={{ iconName: 'Down' }}
                title="Mover bloco para baixo"
                disabled={index === linkedChildForms.length - 1}
                onClick={() => moveLinked(index, index + 1)}
              />
              <DefaultButton text="Remover este bloco" onClick={() => removeLinked(cfg.id)} />
            </Stack>
            <Dropdown
              label="Lista filha (SharePoint)"
              options={listOpts}
              selectedKey={cfg.listTitle.trim() || ''}
              onChange={(_, o) => {
                if (!o) return;
                const title = String(o.key);
                onLinkedChildFormsChange(
                  patchChildById(linkedChildForms, cfg.id, {
                    listTitle: title,
                    parentLookupFieldInternalName: '',
                  })
                );
                void loadChildMeta(cfg.id, title);
              }}
              disabled={listsLoading}
            />
            {loading && <Spinner label="A carregar campos da lista…" />}
            <Dropdown
              label="Campo Lookup para a lista principal"
              options={lookupDropdownOptions}
              selectedKey={cfg.parentLookupFieldInternalName.trim() || ''}
              onChange={(_, o) =>
                o &&
                onLinkedChildFormsChange(
                  patchChildById(linkedChildForms, cfg.id, {
                    parentLookupFieldInternalName: String(o.key),
                  })
                )
              }
              disabled={!cfg.listTitle.trim() || loading || !lookupOpts.length}
            />
            {cfg.title?.trim() && (
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Rótulo JSON legado: «{cfg.title}»
                </Text>
                <DefaultButton
                  text="Usar só o nome da lista"
                  onClick={() =>
                    onLinkedChildFormsChange(
                      patchChildById(linkedChildForms, cfg.id, { title: undefined })
                    )
                  }
                />
              </Stack>
            )}
            <Stack horizontal tokens={{ childrenGap: 12 }} wrap>
              <Dropdown
                label="Ordem de exibição"
                options={ORDER_DROPDOWN}
                selectedKey={cfg.order === undefined ? '__auto' : String(cfg.order)}
                onChange={(_, o) => {
                  if (!o) return;
                  const k = String(o.key);
                  onLinkedChildFormsChange(
                    patchChildById(linkedChildForms, cfg.id, {
                      order: k === '__auto' ? undefined : parseInt(k, 10),
                    })
                  );
                }}
                styles={{ root: { minWidth: 200 } }}
              />
              <Dropdown
                label="Mínimo de linhas"
                options={MIN_ROWS_DROPDOWN}
                selectedKey={String(cfg.minRows ?? 0)}
                onChange={(_, o) =>
                  o &&
                  onLinkedChildFormsChange(
                    patchChildById(linkedChildForms, cfg.id, {
                      minRows: parseInt(String(o.key), 10),
                    })
                  )
                }
                styles={{ root: { minWidth: 200 } }}
              />
              <Dropdown
                label="Máximo de linhas"
                options={MAX_ROWS_DROPDOWN}
                selectedKey={cfg.maxRows === undefined ? '__none' : String(cfg.maxRows)}
                onChange={(_, o) => {
                  if (!o) return;
                  const k = String(o.key);
                  onLinkedChildFormsChange(
                    patchChildById(linkedChildForms, cfg.id, {
                      maxRows: k === '__none' ? undefined : parseInt(k, 10),
                    })
                  );
                }}
                styles={{ root: { minWidth: 200 } }}
              />
            </Stack>
            <Toggle
              label="Secção recolhida por defeito (reservado)"
              checked={cfg.collapsedDefault === true}
              onChange={(_, c) =>
                onLinkedChildFormsChange(
                  patchChildById(linkedChildForms, cfg.id, { collapsedDefault: !!c })
                )
              }
            />
            <Text variant="small" styles={{ root: { fontWeight: 600, marginTop: 8 } }}>
              Campos na etapa «Geral» (ordem)
            </Text>
            <Stack styles={{ root: { border: '1px solid #edebe9', padding: 8, borderRadius: 4 } }} tokens={{ childrenGap: 6 }}>
              {mainNames.length === 0 && (
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Nenhum campo. Adicione abaixo.
                </Text>
              )}
              {mainNames.map((name, ni) => {
                const m = meta.find((x) => x.InternalName === name);
                return (
                  <Stack key={name} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                    <Text variant="small" styles={{ root: { minWidth: 180 } }}>
                      {m ? `${m.Title} (${name})` : name}
                    </Text>
                    <IconButton
                      iconProps={{ iconName: 'Up' }}
                      title="Subir"
                      disabled={ni === 0}
                      onClick={() => {
                        const fn = mainNames.slice();
                        const t = fn[ni - 1];
                        fn[ni - 1] = fn[ni];
                        fn[ni] = t;
                        let next = setMainStepFieldNames(cfg, fn);
                        next = ensureFieldConfigsForNames(next, fn);
                        onLinkedChildFormsChange(
                          linkedChildForms.map((c) => (c.id === cfg.id ? next : c))
                        );
                      }}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Down' }}
                      title="Descer"
                      disabled={ni === mainNames.length - 1}
                      onClick={() => {
                        const fn = mainNames.slice();
                        const t = fn[ni + 1];
                        fn[ni + 1] = fn[ni];
                        fn[ni] = t;
                        let next = setMainStepFieldNames(cfg, fn);
                        next = ensureFieldConfigsForNames(next, fn);
                        onLinkedChildFormsChange(
                          linkedChildForms.map((c) => (c.id === cfg.id ? next : c))
                        );
                      }}
                    />
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Remover"
                      onClick={() => {
                        const fn = mainNames.filter((_, j) => j !== ni);
                        let next = setMainStepFieldNames(cfg, fn);
                        next = {
                          ...next,
                          fields: next.fields.filter((f) => fn.indexOf(f.internalName) !== -1),
                        };
                        onLinkedChildFormsChange(
                          linkedChildForms.map((c) => (c.id === cfg.id ? next : c))
                        );
                      }}
                    />
                  </Stack>
                );
              })}
            </Stack>
            <Dropdown
              placeholder="Adicionar campo à etapa Geral…"
              options={addOptions.length ? addOptions : [{ key: '__none', text: '— sem campos disponíveis —' }]}
              selectedKey=""
              onChange={(_, o) => {
                if (!o || o.key === '__none') return;
                const key = String(o.key);
                if (mainNames.indexOf(key) !== -1) return;
                const fn = mainNames.concat([key]);
                let next = setMainStepFieldNames(cfg, fn);
                next = ensureFieldConfigsForNames(next, fn);
                onLinkedChildFormsChange(
                  linkedChildForms.map((c) => (c.id === cfg.id ? next : c))
                );
              }}
            />
          </Stack>
        );
      })}
    </Stack>
  );
}
