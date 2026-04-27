import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { Stack, Text, Dropdown, IDropdownOption, DefaultButton, Spinner, SpinnerSize } from '@fluentui/react';
import { FieldsService, ListsService } from '../../../../services';
import type { IListPageLinkedListBinding } from '../../core/config/types';
import { discoverListsWithLookupToMain, type IDiscoveredLinkedList } from '../../core/listPage/linkedListDiscovery';

export interface IListPageBlockLinkedBindingEditorProps {
  layoutPanelOpen: boolean;
  mainListTitle: string;
  binding: IListPageLinkedListBinding | undefined;
  onBindingChange: (next: IListPageLinkedListBinding | undefined) => void;
}

export const ListPageBlockLinkedBindingEditor: React.FC<IListPageBlockLinkedBindingEditorProps> = ({
  layoutPanelOpen,
  mainListTitle,
  binding,
  onBindingChange,
}) => {
  const listsService = useMemo(() => new ListsService(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);
  const [rows, setRows] = useState<IDiscoveredLinkedList[]>([]);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState<string | undefined>(undefined);

  const load = useCallback(async (): Promise<void> => {
    const t = mainListTitle.trim();
    if (!t) {
      setRows([]);
      return;
    }
    setLoading(true);
    setErr(undefined);
    try {
      const r = await discoverListsWithLookupToMain(t, listsService, fieldsService);
      setRows(r);
    } catch (e) {
      setErr(e instanceof Error ? e.message : String(e));
      setRows([]);
    } finally {
      setLoading(false);
    }
  }, [mainListTitle, listsService, fieldsService]);

  useEffect(() => {
    if (!layoutPanelOpen || !mainListTitle.trim()) {
      setRows([]);
      setErr(undefined);
      return;
    }
    let cancelled = false;
    void (async (): Promise<void> => {
      setLoading(true);
      setErr(undefined);
      try {
        const r = await discoverListsWithLookupToMain(mainListTitle.trim(), listsService, fieldsService);
        if (!cancelled) setRows(r);
      } catch (e) {
        if (!cancelled) {
          setErr(e instanceof Error ? e.message : String(e));
          setRows([]);
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [layoutPanelOpen, mainListTitle, listsService, fieldsService]);

  const listOptions = useMemo((): IDropdownOption[] => {
    const opts: IDropdownOption[] = [{ key: '__main', text: 'Lista principal (vista)' }];
    const seen = new Set<string>();
    for (let i = 0; i < rows.length; i++) {
      const title = rows[i].listTitle;
      if (seen.has(title)) continue;
      seen.add(title);
      opts.push({ key: title, text: title });
    }
    const cur = binding?.listTitle?.trim();
    if (cur && !seen.has(cur)) {
      opts.push({ key: cur, text: `${cur} (configurado)` });
    }
    return opts;
  }, [rows, binding?.listTitle]);

  const selectedListKey = binding?.listTitle?.trim() ? binding.listTitle.trim() : '__main';

  const lookupOptions = useMemo((): IDropdownOption[] => {
    const lt = binding?.listTitle?.trim();
    if (!lt) return [];
    const row = rows.find((r) => r.listTitle === lt);
    const base = row?.lookupFields ?? [];
    const opts = base.map((f) => ({ key: f.internalName, text: `${f.title} (${f.internalName})` }));
    const saved = binding?.parentLookupFieldInternalName?.trim();
    if (saved && !opts.some((o) => String(o.key) === saved)) {
      return opts.concat([{ key: saved, text: `${saved} (guardado)` }]);
    }
    return opts;
  }, [rows, binding?.listTitle, binding?.parentLookupFieldInternalName]);

  const onListChange = (_: unknown, opt?: IDropdownOption): void => {
    if (!opt) return;
    const k = String(opt.key);
    if (k === '__main') {
      onBindingChange(undefined);
      return;
    }
    const row = rows.find((r) => r.listTitle === k);
    const firstLk = row?.lookupFields[0]?.internalName ?? '';
    onBindingChange({ listTitle: k, parentLookupFieldInternalName: firstLk });
  };

  const onLookupChange = (_: unknown, opt?: IDropdownOption): void => {
    if (!opt || !binding?.listTitle?.trim()) return;
    onBindingChange({
      listTitle: binding.listTitle.trim(),
      parentLookupFieldInternalName: String(opt.key),
    });
  };

  return (
    <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 6, paddingTop: 8, borderTop: '1px solid #edebe9' } }}>
      <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
        Lista vinculada (Lookup para a principal)
      </Text>
    
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
        <DefaultButton text="Atualizar listas" onClick={() => void load()} disabled={loading || !mainListTitle.trim()} />
        {loading ? <Spinner size={SpinnerSize.small} /> : null}
      </Stack>
      {err ? (
        <Text variant="small" styles={{ root: { color: '#a4262c' } }}>
          {err}
        </Text>
      ) : null}
      <Dropdown
        label="Origem dos dados"
        selectedKey={selectedListKey === '' ? '__main' : selectedListKey}
        options={listOptions}
        onChange={onListChange}
        disabled={!mainListTitle.trim()}
      />
      {binding?.listTitle?.trim() ? (
        <Dropdown
          label="Campo Lookup (para a lista principal)"
          selectedKey={binding.parentLookupFieldInternalName?.trim() || undefined}
          options={lookupOptions}
          onChange={onLookupChange}
          disabled={lookupOptions.length === 0}
        />
      ) : null}
    </Stack>
  );
};
