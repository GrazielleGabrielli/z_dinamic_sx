import * as React from 'react';
import { useEffect, useMemo, useState, useCallback } from 'react';
import {
  Stack,
  Text,
  Spinner,
  MessageBar,
  MessageBarType,
  DetailsList,
  SelectionMode,
  Dropdown,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import { ItemsService } from '../../../../services';
import { normalizeItemsQuerySelectExpand } from '../../../../services/items/ItemsService';
import { readListItemId } from '../../../../services/items/listItemId';
import type { TFormManagerBrowseLayoutControlKind } from '../../core/config/types/formManager';
import { TableCardsLayoutToggle, type TTableCardsLayoutKind } from '../DataTable/TableCardsLayoutToggle';

function formatCellValue(item: Record<string, unknown>, m: IFieldMetadata): string {
  const v = item[m.InternalName];
  if (v == null || v === '') return '';
  if (typeof v === 'object' && v !== null) {
    if (!Array.isArray(v) && 'Title' in (v as Record<string, unknown>)) {
      return String((v as { Title?: string }).Title ?? '').trim();
    }
    if (Array.isArray(v)) {
      return v
        .map((x) =>
          x !== null && typeof x === 'object' && 'Title' in (x as Record<string, unknown>)
            ? String((x as { Title?: string }).Title ?? '').trim()
            : String(x ?? '')
        )
        .filter(Boolean)
        .join('; ');
    }
    return '';
  }
  if (m.MappedType === 'boolean') return v ? 'Sim' : 'Não';
  const s = String(v);
  if (m.MappedType === 'multiline' && s.length > 120) return `${s.slice(0, 117)}…`;
  return s;
}

export interface IFormManagerBrowseListProps {
  listTitle: string;
  listWebServerRelativeUrl?: string;
  columns: IFieldMetadata[];
  select: string[];
  expand: string[];
  fieldMetadata: IFieldMetadata[];
  itemsService: ItemsService;
  defaultLayoutKind: TTableCardsLayoutKind;
  layoutControl?: TFormManagerBrowseLayoutControlKind;
  refreshSignal?: number;
  selectedItemId?: number;
  onSelectRow: (id: number) => void;
}

export const FormManagerBrowseList: React.FC<IFormManagerBrowseListProps> = ({
  listTitle,
  listWebServerRelativeUrl,
  columns,
  select,
  expand,
  fieldMetadata,
  itemsService,
  defaultLayoutKind,
  layoutControl = 'segmented',
  refreshSignal,
  selectedItemId,
  onSelectRow,
}) => {
  const [layoutKind, setLayoutKind] = useState<TTableCardsLayoutKind>(defaultLayoutKind);
  const [items, setItems] = useState<Record<string, unknown>[]>([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState<string | undefined>(undefined);

  useEffect(() => {
    setLayoutKind(defaultLayoutKind);
  }, [defaultLayoutKind]);

  const { select: normSelect, expand: normExpand } = useMemo(
    () => normalizeItemsQuerySelectExpand(select, expand, fieldMetadata),
    [select, expand, fieldMetadata]
  );

  const reload = useCallback(() => {
    const t = listTitle.trim();
    if (!t) {
      setItems([]);
      setLoading(false);
      return;
    }
    setLoading(true);
    setErr(undefined);
    void (async (): Promise<void> => {
      try {
        const data = await itemsService.getItems(t, {
          select: normSelect,
          expand: normExpand.length ? normExpand : undefined,
          orderBy: { field: 'Id', ascending: false },
          top: 200,
          fieldMetadata,
          ...(listWebServerRelativeUrl ? { webServerRelativeUrl: listWebServerRelativeUrl } : {}),
        });
        setItems(Array.isArray(data) ? data : []);
      } catch (e) {
        setErr(e instanceof Error ? e.message : String(e));
        setItems([]);
      } finally {
        setLoading(false);
      }
    })();
  }, [listTitle, normSelect, normExpand, fieldMetadata, itemsService, listWebServerRelativeUrl]);

  useEffect(() => {
    reload();
  }, [reload, refreshSignal]);

  const detailColumns = useMemo(() => {
    const idCol = {
      key: 'Id',
      name: 'ID',
      fieldName: 'Id',
      minWidth: 52,
      maxWidth: 72,
      isResizable: true,
    };
    const rest = columns.map((m) => ({
      key: m.InternalName,
      name: m.Title,
      minWidth: 96,
      isResizable: true,
      onRender: (row: Record<string, unknown>) => <span>{formatCellValue(row, m)}</span>,
    }));
    return [idCol, ...rest];
  }, [columns]);

  const pickRow = useCallback(
    (row: Record<string, unknown> | undefined) => {
      if (!row) return;
      const id = readListItemId(row);
      if (id != null) onSelectRow(id);
    },
    [onSelectRow]
  );

  const layoutPicker =
    layoutControl === 'compactDropdown' ? (
      <Dropdown
        selectedKey={layoutKind}
        onChange={(_, o) => o && setLayoutKind(String(o.key) as TTableCardsLayoutKind)}
        options={[
          { key: 'table', text: 'Tabela' },
          { key: 'cards', text: 'Cartões' },
        ]}
        styles={{ root: { width: 140 } }}
      />
    ) : (
      <TableCardsLayoutToggle value={layoutKind} onChange={setLayoutKind} />
    );

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { maxWidth: '100%' } }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap tokens={{ childrenGap: 8 }}>
        <Text variant="smallPlus" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
          Registos da lista
        </Text>
        {layoutPicker}
      </Stack>
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      {loading && <Spinner label="A carregar itens…" />}
      {!loading && !err && items.length === 0 && (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Não há itens para mostrar.
        </Text>
      )}
      {!loading && !err && layoutKind === 'table' && items.length > 0 && (
        <DetailsList
          items={items}
          columns={detailColumns}
          selectionMode={SelectionMode.none}
          isHeaderVisible
          compact
          onActiveItemChanged={pickRow}
        />
      )}
      {!loading && !err && layoutKind === 'cards' && items.length > 0 && (
        <div
          style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fill, minmax(240px, 1fr))',
            gap: 12,
          }}
        >
          {items.map((row, i) => {
            const id = readListItemId(row);
            const sel = selectedItemId != null && id === selectedItemId;
            return (
              <button
                key={id != null ? String(id) : `row-${i}`}
                type="button"
                onClick={() => pickRow(row)}
                style={{
                  textAlign: 'left',
                  padding: 14,
                  borderRadius: 8,
                  border: `1px solid ${sel ? '#0078d4' : '#edebe9'}`,
                  background: sel ? '#f3f9ff' : '#ffffff',
                  boxShadow: sel ? '0 0 0 1px rgba(0,120,212,0.35)' : '0 1px 2px rgba(0,0,0,0.04)',
                  cursor: 'pointer',
                }}
              >
                <Text variant="small" styles={{ root: { fontWeight: 700, color: '#323130', display: 'block' } }}>
                  #{id ?? '—'}
                </Text>
                {columns.slice(0, 5).map((m) => (
                  <Text
                    key={m.InternalName}
                    variant="small"
                    styles={{ root: { color: '#605e5c', marginTop: 6, lineHeight: 1.35 } }}
                  >
                    <span style={{ fontWeight: 600, color: '#323130' }}>{m.Title}: </span>
                    {formatCellValue(row, m) || '—'}
                  </Text>
                ))}
              </button>
            );
          })}
        </div>
      )}
    </Stack>
  );
};
