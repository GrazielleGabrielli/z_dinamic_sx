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
  Icon,
  Dropdown,
  type IColumn,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import { ItemsService, normalizeItemsQuerySelectExpand } from '../../../../services/items/ItemsService';
import { readListItemId } from '../../../../services/items/listItemId';
import type {
  TFormManagerBrowseLayoutKind,
  TFormManagerBrowseLayoutControlKind,
} from '../../core/config/types/formManager';

function formatCellValue(item: Record<string, unknown>, m: IFieldMetadata): string {
  const v = item[m.InternalName];
  if (v == null || v === '') return '';
  if (typeof v === 'object' && v !== null) {
    if (!Array.isArray(v) && 'Title' in (v as object)) {
      return String((v as { Title?: string }).Title ?? '').trim();
    }
    if (Array.isArray(v)) {
      return v
        .map((x) =>
          x !== null && typeof x === 'object' && 'Title' in (x as object)
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

function BrowseLayoutSegmented(props: {
  value: TFormManagerBrowseLayoutKind;
  onChange: (v: TFormManagerBrowseLayoutKind) => void;
}): JSX.Element {
  const { value, onChange } = props;
  const track: React.CSSProperties = {
    display: 'inline-flex',
    alignItems: 'center',
    background: '#e8e8e8',
    borderRadius: 9,
    padding: 3,
    gap: 2,
  };
  const btn = (active: boolean): React.CSSProperties => ({
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 36,
    height: 32,
    border: 'none',
    borderRadius: 7,
    cursor: 'pointer',
    background: active ? '#ffffff' : 'transparent',
    boxShadow: active ? '0 1px 2px rgba(0,0,0,0.06)' : undefined,
    lineHeight: 0,
    color: '#323130',
  });
  return (
    <div style={track} role="group" aria-label="Modo de visualização">
      <button
        type="button"
        aria-pressed={value === 'table'}
        title="Tabela"
        style={btn(value === 'table')}
        onClick={() => onChange('table')}
      >
        <Icon iconName="BulletedList" styles={{ root: { fontSize: 16 } }} />
      </button>
      <button
        type="button"
        aria-pressed={value === 'cards'}
        title="Cartões"
        style={btn(value === 'cards')}
        onClick={() => onChange('cards')}
      >
        <Icon iconName="Tiles" styles={{ root: { fontSize: 16 } }} />
      </button>
    </div>
  );
}

export interface IFormManagerBrowseListProps {
  listTitle: string;
  listWebServerRelativeUrl?: string;
  columns: IFieldMetadata[];
  select: string[];
  expand: string[];
  fieldMetadata: IFieldMetadata[];
  itemsService: ItemsService;
  defaultLayoutKind: TFormManagerBrowseLayoutKind;
  layoutControl?: TFormManagerBrowseLayoutControlKind;
  refreshSignal: number;
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
  const [layoutKind, setLayoutKind] = useState<TFormManagerBrowseLayoutKind>(defaultLayoutKind);
  const [items, setItems] = useState<Record<string, unknown>[]>([]);
  const [loading, setLoading] = useState(true);
  const [err, setErr] = useState<string | undefined>(undefined);

  useEffect(() => {
    setLayoutKind(defaultLayoutKind);
  }, [defaultLayoutKind]);

  const { normSelect, normExpand } = useMemo(() => {
    const n = normalizeItemsQuerySelectExpand(select, expand, fieldMetadata);
    return { normSelect: n.select, normExpand: n.expand };
  }, [select, expand, fieldMetadata]);

  const reload = useCallback((): void => {
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
        const data = await itemsService.getItems<Record<string, unknown>>(t, {
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
  }, [
    listTitle,
    normSelect,
    normExpand,
    fieldMetadata,
    itemsService,
    listWebServerRelativeUrl,
  ]);

  useEffect(() => {
    reload();
  }, [reload, refreshSignal]);

  const detailColumns: IColumn[] = useMemo(() => {
    const idCol: IColumn = {
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
      onRender: (row: Record<string, unknown>) => (
        <span>{formatCellValue(row, m)}</span>
      ),
    }));
    return [idCol, ...rest];
  }, [columns]);

  const pickRow = useCallback(
    (row: Record<string, unknown> | null): void => {
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
        onChange={(_, o) => o && setLayoutKind(String(o.key) as TFormManagerBrowseLayoutKind)}
        options={[
          { key: 'table', text: 'Tabela' },
          { key: 'cards', text: 'Cartões' },
        ]}
        styles={{ root: { width: 140 } }}
      />
    ) : (
      <BrowseLayoutSegmented value={layoutKind} onChange={setLayoutKind} />
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
