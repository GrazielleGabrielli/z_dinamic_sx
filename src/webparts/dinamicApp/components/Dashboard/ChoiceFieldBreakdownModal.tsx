import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Modal,
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  ChoiceGroup,
  IChoiceGroupOption,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import { FieldsService, ItemsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import type { IDashboardCardConfig, IChartSeriesConfig, TFilterOperator } from '../../core/config/types';
import { getDefaultDashboardCardStyle } from '../../core/dashboard/utils';

function padHex2(n: number): string {
  const s = n.toString(16);
  return s.length >= 2 ? s.slice(-2) : `0${s}`;
}

function randomHexColor(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.getRandomValues === 'function') {
    const buf = new Uint8Array(3);
    crypto.getRandomValues(buf);
    return `#${padHex2(buf[0])}${padHex2(buf[1])}${padHex2(buf[2])}`;
  }
  const n = Math.floor(Math.random() * 0xffffff);
  let h = n.toString(16);
  while (h.length < 6) h = `0${h}`;
  return `#${h}`;
}

function pickDistinctRandomHexColors(count: number): string[] {
  const used = new Set<string>();
  const out: string[] = [];
  for (let i = 0; i < count; i++) {
    let c = randomHexColor();
    let tries = 0;
    while (used.has(c) && tries < 80) {
      c = randomHexColor();
      tries++;
    }
    used.add(c);
    out.push(c);
  }
  return out;
}

const MERGE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'append', text: 'Acrescentar aos existentes' },
  { key: 'replace', text: 'Substituir todos' },
];

function isChoiceLike(f: IFieldMetadata): boolean {
  return f.MappedType === 'choice' || f.MappedType === 'multichoice';
}

function isLookupSingle(f: IFieldMetadata): boolean {
  return f.MappedType === 'lookup' && Boolean(f.LookupList);
}

function filterOperatorForField(f: IFieldMetadata): TFilterOperator {
  return f.MappedType === 'multichoice' ? 'contains' : 'eq';
}

function choiceSlug(v: string, index: number): string {
  const base = v.replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '').slice(0, 28);
  return `${base || 'opt'}_${index}`;
}

function buildCardsFromChoiceField(field: IFieldMetadata, baseTime: number): IDashboardCardConfig[] {
  const choices = field.Choices ?? [];
  const op = filterOperatorForField(field);
  const baseStyle = getDefaultDashboardCardStyle();
  const colors = pickDistinctRandomHexColors(choices.length);
  return choices.map((value, i) => ({
    id: `card_${field.InternalName}_${choiceSlug(value, i)}_${baseTime}`,
    title: value,
    subtitle: '',
    aggregate: 'count',
    filters: [{ field: field.InternalName, operator: op, value }],
    emptyValueText: 'Nenhum item',
    errorText: 'Erro ao carregar',
    loadingText: 'Carregando...',
    style: { ...baseStyle, borderColor: colors[i], valueColor: colors[i] },
  }));
}

function buildSeriesFromChoiceField(field: IFieldMetadata, baseTime: number): IChartSeriesConfig[] {
  const choices = field.Choices ?? [];
  const op = filterOperatorForField(field);
  const colors = pickDistinctRandomHexColors(choices.length);
  return choices.map((value, i) => ({
    id: `series_${field.InternalName}_${choiceSlug(value, i)}_${baseTime}`,
    label: value,
    aggregate: 'count',
    filters: [{ field: field.InternalName, operator: op, value }],
    color: colors[i],
  }));
}

function buildCardsFromLookupField(
  field: IFieldMetadata,
  rows: { id: number; label: string }[],
  baseTime: number
): IDashboardCardConfig[] {
  const baseStyle = getDefaultDashboardCardStyle();
  const colors = pickDistinctRandomHexColors(rows.length);
  return rows.map((row, i) => ({
    id: `card_${field.InternalName}_${row.id}_${choiceSlug(row.label, i)}_${baseTime}`,
    title: row.label,
    subtitle: '',
    aggregate: 'count',
    filters: [{ field: field.InternalName, operator: 'eq', value: String(row.id) }],
    emptyValueText: 'Nenhum item',
    errorText: 'Erro ao carregar',
    loadingText: 'Carregando...',
    style: { ...baseStyle, borderColor: colors[i], valueColor: colors[i] },
  }));
}

function buildSeriesFromLookupField(
  field: IFieldMetadata,
  rows: { id: number; label: string }[],
  baseTime: number
): IChartSeriesConfig[] {
  const colors = pickDistinctRandomHexColors(rows.length);
  return rows.map((row, i) => ({
    id: `series_${field.InternalName}_${row.id}_${choiceSlug(row.label, i)}_${baseTime}`,
    label: row.label,
    aggregate: 'count',
    filters: [{ field: field.InternalName, operator: 'eq', value: String(row.id) }],
    color: colors[i],
  }));
}

export type TChoiceBreakdownTarget = 'cards' | 'series';

export interface IChoiceFieldBreakdownModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  listTitle: string;
  listWebServerRelativeUrl?: string;
  target: TChoiceBreakdownTarget;
  onApply: (items: IDashboardCardConfig[] | IChartSeriesConfig[], mergeMode: 'append' | 'replace') => void;
}

export const ChoiceFieldBreakdownModal: React.FC<IChoiceFieldBreakdownModalProps> = ({
  isOpen,
  onDismiss,
  listTitle,
  listWebServerRelativeUrl,
  target,
  onApply,
}) => {
  const lw = listWebServerRelativeUrl?.trim() || undefined;
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [fields, setFields] = useState<IFieldMetadata[]>([]);
  const [selectedKey, setSelectedKey] = useState<string>('');
  const [mergeKey, setMergeKey] = useState<'append' | 'replace'>('append');
  const [lookupRows, setLookupRows] = useState<{ id: number; label: string }[]>([]);
  const [lookupLoading, setLookupLoading] = useState(false);

  const fieldsService = useMemo(() => new FieldsService(), []);
  const itemsService = useMemo(() => new ItemsService(), []);

  useEffect(() => {
    if (!isOpen || !listTitle.trim()) {
      setFields([]);
      setSelectedKey('');
      setError(undefined);
      return;
    }
    setLoading(true);
    setError(undefined);
    fieldsService
      .getVisibleFields(listTitle.trim(), lw)
      .then((all) => {
        const pick = all.filter((f) => isChoiceLike(f) || isLookupSingle(f));
        setFields(pick);
        setSelectedKey(pick[0]?.InternalName ?? '');
        setLoading(false);
      })
      .catch((e) => {
        setError(e instanceof Error ? e.message : String(e));
        setFields([]);
        setLoading(false);
      });
  }, [isOpen, listTitle, lw]);

  useEffect(() => {
    let cancelled = false;
    if (!selectedKey) {
      setLookupRows([]);
      setLookupLoading(false);
      return (): void => {
        cancelled = true;
      };
    }
    let sel: IFieldMetadata | undefined;
    for (let i = 0; i < fields.length; i++) {
      if (fields[i].InternalName === selectedKey) sel = fields[i];
    }
    if (!sel || !isLookupSingle(sel)) {
      setLookupRows([]);
      setLookupLoading(false);
      return (): void => {
        cancelled = true;
      };
    }
    const m = sel;
    const lf = m.LookupField || 'Title';
    const listGuid = String(m.LookupList ?? '');
    setLookupLoading(true);
    setLookupRows([]);
    itemsService
      .getItems<Record<string, unknown>>(listGuid, {
        select: ['Id', lf],
        top: 500,
        webServerRelativeUrl: lw,
      })
      .then((rows) => {
        if (cancelled) return;
        const out: { id: number; label: string }[] = [];
        for (let i = 0; i < rows.length; i++) {
          const r = rows[i];
          const rawId = r.Id ?? r.id;
          const id = typeof rawId === 'number' ? rawId : parseInt(String(rawId), 10);
          if (Number.isNaN(id)) continue;
          const lv = r[lf];
          const label = lv !== undefined && lv !== null ? String(lv) : `#${id}`;
          out.push({ id, label });
        }
        out.sort((a, b) => a.label.localeCompare(b.label, undefined, { sensitivity: 'base' }));
        setLookupRows(out);
        setLookupLoading(false);
      })
      .catch(() => {
        if (!cancelled) {
          setLookupRows([]);
          setLookupLoading(false);
        }
      });
    return (): void => {
      cancelled = true;
    };
  }, [selectedKey, fields, lw, itemsService]);

  function fieldKind(f: IFieldMetadata | undefined): 'choice' | 'lookup' | '' {
    if (!f) return '';
    if (isLookupSingle(f)) return 'lookup';
    if (isChoiceLike(f)) return 'choice';
    return '';
  }

  const dropdownOptions: IDropdownOption[] = useMemo(
    () =>
      fields.map((f) => {
        const kind =
          f.MappedType === 'multichoice'
            ? 'MultiChoice'
            : f.MappedType === 'choice'
              ? 'Choice'
              : 'Lookup';
        return {
          key: f.InternalName,
          text: `${f.Title} (${f.InternalName}) — ${kind}`,
        };
      }),
    [fields]
  );

  const selectedField = useMemo((): IFieldMetadata | undefined => {
    for (let i = 0; i < fields.length; i++) {
      if (fields[i].InternalName === selectedKey) return fields[i];
    }
    return undefined;
  }, [fields, selectedKey]);

  const kindSel = fieldKind(selectedField);
  const choiceCount = selectedField?.Choices?.length ?? 0;
  const optionHint =
    kindSel === 'lookup'
      ? lookupLoading
        ? 'A carregar itens da lista referenciada…'
        : `${lookupRows.length} registo(s) na lista lookup.`
      : `${choiceCount} opção(ões) na definição do campo.`;

  const handleApply = (): void => {
    if (!selectedField) return;
    const t = Date.now();
    if (kindSel === 'choice') {
      if (choiceCount === 0) return;
      if (target === 'cards') onApply(buildCardsFromChoiceField(selectedField, t), mergeKey);
      else onApply(buildSeriesFromChoiceField(selectedField, t), mergeKey);
      onDismiss();
      return;
    }
    if (kindSel === 'lookup') {
      if (lookupRows.length === 0 || lookupLoading) return;
      if (target === 'cards') onApply(buildCardsFromLookupField(selectedField, lookupRows, t), mergeKey);
      else onApply(buildSeriesFromLookupField(selectedField, lookupRows, t), mergeKey);
      onDismiss();
    }
  };

  const canApply = Boolean(
    selectedField &&
      ((kindSel === 'choice' && choiceCount > 0) ||
        (kindSel === 'lookup' && lookupRows.length > 0 && !lookupLoading))
  );

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking
      styles={{ main: { maxWidth: 480, margin: 'auto' } }}
    >
      <div style={{ padding: 24 }}>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600, marginBottom: 8, display: 'block' } }}>
          Avançada — gerar por campo Choice ou Lookup
        </Text>
        <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 16, display: 'block' } }}>
          {target === 'cards'
            ? 'Será criado um card por opção (ou por item lookup), cada um com contagem e filtro.'
            : 'Será criada uma série por opção (ou por item lookup), cada uma com contagem e filtro.'}
          {' '}
          Em MultiChoice o filtro usa &quot;contém&quot; no texto. Em Lookup o filtro usa o Id do item referenciado.
        </Text>

        {loading && (
          <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" styles={{ root: { marginBottom: 16 } }}>
            <Spinner size={SpinnerSize.small} />
            <Text variant="small">Carregando campos...</Text>
          </Stack>
        )}

        {error && (
          <Text variant="small" styles={{ root: { color: '#d13438', marginBottom: 12, display: 'block' } }}>
            {error}
          </Text>
        )}

        {!loading && fields.length === 0 && !error && (
          <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 16 } }}>
            Nenhum campo Choice, MultiChoice ou Lookup simples visível nesta lista.
          </Text>
        )}

        {!loading && fields.length > 0 && (
          <>
            <Dropdown
              label="Campo"
              options={dropdownOptions}
              selectedKey={selectedKey || undefined}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) =>
                opt && setSelectedKey(String(opt.key))
              }
              styles={{ root: { marginBottom: 12 } }}
            />
            {selectedField && (
              <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center" styles={{ root: { marginBottom: 12 } }}>
                {lookupLoading && <Spinner size={SpinnerSize.small} />}
                <Text variant="small" styles={{ root: { color: '#605e5c', display: 'block' } }}>{optionHint}</Text>
              </Stack>
            )}
            <ChoiceGroup
              label="Como aplicar"
              options={MERGE_OPTIONS}
              selectedKey={mergeKey}
              onChange={(_: React.FormEvent<HTMLElement | HTMLInputElement> | undefined, opt?: IChoiceGroupOption) => {
                if (opt?.key === 'append' || opt?.key === 'replace') setMergeKey(opt.key);
              }}
              styles={{ root: { marginBottom: 16 } }}
            />
          </>
        )}

        <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="end">
          <DefaultButton text="Cancelar" onClick={onDismiss} />
          <PrimaryButton text="Gerar" onClick={handleApply} disabled={!canApply || loading || lookupLoading} />
        </Stack>
      </div>
    </Modal>
  );
};
