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
import { FieldsService } from '../../../../services';
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

function filterOperatorForField(f: IFieldMetadata): TFilterOperator {
  return f.MappedType === 'multichoice' ? 'contains' : 'eq';
}

function choiceSlug(v: string, index: number): string {
  const base = v.replace(/[^a-zA-Z0-9]+/g, '_').replace(/^_|_$/g, '').slice(0, 28);
  return `${base || 'opt'}_${index}`;
}

function buildCardsFromField(field: IFieldMetadata, baseTime: number): IDashboardCardConfig[] {
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

function buildSeriesFromField(field: IFieldMetadata, baseTime: number): IChartSeriesConfig[] {
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

  const fieldsService = useMemo(() => new FieldsService(), []);

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
        const choiceFields = all.filter(isChoiceLike);
        setFields(choiceFields);
        setSelectedKey(choiceFields[0]?.InternalName ?? '');
        setLoading(false);
      })
      .catch((e) => {
        setError(e instanceof Error ? e.message : String(e));
        setFields([]);
        setLoading(false);
      });
  }, [isOpen, listTitle, lw]);

  const dropdownOptions: IDropdownOption[] = useMemo(
    () =>
      fields.map((f) => ({
        key: f.InternalName,
        text: `${f.Title} (${f.InternalName}) — ${f.MappedType === 'multichoice' ? 'MultiChoice' : 'Choice'}`,
      })),
    [fields]
  );

  const selectedField = useMemo((): IFieldMetadata | undefined => {
    for (let i = 0; i < fields.length; i++) {
      if (fields[i].InternalName === selectedKey) return fields[i];
    }
    return undefined;
  }, [fields, selectedKey]);

  const choiceCount = selectedField?.Choices?.length ?? 0;

  const handleApply = (): void => {
    if (!selectedField || choiceCount === 0) return;
    const t = Date.now();
    if (target === 'cards') {
      onApply(buildCardsFromField(selectedField, t), mergeKey);
    } else {
      onApply(buildSeriesFromField(selectedField, t), mergeKey);
    }
    onDismiss();
  };

  const canApply = Boolean(selectedField && choiceCount > 0);

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={onDismiss}
      isBlocking
      styles={{ main: { maxWidth: 480, margin: 'auto' } }}
    >
      <div style={{ padding: 24 }}>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600, marginBottom: 8, display: 'block' } }}>
          Avançada — gerar por campo Choice
        </Text>
        <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 16, display: 'block' } }}>
          {target === 'cards'
            ? 'Será criado um card por opção do campo, cada um com contagem e filtro no valor.'
            : 'Será criada uma série por opção do campo, cada uma com contagem e filtro no valor.'}
          {' '}
          Em MultiChoice o filtro usa &quot;contém&quot; no texto armazenado.
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
            Nenhum campo Choice ou MultiChoice visível nesta lista.
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
              <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 12, display: 'block' } }}>
                {choiceCount} opção(ões) na definição do campo.
              </Text>
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
          <PrimaryButton text="Gerar" onClick={handleApply} disabled={!canApply || loading} />
        </Stack>
      </div>
    </Modal>
  );
};
