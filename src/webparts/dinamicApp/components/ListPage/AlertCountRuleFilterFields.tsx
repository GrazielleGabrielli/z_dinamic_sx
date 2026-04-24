import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import { Dropdown, IDropdownOption, Link, MessageBar, MessageBarType, Stack, Text, TextField } from '@fluentui/react';
import type { IListPageAlertCountRule, TListPageAlertCountFilterFieldOp } from '../../core/config/types';
import type { IFieldMetadata } from '../../../../services/shared/types';
import { FieldsService } from '../../../../services';
import {
  isFieldEligibleForAlertCountFilter,
  mergeCountRuleODataFromStructured,
  operatorsForCountFilterField,
} from '../../core/listPage/alertCountRuleFilterOData';

const BOOLEAN_OPTS: IDropdownOption[] = [
  { key: '1', text: 'Sim' },
  { key: '0', text: 'Não' },
];

const OP_LABELS: Partial<Record<TListPageAlertCountFilterFieldOp, string>> = {
  eq: 'Igual a',
  ne: 'Diferente de',
  gt: 'Maior que',
  ge: 'Maior ou igual',
  lt: 'Menor que',
  le: 'Menor ou igual',
  contains: 'Contém (texto)',
};

function opDropdownOptions(ops: TListPageAlertCountFilterFieldOp[]): IDropdownOption[] {
  return ops.map((k) => ({ key: k, text: OP_LABELS[k] ?? k }));
}

export interface IAlertCountRuleFilterFieldsProps {
  listTitle: string;
  fields: IFieldMetadata[] | undefined;
  rule: IListPageAlertCountRule;
  onRuleChange: (next: IListPageAlertCountRule) => void;
}

export const AlertCountRuleFilterFields: React.FC<IAlertCountRuleFilterFieldsProps> = ({
  listTitle,
  fields,
  rule,
  onRuleChange,
}) => {
  const fieldsSvc = useMemo(() => new FieldsService(), []);
  const byInternal = useMemo(
    () => new Map((fields ?? []).map((f) => [f.InternalName, f])),
    [fields]
  );

  const [choiceExtra, setChoiceExtra] = useState<string[]>([]);
  const selectedMeta = rule.countFilterField ? byInternal.get(rule.countFilterField) : undefined;

  useEffect(() => {
    if (!listTitle.trim() || !rule.countFilterField || !selectedMeta) {
      setChoiceExtra([]);
      return;
    }
    const mt = selectedMeta.MappedType;
    if (mt !== 'choice' && mt !== 'multichoice') {
      setChoiceExtra([]);
      return;
    }
    if (selectedMeta.Choices && selectedMeta.Choices.length > 0) {
      setChoiceExtra(selectedMeta.Choices);
      return;
    }
    let cancelled = false;
    fieldsSvc
      .getFieldOptions(listTitle.trim(), rule.countFilterField)
      .then((opts) => {
        if (!cancelled) setChoiceExtra(opts ?? []);
      })
      .catch(() => {
        if (!cancelled) setChoiceExtra([]);
      });
    return () => {
      cancelled = true;
    };
  }, [listTitle, rule.countFilterField, selectedMeta, fieldsSvc]);

  const patch = useCallback(
    (p: Partial<IListPageAlertCountRule>) => {
      const next = { ...rule, ...p };
      onRuleChange(mergeCountRuleODataFromStructured(next, byInternal));
    },
    [rule, onRuleChange, byInternal]
  );

  const fieldOptions: IDropdownOption[] = useMemo(() => {
    const base: IDropdownOption[] = [{ key: '', text: '(Todos os itens — sem filtro por campo)' }];
    const list = (fields ?? []).filter(isFieldEligibleForAlertCountFilter);
    for (let i = 0; i < list.length; i++) {
      const f = list[i];
      base.push({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` });
    }
    return base;
  }, [fields]);

  const hasStructured = Boolean(rule.countFilterField?.trim());
  const legacyManual =
    rule.countFilterUseManualOdata === true ||
    (Boolean((rule.odataFilter ?? '').trim()) && !hasStructured);

  const allowedOps = operatorsForCountFilterField(selectedMeta);
  const opOptions = opDropdownOptions(allowedOps);
  const choiceLike =
    selectedMeta &&
    (selectedMeta.MappedType === 'choice' ||
      selectedMeta.MappedType === 'multichoice' ||
      selectedMeta.MappedType === 'taxonomy' ||
      selectedMeta.MappedType === 'taxonomymulti');
  const choiceValues =
    choiceLike && (choiceExtra.length > 0 || (selectedMeta.Choices?.length ?? 0) > 0)
      ? choiceExtra.length > 0
        ? choiceExtra
        : selectedMeta?.Choices ?? []
      : [];

  const onFieldKey = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    const key = String(opt?.key ?? '');
    if (!key) {
      onRuleChange({
        ...rule,
        countFilterField: undefined,
        countFilterFieldOp: undefined,
        countFilterValue: undefined,
        odataFilter: undefined,
        countFilterUseManualOdata: false,
      });
      return;
    }
    const meta = byInternal.get(key);
    const ops = operatorsForCountFilterField(meta);
    const next: IListPageAlertCountRule = {
      ...rule,
      countFilterField: key,
      countFilterFieldOp: ops[0],
      countFilterValue: meta?.MappedType === 'boolean' ? '0' : '',
      countFilterUseManualOdata: false,
    };
    onRuleChange(mergeCountRuleODataFromStructured(next, byInternal));
  };

  const onOpKey = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    const k = String(opt?.key ?? 'eq') as TListPageAlertCountFilterFieldOp;
    patch({ countFilterFieldOp: k });
  };

  const switchToBuilder = (): void => {
    onRuleChange({
      ...rule,
      odataFilter: undefined,
      countFilterField: undefined,
      countFilterFieldOp: undefined,
      countFilterValue: undefined,
      countFilterUseManualOdata: false,
    });
  };

  const valueControl = (): React.ReactNode => {
    if (!selectedMeta || !hasStructured) return null;
    const mt = selectedMeta.MappedType;
    if (mt === 'boolean') {
      return (
        <Dropdown
          label="Valor"
          selectedKey={rule.countFilterValue === '0' || rule.countFilterValue === '1' ? rule.countFilterValue : '0'}
          options={BOOLEAN_OPTS}
          onChange={(_, o) => o && patch({ countFilterValue: String(o.key) })}
        />
      );
    }
    if (mt === 'choice' || mt === 'multichoice' || mt === 'taxonomy' || mt === 'taxonomymulti') {
      const opts: IDropdownOption[] = choiceValues.map((c) => ({ key: c, text: c }));
      if (opts.length === 0) {
        return (
          <TextField
            label="Valor"
            value={rule.countFilterValue ?? ''}
            onChange={(_, v) => patch({ countFilterValue: v ?? '' })}
            description="Sem opções carregadas; escreva o texto exato da opção."
          />
        );
      }
      return (
        <Dropdown
          label="Valor"
          selectedKey={rule.countFilterValue ?? ''}
          options={opts}
          onChange={(_, o) => o && patch({ countFilterValue: String(o.key) })}
        />
      );
    }
    if (mt === 'number' || mt === 'currency') {
      return (
        <TextField
          label="Valor (número)"
          type="number"
          value={rule.countFilterValue ?? ''}
          onChange={(_, v) => patch({ countFilterValue: v ?? '' })}
        />
      );
    }
    if (mt === 'lookup' || mt === 'user') {
      return (
        <TextField
          label="Id do item referenciado"
          type="number"
          value={rule.countFilterValue ?? ''}
          onChange={(_, v) => patch({ countFilterValue: v ?? '' })}
          description={`Gera filtro no campo «${selectedMeta.InternalName}Id».`}
        />
      );
    }
    if (mt === 'datetime') {
      return (
        <TextField
          label="Data/hora (ISO 8601)"
          value={rule.countFilterValue ?? ''}
          onChange={(_, v) => patch({ countFilterValue: v ?? '' })}
          placeholder="2025-01-15T12:00:00Z"
          description="Ex.: 2025-01-15T12:00:00Z — é acrescentado o prefixo datetime'…' no OData."
        />
      );
    }
    return (
      <TextField
        label="Valor"
        value={rule.countFilterValue ?? ''}
        onChange={(_, v) => patch({ countFilterValue: v ?? '' })}
      />
    );
  };

  if (!listTitle.trim()) {
    return (
      <Text variant="small" styles={{ root: { color: '#a4262c' } }}>
        Defina o título da lista na vista para filtrar por campo.
      </Text>
    );
  }

  if (legacyManual) {
    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <MessageBar messageBarType={MessageBarType.warning}>
          Filtro OData em texto livre. Pode editar abaixo ou voltar ao construtor por campos.
        </MessageBar>
        <TextField
          label="Filtro OData"
          multiline
          rows={3}
          value={rule.odataFilter ?? ''}
          onChange={(_, v) =>
            onRuleChange({
              ...rule,
              odataFilter: (v ?? '').trim() || undefined,
              countFilterField: undefined,
              countFilterFieldOp: undefined,
              countFilterValue: undefined,
              countFilterUseManualOdata: true,
            })
          }
        />
        <Link onClick={switchToBuilder}>Usar construtor por campos (substitui o texto manual)</Link>
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 10 }}>
      <Dropdown
        label="Campo da lista"
        selectedKey={rule.countFilterField ?? ''}
        options={fieldOptions}
        onChange={onFieldKey}
        disabled={!fields?.length}
      />
      {!fields?.length ? (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          A carregar campos…
        </Text>
      ) : null}
      {hasStructured ? (
        <>
          <Dropdown
            label="Condição no campo"
            selectedKey={
              allowedOps.indexOf((rule.countFilterFieldOp ?? 'eq') as TListPageAlertCountFilterFieldOp) !== -1
                ? rule.countFilterFieldOp
                : allowedOps[0]
            }
            options={opOptions}
            onChange={onOpKey}
          />
          {valueControl()}
          {rule.odataFilter ? (
            <Text
              variant="small"
              styles={{
                root: {
                  color: '#605e5c',
                  fontFamily: 'monospace',
                  wordBreak: 'break-all',
                  lineHeight: 1.45,
                },
              }}
            >
              OData: {rule.odataFilter}
            </Text>
          ) : null}
        </>
      ) : (
        <Link
          onClick={() =>
            onRuleChange({
              ...rule,
              countFilterUseManualOdata: true,
              countFilterField: undefined,
              countFilterFieldOp: undefined,
              countFilterValue: undefined,
              odataFilter: rule.odataFilter ?? '',
            })
          }
        >
          Editar OData em texto livre
        </Link>
      )}
    </Stack>
  );
};
