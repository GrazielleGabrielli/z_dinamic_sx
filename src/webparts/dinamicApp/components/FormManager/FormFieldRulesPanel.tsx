import * as React from 'react';
import { useEffect, useState, useCallback } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  Checkbox,
  Dropdown,
  IDropdownOption,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type { IFormFieldConfig, TFormManagerFormMode, TFormRule } from '../../core/config/types/formManager';
import {
  buildFieldUiRules,
  emptyFieldRuleEditorState,
  fieldRuleStateFromRules,
  mergeFieldRuleEditorState,
  type IFieldRuleEditorState,
  templateFieldRulesDateNotPast,
  templateFieldRulesEmail,
} from '../../core/formManager/formManagerVisualModel';

export interface IFormFieldRulesPanelProps {
  isOpen: boolean;
  internalName: string;
  fieldConfig: IFormFieldConfig;
  meta: IFieldMetadata | undefined;
  rules: TFormRule[];
  fieldOptions: IDropdownOption[];
  onDismiss: () => void;
  onApply: (nextField: IFormFieldConfig, editor: IFieldRuleEditorState) => void;
}

const MODE_OPTS: { key: TFormManagerFormMode; label: string }[] = [
  { key: 'create', label: 'Criar' },
  { key: 'edit', label: 'Editar' },
  { key: 'view', label: 'Ver' },
];

const ALL_MODES: TFormManagerFormMode[] = ['create', 'edit', 'view'];

export const FormFieldRulesPanel: React.FC<IFormFieldRulesPanelProps> = ({
  isOpen,
  internalName,
  fieldConfig,
  meta,
  rules,
  fieldOptions,
  onDismiss,
  onApply,
}) => {
  const [fc, setFc] = useState<IFormFieldConfig>(fieldConfig);
  const [ed, setEd] = useState<IFieldRuleEditorState>(() => emptyFieldRuleEditorState());

  useEffect(() => {
    if (!isOpen) return;
    setFc({ ...fieldConfig });
    setEd(fieldRuleStateFromRules(internalName, rules));
  }, [isOpen, internalName, fieldConfig, rules]);

  const mt = meta?.MappedType ?? 'unknown';
  const title = meta?.Title ?? internalName;

  const toggleModeRow = useCallback((m: TFormManagerFormMode, checked: boolean) => {
    setEd((prev) => {
      let next = prev.modes.length === 0 ? ALL_MODES.slice() : prev.modes.slice();
      if (checked) {
        if (next.indexOf(m) === -1) next.push(m);
      } else {
        next = next.filter((x) => x !== m);
      }
      if (next.length === ALL_MODES.length) return { ...prev, modes: [] };
      return { ...prev, modes: next };
    });
  }, []);

  const modeRowChecked = useCallback((m: TFormManagerFormMode): boolean => {
    return ed.modes.length === 0 || ed.modes.indexOf(m) !== -1;
  }, [ed.modes]);

  const handleApply = (): void => {
    onApply(fc, ed);
    onDismiss();
  };

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.medium}
      headerText={`Configurar regras — ${title}`}
      onDismiss={onDismiss}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {internalName} · {mt}
          {fc.sectionId ? ` · etapa ${fc.sectionId}` : ''}
        </Text>
        <Text variant="small">Aplicar regras geradas apenas nos modos:</Text>
        <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
          {MODE_OPTS.map((m) => (
            <Checkbox
              key={m.key}
              label={m.label}
              checked={modeRowChecked(m.key)}
              onChange={(_, c) => toggleModeRow(m.key, !!c)}
            />
          ))}
        </Stack>
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Vazio = todos os modos. Desmarque um para restringir.
        </Text>
        <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
          <DefaultButton
            text="Modelo: data não no passado"
            onClick={() => setEd((prev) => mergeFieldRuleEditorState(prev, templateFieldRulesDateNotPast()))}
          />
          <DefaultButton
            text="Modelo: validar e-mail"
            onClick={() => setEd((prev) => mergeFieldRuleEditorState(prev, templateFieldRulesEmail()))}
          />
        </Stack>
        {(mt === 'text' || mt === 'multiline' || mt === 'url') && (
          <TextField
            label="Placeholder"
            value={fc.placeholder ?? ''}
            onChange={(_, v) => setFc((p) => ({ ...p, placeholder: v || undefined }))}
          />
        )}
        <TextField
          label="Texto de ajuda (campo)"
          multiline
          rows={2}
          value={fc.helpText ?? ''}
          onChange={(_, v) => setFc((p) => ({ ...p, helpText: v || undefined }))}
        />
        <TextField
          label="Valor padrão (token ou texto; aplica se vazio)"
          value={ed.defaultValue}
          onChange={(_, v) => setEd((p) => ({ ...p, defaultValue: v ?? '' }))}
        />
        {(mt === 'text' || mt === 'multiline' || mt === 'url' || mt === 'unknown') && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Validação de texto</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              <TextField
                label="Mín. caracteres"
                value={ed.validateValue.minLength}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, minLength: v ?? '' } }))
                }
              />
              <TextField
                label="Máx. caracteres"
                value={ed.validateValue.maxLength}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, maxLength: v ?? '' } }))
                }
              />
            </Stack>
            <TextField
              label="Regex (padrão)"
              value={ed.validateValue.pattern}
              onChange={(_, v) =>
                setEd((p) => ({ ...p, validateValue: { ...p.validateValue, pattern: v ?? '' } }))
              }
            />
            <TextField
              label="Mensagem se falhar o padrão"
              value={ed.validateValue.patternMessage}
              onChange={(_, v) =>
                setEd((p) => ({ ...p, validateValue: { ...p.validateValue, patternMessage: v ?? '' } }))
              }
            />
          </Stack>
        )}
        {(mt === 'number' || mt === 'currency') && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Validação numérica</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              <TextField
                label="Mínimo"
                type="number"
                value={ed.validateValue.minNumber}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, minNumber: v ?? '' } }))
                }
              />
              <TextField
                label="Máximo"
                type="number"
                value={ed.validateValue.maxNumber}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, maxNumber: v ?? '' } }))
                }
              />
            </Stack>
            <TextField
              label="Expressão calculada (setComputed)"
              multiline
              rows={2}
              value={ed.computedExpression}
              onChange={(_, v) => setEd((p) => ({ ...p, computedExpression: v ?? '' }))}
            />
          </Stack>
        )}
        {mt === 'datetime' && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Validação de data</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              <TextField
                label="Mín. dias a partir de hoje"
                value={ed.validateDate.minDaysFromToday}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateDate: { ...p.validateDate, minDaysFromToday: v ?? '' } }))
                }
              />
              <TextField
                label="Máx. dias a partir de hoje"
                value={ed.validateDate.maxDaysFromToday}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateDate: { ...p.validateDate, maxDaysFromToday: v ?? '' } }))
                }
              />
            </Stack>
            <Checkbox
              label="Bloquear fins de semana"
              checked={ed.validateDate.blockWeekends}
              onChange={(_, c) =>
                setEd((p) => ({ ...p, validateDate: { ...p.validateDate, blockWeekends: !!c } }))
              }
            />
            <Dropdown
              label="Data &gt;= campo"
              options={[{ key: '', text: '—' }, ...fieldOptions]}
              selectedKey={ed.validateDate.gteField || ''}
              onChange={(_, o) =>
                setEd((p) => ({
                  ...p,
                  validateDate: { ...p.validateDate, gteField: o ? String(o.key) : '' },
                }))
              }
            />
            <Dropdown
              label="Data &lt;= campo"
              options={[{ key: '', text: '—' }, ...fieldOptions]}
              selectedKey={ed.validateDate.lteField || ''}
              onChange={(_, o) =>
                setEd((p) => ({
                  ...p,
                  validateDate: { ...p.validateDate, lteField: o ? String(o.key) : '' },
                }))
              }
            />
            <TextField
              label="Mensagem de erro"
              value={ed.validateDate.message}
              onChange={(_, v) =>
                setEd((p) => ({ ...p, validateDate: { ...p.validateDate, message: v ?? '' } }))
              }
            />
          </Stack>
        )}
        {(mt === 'choice' || mt === 'multichoice') && (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Condições “se valor = X” use a aba Regras condicionais. Aqui: padrão e validação de texto se aplicável.
          </Text>
        )}
        {(mt === 'lookup' || mt === 'lookupmulti' || mt === 'user' || mt === 'usermulti') && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Lookup / usuário</Text>
            <Dropdown
              label="Campo pai (filtro)"
              options={[{ key: '', text: '—' }, ...fieldOptions]}
              selectedKey={ed.filterLookup.parentField || ''}
              onChange={(_, o) =>
                setEd((p) => ({
                  ...p,
                  filterLookup: { ...p.filterLookup, parentField: o ? String(o.key) : '' },
                }))
              }
            />
            <TextField
              label="Modelo OData (use {'{parent}'} para o Id do pai)"
              multiline
              rows={2}
              value={ed.filterLookup.odataFilterTemplate}
              onChange={(_, v) =>
                setEd((p) => ({
                  ...p,
                  filterLookup: { ...p.filterLookup, odataFilterTemplate: v ?? '' },
                }))
              }
            />
          </Stack>
        )}
        {mt === 'boolean' && (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Use valor padrão acima (true/false). Visibilidade condicional: aba Regras condicionais.
          </Text>
        )}
        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Limpar ao mudar outro campo</Text>
          <Checkbox
            label="Ativar"
            checked={ed.clearOnChange.enabled}
            onChange={(_, c) =>
              setEd((p) => ({ ...p, clearOnChange: { ...p.clearOnChange, enabled: !!c } }))
            }
          />
          <Dropdown
            label="Campo que dispara a limpeza"
            options={[{ key: '', text: '—' }, ...fieldOptions]}
            selectedKey={ed.clearOnChange.triggerField || ''}
            onChange={(_, o) =>
              setEd((p) => ({
                ...p,
                clearOnChange: { ...p.clearOnChange, triggerField: o ? String(o.key) : '' },
              }))
            }
            disabled={!ed.clearOnChange.enabled}
          />
          <TextField
            label="Campos a limpar (internos, separados por vírgula)"
            value={ed.clearOnChange.clearFieldsText}
            onChange={(_, v) =>
              setEd((p) => ({
                ...p,
                clearOnChange: { ...p.clearOnChange, clearFieldsText: v ?? '' },
              }))
            }
            disabled={!ed.clearOnChange.enabled}
          />
        </Stack>
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Pré-visualização: {buildFieldUiRules(internalName, ed).length} regra(s) gerada(s) para este campo.
        </Text>
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Aplicar" onClick={handleApply} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      </Stack>
    </Panel>
  );
};
