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
import type {
  IFormFieldConfig,
  TFormManagerFormMode,
  TFormConditionOp,
  TFormRule,
} from '../../core/config/types/formManager';
import { FORM_ATTACHMENTS_FIELD_INTERNAL, isFormBannerFieldConfig } from '../../core/config/types/formManager';
import {
  buildFieldUiRules,
  CONDITION_OP_OPTIONS,
  emptyFieldRuleEditorState,
  fieldRuleStateFromRules,
  mergeFieldRuleEditorState,
  type IFieldRuleEditorState,
  type IWhenUi,
  templateFieldRulesDateNotPast,
  templateFieldRulesEmail,
} from '../../core/formManager/formManagerVisualModel';
import { FormManagerCollapseSection } from './FormManagerComponentsTab';

const TEXT_RULES_COLLAPSE_IDS = {
  display: 'textRulesDisplay',
  validation: 'textRulesValidation',
  transform: 'textRulesTransform',
  masks: 'textRulesMasks',
  conditionals: 'textRulesConditionals',
} as const;

export interface IFormFieldRulesPanelProps {
  isOpen: boolean;
  internalName: string;
  fieldConfig: IFormFieldConfig;
  meta: IFieldMetadata | undefined;
  rules: TFormRule[];
  fieldOptions: IDropdownOption[];
  /** Pastas da árvore em Anexos (biblioteca); para valor calculado = URL da pasta. */
  attachmentLibraryFolderOptions?: IDropdownOption[];
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
  attachmentLibraryFolderOptions = [],
  onDismiss,
  onApply,
}) => {
  const [fc, setFc] = useState<IFormFieldConfig>(fieldConfig);
  const [ed, setEd] = useState<IFieldRuleEditorState>(() => emptyFieldRuleEditorState());
  const [textRulesOpen, setTextRulesOpen] = useState<Record<string, boolean>>({});

  useEffect(() => {
    if (!isOpen) return;
    setFc({ ...fieldConfig });
    const st = fieldRuleStateFromRules(internalName, rules);
    const df = String(fieldOptions[0]?.key ?? 'Title');
    if (!st.disableWhenActive && !st.enableWhenActive) {
      st.disableWhenUi = { ...st.disableWhenUi, field: df };
      st.enableWhenUi = { ...st.enableWhenUi, field: df };
    }
    setEd(st);
  }, [isOpen, internalName, fieldConfig, rules, fieldOptions]);

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

  const toggleTextRulesSection = useCallback((id: string): void => {
    setTextRulesOpen((prev) => ({ ...prev, [id]: !prev[id] }));
  }, []);
  const isTextRulesOpen = useCallback((id: string): boolean => textRulesOpen[id] === true, [textRulesOpen]);

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
        {mt !== 'text' && (
          <>
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
            {(mt === 'multiline' || mt === 'url') && (
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
            {internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL && !isFormBannerFieldConfig(fieldConfig) && (
              <Stack tokens={{ childrenGap: 8 }} styles={{ root: { borderTop: '1px solid #edebe9', paddingTop: 12 } }}>
                <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                  Valor calculado (setComputed)
                </Text>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Número: expressão com {'{{NomeInterno}}'} e + − * / ( ). Texto: comece com{' '}
                  <code style={{ fontSize: 12 }}>str:</code> e use {'{{campo}}'}; dentro do texto use{' '}
                  <code style={{ fontSize: 12 }}>[me]</code> (id do utilizador),{' '}
                  <code style={{ fontSize: 12 }}>[myName]</code>, <code style={{ fontSize: 12 }}>[myEmail]</code>,{' '}
                  <code style={{ fontSize: 12 }}>[today]</code>, <code style={{ fontSize: 12 }}>[siteTitle]</code>,{' '}
                  <code style={{ fontSize: 12 }}>[query:chave]</code> (URL), etc. Só um token:{' '}
                  <code style={{ fontSize: 12 }}>[myEmail]</code> sem prefixo. URL de pasta de anexos: dropdown abaixo ou{' '}
                  <code style={{ fontSize: 12 }}>attfolder:idDoNo</code> (com item já gravado e pastas configuradas em Anexos).
                </Text>
                {attachmentLibraryFolderOptions.length > 0 && (
                  <Dropdown
                    label="URL da pasta na biblioteca de anexos"
                    options={[{ key: '', text: '— Não usar (expressão manual)' }, ...attachmentLibraryFolderOptions]}
                    selectedKey={ed.computedAttachmentFolderNodeId || ''}
                    onChange={(_, o) => {
                      const k = o ? String(o.key) : '';
                      setEd((p) => ({
                        ...p,
                        computedAttachmentFolderNodeId: k,
                        computedExpression: k ? '' : p.computedExpression,
                      }));
                    }}
                  />
                )}
                <TextField
                  label={
                    ed.computedAttachmentFolderNodeId ? 'Expressão (desative a pasta acima para editar)' : 'Expressão'
                  }
                  multiline
                  rows={3}
                  value={ed.computedAttachmentFolderNodeId ? '' : ed.computedExpression}
                  disabled={!!ed.computedAttachmentFolderNodeId}
                  onChange={(_, v) =>
                    setEd((p) => ({
                      ...p,
                      computedExpression: v ?? '',
                      computedAttachmentFolderNodeId: '',
                    }))
                  }
                />
              </Stack>
            )}
            <Stack tokens={{ childrenGap: 8 }} styles={{ root: { borderTop: '1px solid #edebe9', paddingTop: 12 } }}>
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                Desativar / ativar o campo
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Condição no mesmo estilo das regras condicionais. Se ambas forem verdadeiras, «Tornar editável quando»
                prevalece sobre «Desativar quando».
              </Text>
              <Checkbox
                label="Desativar este campo quando a condição for verdadeira"
                checked={ed.disableWhenActive}
                onChange={(_, c) => setEd((p) => ({ ...p, disableWhenActive: !!c }))}
              />
              <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Dropdown
                  label="Campo"
                  options={fieldOptions}
                  selectedKey={ed.disableWhenUi.field}
                  disabled={!ed.disableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      disableWhenUi: { ...p.disableWhenUi, field: String(o.key) },
                    }))
                  }
                  styles={{ dropdown: { width: 160 } }}
                />
                <Dropdown
                  label="Operador"
                  options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                  selectedKey={ed.disableWhenUi.op}
                  disabled={!ed.disableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      disableWhenUi: { ...p.disableWhenUi, op: o.key as TFormConditionOp },
                    }))
                  }
                  styles={{ dropdown: { width: 150 } }}
                />
                <Dropdown
                  label="Comparar"
                  options={[
                    { key: 'literal', text: 'Texto fixo' },
                    { key: 'field', text: 'Campo' },
                    { key: 'token', text: 'Token' },
                  ]}
                  selectedKey={ed.disableWhenUi.compareKind}
                  disabled={!ed.disableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      disableWhenUi: { ...p.disableWhenUi, compareKind: o.key as IWhenUi['compareKind'] },
                    }))
                  }
                  styles={{ dropdown: { width: 112 } }}
                />
                <TextField
                  label="Valor"
                  value={ed.disableWhenUi.compareValue}
                  disabled={
                    !ed.disableWhenActive ||
                    ed.disableWhenUi.op === 'isEmpty' ||
                    ed.disableWhenUi.op === 'isFilled' ||
                    ed.disableWhenUi.op === 'isTrue' ||
                    ed.disableWhenUi.op === 'isFalse'
                  }
                  onChange={(_, v) =>
                    setEd((p) => ({
                      ...p,
                      disableWhenUi: { ...p.disableWhenUi, compareValue: v ?? '' },
                    }))
                  }
                  styles={{ fieldGroup: { minWidth: 120 } }}
                />
              </Stack>
              <Checkbox
                label="Tornar editável quando a condição for verdadeira (sobrepor desativação acima)"
                checked={ed.enableWhenActive}
                onChange={(_, c) => setEd((p) => ({ ...p, enableWhenActive: !!c }))}
              />
              <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Dropdown
                  label="Campo"
                  options={fieldOptions}
                  selectedKey={ed.enableWhenUi.field}
                  disabled={!ed.enableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      enableWhenUi: { ...p.enableWhenUi, field: String(o.key) },
                    }))
                  }
                  styles={{ dropdown: { width: 160 } }}
                />
                <Dropdown
                  label="Operador"
                  options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                  selectedKey={ed.enableWhenUi.op}
                  disabled={!ed.enableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      enableWhenUi: { ...p.enableWhenUi, op: o.key as TFormConditionOp },
                    }))
                  }
                  styles={{ dropdown: { width: 150 } }}
                />
                <Dropdown
                  label="Comparar"
                  options={[
                    { key: 'literal', text: 'Texto fixo' },
                    { key: 'field', text: 'Campo' },
                    { key: 'token', text: 'Token' },
                  ]}
                  selectedKey={ed.enableWhenUi.compareKind}
                  disabled={!ed.enableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      enableWhenUi: { ...p.enableWhenUi, compareKind: o.key as IWhenUi['compareKind'] },
                    }))
                  }
                  styles={{ dropdown: { width: 112 } }}
                />
                <TextField
                  label="Valor"
                  value={ed.enableWhenUi.compareValue}
                  disabled={
                    !ed.enableWhenActive ||
                    ed.enableWhenUi.op === 'isEmpty' ||
                    ed.enableWhenUi.op === 'isFilled' ||
                    ed.enableWhenUi.op === 'isTrue' ||
                    ed.enableWhenUi.op === 'isFalse'
                  }
                  onChange={(_, v) =>
                    setEd((p) => ({
                      ...p,
                      enableWhenUi: { ...p.enableWhenUi, compareValue: v ?? '' },
                    }))
                  }
                  styles={{ fieldGroup: { minWidth: 120 } }}
                />
              </Stack>
            </Stack>
          </>
        )}
        {mt === 'text' && (
          <Stack tokens={{ childrenGap: 10 }}>
            <FormManagerCollapseSection
              title="Exibição"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.display)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.display)}
            >
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
              <TextField
                label="Placeholder"
                value={fc.placeholder ?? ''}
                onChange={(_, v) => setFc((p) => ({ ...p, placeholder: v || undefined }))}
              />
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
              {internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL && !isFormBannerFieldConfig(fieldConfig) && (
                <Stack tokens={{ childrenGap: 8 }}>
                  <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                    Valor calculado (setComputed)
                  </Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Número: expressão com {'{{NomeInterno}}'} e + − * / ( ). Texto: comece com{' '}
                    <code style={{ fontSize: 12 }}>str:</code> e use {'{{campo}}'}; dentro do texto use{' '}
                    <code style={{ fontSize: 12 }}>[me]</code> (id do utilizador),{' '}
                    <code style={{ fontSize: 12 }}>[myName]</code>, <code style={{ fontSize: 12 }}>[myEmail]</code>,{' '}
                    <code style={{ fontSize: 12 }}>[today]</code>, <code style={{ fontSize: 12 }}>[siteTitle]</code>,{' '}
                    <code style={{ fontSize: 12 }}>[query:chave]</code> (URL), etc. Só um token:{' '}
                    <code style={{ fontSize: 12 }}>[myEmail]</code> sem prefixo. URL de pasta de anexos: dropdown abaixo ou{' '}
                    <code style={{ fontSize: 12 }}>attfolder:idDoNo</code> (com item já gravado e pastas configuradas em Anexos).
                  </Text>
                  {attachmentLibraryFolderOptions.length > 0 && (
                    <Dropdown
                      label="URL da pasta na biblioteca de anexos"
                      options={[{ key: '', text: '— Não usar (expressão manual)' }, ...attachmentLibraryFolderOptions]}
                      selectedKey={ed.computedAttachmentFolderNodeId || ''}
                      onChange={(_, o) => {
                        const k = o ? String(o.key) : '';
                        setEd((p) => ({
                          ...p,
                          computedAttachmentFolderNodeId: k,
                          computedExpression: k ? '' : p.computedExpression,
                        }));
                      }}
                    />
                  )}
                  <TextField
                    label={
                      ed.computedAttachmentFolderNodeId ? 'Expressão (desative a pasta acima para editar)' : 'Expressão'
                    }
                    multiline
                    rows={3}
                    value={ed.computedAttachmentFolderNodeId ? '' : ed.computedExpression}
                    disabled={!!ed.computedAttachmentFolderNodeId}
                    onChange={(_, v) =>
                      setEd((p) => ({
                        ...p,
                        computedExpression: v ?? '',
                        computedAttachmentFolderNodeId: '',
                      }))
                    }
                  />
                </Stack>
              )}
              <Checkbox
                label="Somente leitura"
                checked={fc.readOnly === true}
                onChange={(_, c) =>
                  setFc((p) => ({
                    ...p,
                    ...(c ? { readOnly: true } : { readOnly: undefined }),
                  }))
                }
              />
              <Checkbox
                label="Ocultar no formulário"
                checked={fc.visible === false}
                onChange={(_, c) =>
                  setFc((p) => {
                    const next: IFormFieldConfig = { ...p };
                    if (c) next.visible = false;
                    else delete next.visible;
                    return next;
                  })
                }
              />
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Validação"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.validation)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.validation)}
            >
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
              <Checkbox
                label="Obrigatório"
                checked={fc.required === true}
                onChange={(_, c) =>
                  setFc((p) => ({
                    ...p,
                    ...(c ? { required: true } : { required: undefined }),
                  }))
                }
              />
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
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Pré-visualização: {buildFieldUiRules(internalName, ed).length} regra(s) gerada(s) para este campo.
              </Text>
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Transformação"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.transform)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.transform)}
            >
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Maiúsculas, minúsculas e capitalização não estão disponíveis nesta UI; o motor do formulário ainda não
                expõe essas opções. Use o JSON do gestor quando existir suporte.
              </Text>
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Máscaras"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.masks)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.masks)}
            >
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Máscaras (CPF, telefone, CEP, CNPJ ou personalizada) não estão disponíveis nesta UI nem no motor atual.
                Configure no JSON do gestor quando existir suporte.
              </Text>
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Condicionais"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.conditionals)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.conditionals)}
            >
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Mostrar ou ocultar este campo consoante outro («se o campo X for Y») usa regras globais (
                <code style={{ fontSize: 12 }}>setVisibility</code>, cartões em JSON na aba «JSON»). Desativar ou
                tornar editável consoante outro campo configura-se abaixo; se ambas as condições forem verdadeiras,
                «Tornar editável quando» prevalece sobre «Desativar quando».
              </Text>
              <Checkbox
                label="Desativar este campo quando a condição for verdadeira"
                checked={ed.disableWhenActive}
                onChange={(_, c) => setEd((p) => ({ ...p, disableWhenActive: !!c }))}
              />
              <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Dropdown
                  label="Campo"
                  options={fieldOptions}
                  selectedKey={ed.disableWhenUi.field}
                  disabled={!ed.disableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      disableWhenUi: { ...p.disableWhenUi, field: String(o.key) },
                    }))
                  }
                  styles={{ dropdown: { width: 160 } }}
                />
                <Dropdown
                  label="Operador"
                  options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                  selectedKey={ed.disableWhenUi.op}
                  disabled={!ed.disableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      disableWhenUi: { ...p.disableWhenUi, op: o.key as TFormConditionOp },
                    }))
                  }
                  styles={{ dropdown: { width: 150 } }}
                />
                <Dropdown
                  label="Comparar"
                  options={[
                    { key: 'literal', text: 'Texto fixo' },
                    { key: 'field', text: 'Campo' },
                    { key: 'token', text: 'Token' },
                  ]}
                  selectedKey={ed.disableWhenUi.compareKind}
                  disabled={!ed.disableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      disableWhenUi: { ...p.disableWhenUi, compareKind: o.key as IWhenUi['compareKind'] },
                    }))
                  }
                  styles={{ dropdown: { width: 112 } }}
                />
                <TextField
                  label="Valor"
                  value={ed.disableWhenUi.compareValue}
                  disabled={
                    !ed.disableWhenActive ||
                    ed.disableWhenUi.op === 'isEmpty' ||
                    ed.disableWhenUi.op === 'isFilled' ||
                    ed.disableWhenUi.op === 'isTrue' ||
                    ed.disableWhenUi.op === 'isFalse'
                  }
                  onChange={(_, v) =>
                    setEd((p) => ({
                      ...p,
                      disableWhenUi: { ...p.disableWhenUi, compareValue: v ?? '' },
                    }))
                  }
                  styles={{ fieldGroup: { minWidth: 120 } }}
                />
              </Stack>
              <Checkbox
                label="Tornar editável quando a condição for verdadeira (sobrepor desativação acima)"
                checked={ed.enableWhenActive}
                onChange={(_, c) => setEd((p) => ({ ...p, enableWhenActive: !!c }))}
              />
              <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Dropdown
                  label="Campo"
                  options={fieldOptions}
                  selectedKey={ed.enableWhenUi.field}
                  disabled={!ed.enableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      enableWhenUi: { ...p.enableWhenUi, field: String(o.key) },
                    }))
                  }
                  styles={{ dropdown: { width: 160 } }}
                />
                <Dropdown
                  label="Operador"
                  options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                  selectedKey={ed.enableWhenUi.op}
                  disabled={!ed.enableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      enableWhenUi: { ...p.enableWhenUi, op: o.key as TFormConditionOp },
                    }))
                  }
                  styles={{ dropdown: { width: 150 } }}
                />
                <Dropdown
                  label="Comparar"
                  options={[
                    { key: 'literal', text: 'Texto fixo' },
                    { key: 'field', text: 'Campo' },
                    { key: 'token', text: 'Token' },
                  ]}
                  selectedKey={ed.enableWhenUi.compareKind}
                  disabled={!ed.enableWhenActive}
                  onChange={(_, o) =>
                    o &&
                    setEd((p) => ({
                      ...p,
                      enableWhenUi: { ...p.enableWhenUi, compareKind: o.key as IWhenUi['compareKind'] },
                    }))
                  }
                  styles={{ dropdown: { width: 112 } }}
                />
                <TextField
                  label="Valor"
                  value={ed.enableWhenUi.compareValue}
                  disabled={
                    !ed.enableWhenActive ||
                    ed.enableWhenUi.op === 'isEmpty' ||
                    ed.enableWhenUi.op === 'isFilled' ||
                    ed.enableWhenUi.op === 'isTrue' ||
                    ed.enableWhenUi.op === 'isFalse'
                  }
                  onChange={(_, v) =>
                    setEd((p) => ({
                      ...p,
                      enableWhenUi: { ...p.enableWhenUi, compareValue: v ?? '' },
                    }))
                  }
                  styles={{ fieldGroup: { minWidth: 120 } }}
                />
              </Stack>
            </FormManagerCollapseSection>
          </Stack>
        )}
        {(mt === 'multiline' || mt === 'url' || mt === 'unknown') && (
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
            Condições «se outro campo = X» entre colunas: JSON do gestor. Neste painel: obrigatoriedade e validação de
            texto, quando aplicável.
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
            Use valor padrão acima (true/false). Visibilidade condicional: opções neste painel ou JSON do gestor.
          </Text>
        )}
        {mt !== 'text' && (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Pré-visualização: {buildFieldUiRules(internalName, ed).length} regra(s) gerada(s) para este campo.
          </Text>
        )}
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Aplicar" onClick={handleApply} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      </Stack>
    </Panel>
  );
};
