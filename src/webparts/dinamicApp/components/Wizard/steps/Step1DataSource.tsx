import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Stack,
  Text,
  ChoiceGroup,
  IChoiceGroupOption,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  TextField,
  DefaultButton,
  PrimaryButton,
  Icon,
  Separator,
} from '@fluentui/react';
import { ListsService, IListSummary, FieldsService } from '../../../../../services';
import type { IFieldMetadata } from '../../../../../services/shared/types';
import { IWizardFormState } from '../types';
import type { IAIStepConfig, IAIButtonConfig } from '../../../../../services/ai/AIConfigService';
import { generateFormStructure, generateFormButtons } from '../../../../../services/ai/AIConfigService';

interface IStep1Props {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
  openAiApiKey?: string;
  onAIStructureApply?: (steps: IAIStepConfig[]) => void;
  onAIButtonsApply?: (buttons: IAIButtonConfig[]) => void;
}

type TAIPhase = 'structure' | 'buttons';

const sourceKindOptions: IChoiceGroupOption[] = [
  { key: 'list', text: 'Lista' },
  { key: 'library', text: 'Biblioteca de documentos' },
];

const BEHAVIOR_LABEL: Record<string, string> = {
  submit: 'Enviar',
  draft: 'Rascunho',
  close: 'Fechar',
  actionsOnly: 'Só ações',
};

const MODE_LABEL: Record<string, string> = {
  create: 'Criar',
  edit: 'Editar',
  view: 'Ver',
};

export const Step1DataSource: React.FC<IStep1Props> = ({
  form,
  onChange,
  openAiApiKey,
  onAIStructureApply,
  onAIButtonsApply,
}) => {
  const [allSources, setAllSources] = useState<IListSummary[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);

  const [fieldMeta, setFieldMeta] = useState<IFieldMetadata[]>([]);
  const [fieldsLoading, setFieldsLoading] = useState(false);

  const [aiPhase, setAiPhase] = useState<TAIPhase>('structure');

  const [aiDescription, setAiDescription] = useState('');
  const [aiLoading, setAiLoading] = useState(false);
  const [aiError, setAiError] = useState<string | undefined>(undefined);
  const [aiPreview, setAiPreview] = useState<IAIStepConfig[] | null>(null);
  const [appliedDescription, setAppliedDescription] = useState('');

  const [btnDescription, setBtnDescription] = useState('');
  const [btnLoading, setBtnLoading] = useState(false);
  const [btnError, setBtnError] = useState<string | undefined>(undefined);
  const [btnPreview, setBtnPreview] = useState<IAIButtonConfig[] | null>(null);

  const showAiSection = !!onAIStructureApply;

  useEffect(() => {
    setAllSources([]);
    setLoading(true);
    const service = new ListsService();
    service
      .getLists(false)
      .then((data) => {
        setAllSources(data);
        setLoading(false);
      })
      .catch((err: Error) => {
        setError(`Não foi possível carregar as origens: ${err.message}`);
        setLoading(false);
      });
  }, []);

  useEffect(() => {
    if (!showAiSection || !form.title.trim()) {
      setFieldMeta([]);
      setAiPreview(null);
      setBtnPreview(null);
      setAiPhase('structure');
      return;
    }
    setFieldsLoading(true);
    setAiPreview(null);
    setBtnPreview(null);
    setAiPhase('structure');
    const service = new FieldsService();
    service
      .getVisibleFields(form.title.trim())
      .then((fields) => {
        setFieldMeta(fields);
        setFieldsLoading(false);
      })
      .catch(() => {
        setFieldMeta([]);
        setFieldsLoading(false);
      });
  }, [form.title, showAiSection]);

  const filtered = allSources.filter((l) =>
    form.kind === 'library' ? l.IsLibrary : !l.IsLibrary
  );

  const dropdownOptions: IDropdownOption[] = filtered.map((l) => ({
    key: l.Title,
    text: l.Title,
  }));

  const handleKindChange = (
    _: React.FormEvent<HTMLElement | HTMLInputElement> | undefined,
    opt?: IChoiceGroupOption
  ): void => {
    if (!opt) return;
    onChange({ kind: opt.key as 'list' | 'library', title: '' });
  };

  const handleTitleChange = (
    _: React.FormEvent<HTMLDivElement>,
    opt?: IDropdownOption
  ): void => {
    if (!opt) return;
    onChange({ title: opt.key as string });
  };

  const handleGenerateStructure = async (): Promise<void> => {
    if (!aiDescription.trim() || !form.title.trim() || !openAiApiKey) return;
    setAiError(undefined);
    setAiPreview(null);
    setAiLoading(true);
    try {
      const result = await generateFormStructure(openAiApiKey, {
        description: aiDescription.trim(),
        listTitle: form.title,
        fields: fieldMeta,
      });
      setAiPreview(result.steps);
    } catch (e) {
      setAiError(e instanceof Error ? e.message : String(e));
    } finally {
      setAiLoading(false);
    }
  };

  const handleApplyStructure = (): void => {
    if (!aiPreview || !onAIStructureApply) return;
    onAIStructureApply(aiPreview);
    setAppliedDescription(aiDescription.trim());
    setAiPhase('buttons');
  };

  const handleGenerateButtons = async (): Promise<void> => {
    if (!btnDescription.trim() || !openAiApiKey) return;
    setBtnError(undefined);
    setBtnPreview(null);
    setBtnLoading(true);
    try {
      const result = await generateFormButtons(openAiApiKey, {
        description: btnDescription.trim(),
        systemDescription: appliedDescription,
      });
      setBtnPreview(result.buttons);
    } catch (e) {
      setBtnError(e instanceof Error ? e.message : String(e));
    } finally {
      setBtnLoading(false);
    }
  };

  const handleApplyButtons = (): void => {
    if (!btnPreview || !onAIButtonsApply) return;
    onAIButtonsApply(btnPreview);
    setBtnPreview(null);
    setBtnDescription('');
  };

  const visiblePreviewSteps = (aiPreview ?? []).filter(
    (s) => s.id !== 'ocultos' && s.id !== 'fixos'
  );

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          Fonte de dados
        </Text>
        <Text variant="medium" styles={{ root: { color: '#605e5c', marginTop: 4, display: 'block' } }}>
          Selecione o tipo de origem e a lista ou biblioteca que será usada.
        </Text>
      </Stack.Item>

      <ChoiceGroup
        label="Tipo de origem"
        options={sourceKindOptions}
        selectedKey={form.kind}
        onChange={handleKindChange}
      />

      {loading && <Spinner size={SpinnerSize.medium} label="Carregando origens disponíveis..." />}

      {error !== undefined && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {error}
        </MessageBar>
      )}

      {!loading && error === undefined && (
        <Dropdown
          label={form.kind === 'library' ? 'Biblioteca' : 'Lista'}
          placeholder={`Selecione uma ${form.kind === 'library' ? 'biblioteca' : 'lista'}`}
          options={dropdownOptions}
          selectedKey={form.title || undefined}
          onChange={handleTitleChange}
          disabled={dropdownOptions.length === 0}
        />
      )}

      {!loading && error === undefined && dropdownOptions.length === 0 && (
        <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
          Nenhuma {form.kind === 'library' ? 'biblioteca' : 'lista'} encontrada neste site.
        </Text>
      )}

      {showAiSection && form.title.trim() && (
        <>
          <Separator />

          {/* ── Fase: Estrutura ── */}
          {aiPhase === 'structure' && (
            <Stack tokens={{ childrenGap: 12 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="Robot" styles={{ root: { fontSize: 16, color: '#0078d4' } }} />
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                  Estrutura do formulário com IA
                </Text>
                {fieldsLoading && <Spinner size={SpinnerSize.small} />}
                {!fieldsLoading && fieldMeta.length > 0 && (
                  <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                    {fieldMeta.length} campo(s)
                  </Text>
                )}
              </Stack>

              {!aiPreview && (
                <TextField
                  label="Descreva o formulário"
                  placeholder="Ex: Solicitação de férias onde o colaborador informa o período, tipo de ausência e gestor aprovador."
                  multiline
                  rows={3}
                  value={aiDescription}
                  onChange={(_, v) => setAiDescription(v ?? '')}
                  disabled={aiLoading || fieldsLoading}
                />
              )}

              {aiError && (
                <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setAiError(undefined)}>
                  {aiError}
                </MessageBar>
              )}

              {aiLoading && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Spinner size={SpinnerSize.small} />
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Gerando estrutura...
                  </Text>
                </Stack>
              )}

              {aiPreview && !aiLoading && (
                <Stack tokens={{ childrenGap: 6 }}>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    {visiblePreviewSteps.length} etapa(s) gerada(s) — revise e aplique:
                  </Text>
                  {visiblePreviewSteps.map((s) => (
                    <Stack
                      key={s.id}
                      styles={{
                        root: { border: '1px solid #edebe9', borderRadius: 4, padding: '6px 10px', background: '#f3f2f1' },
                      }}
                    >
                      <Text styles={{ root: { fontWeight: 600, fontSize: 13 } }}>{s.title}</Text>
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        {s.fieldNames.length === 0
                          ? 'Sem campos'
                          : s.fieldNames
                              .map((n) => {
                                const m = fieldMeta.find((f) => f.InternalName === n);
                                return m ? m.Title : n;
                              })
                              .join(' · ')}
                      </Text>
                    </Stack>
                  ))}
                </Stack>
              )}

              <Stack horizontal tokens={{ childrenGap: 8 }}>
                {!aiPreview && (
                  <DefaultButton
                    text="✨ Gerar estrutura"
                    onClick={() => void handleGenerateStructure()}
                    disabled={!aiDescription.trim() || aiLoading || fieldsLoading || !openAiApiKey}
                    title={!openAiApiKey ? 'Configure a chave OpenAI nas propriedades da web part' : undefined}
                  />
                )}
                {aiPreview && (
                  <>
                    <PrimaryButton text="Aplicar estrutura" onClick={handleApplyStructure} />
                    <DefaultButton
                      text="Tentar novamente"
                      onClick={() => { setAiPreview(null); setAiError(undefined); }}
                    />
                  </>
                )}
              </Stack>
            </Stack>
          )}

          {/* ── Fase: Botões ── */}
          {aiPhase === 'buttons' && (
            <Stack tokens={{ childrenGap: 12 }}>
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                <Icon iconName="ButtonControl" styles={{ root: { fontSize: 16, color: '#0078d4' } }} />
                <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                  Botões do formulário com IA
                </Text>
              </Stack>

              <MessageBar messageBarType={MessageBarType.success}>
                Estrutura aplicada com sucesso. Agora descreva os botões do formulário.
              </MessageBar>

              {!btnPreview && (
                <TextField
                  label="Descreva os botões"
                  placeholder={`Ex: Botão "Enviar" (principal, só no modo criar) que salva e fecha. Botão "Fechar" que redireciona para /sites/rh/paginas/ferias. Botão "Salvar Rascunho" (modo criar e editar).`}
                  multiline
                  rows={4}
                  value={btnDescription}
                  onChange={(_, v) => setBtnDescription(v ?? '')}
                  disabled={btnLoading}
                />
              )}

              {btnError && (
                <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setBtnError(undefined)}>
                  {btnError}
                </MessageBar>
              )}

              {btnLoading && (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
                  <Spinner size={SpinnerSize.small} />
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Gerando botões...
                  </Text>
                </Stack>
              )}

              {btnPreview && !btnLoading && (
                <Stack tokens={{ childrenGap: 6 }}>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    {btnPreview.length} botão(ões) gerado(s):
                  </Text>
                  {btnPreview.map((b) => (
                    <Stack
                      key={b.id}
                      horizontal
                      verticalAlign="center"
                      tokens={{ childrenGap: 8 }}
                      styles={{
                        root: { border: '1px solid #edebe9', borderRadius: 4, padding: '6px 10px', background: '#f3f2f1' },
                      }}
                    >
                      <Stack grow tokens={{ childrenGap: 2 }}>
                        <Text styles={{ root: { fontWeight: 600, fontSize: 13 } }}>
                          {b.label}
                          {b.appearance === 'primary' && (
                            <span style={{ marginLeft: 6, fontSize: 11, color: '#0078d4' }}>[principal]</span>
                          )}
                        </Text>
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                          {b.behavior ? BEHAVIOR_LABEL[b.behavior] ?? b.behavior : ''}
                          {b.operation === 'redirect' && b.redirectUrlTemplate
                            ? ` → ${b.redirectUrlTemplate}`
                            : ''}
                          {b.modes && b.modes.length > 0
                            ? ` · ${b.modes.map((m) => MODE_LABEL[m] ?? m).join(', ')}`
                            : ''}
                        </Text>
                      </Stack>
                    </Stack>
                  ))}
                </Stack>
              )}

              <Stack horizontal tokens={{ childrenGap: 8 }}>
                {!btnPreview && (
                  <DefaultButton
                    text="✨ Gerar botões"
                    onClick={() => void handleGenerateButtons()}
                    disabled={!btnDescription.trim() || btnLoading || !openAiApiKey}
                  />
                )}
                {btnPreview && (
                  <>
                    <PrimaryButton text="Aplicar botões" onClick={handleApplyButtons} />
                    <DefaultButton
                      text="Tentar novamente"
                      onClick={() => { setBtnPreview(null); setBtnError(undefined); }}
                    />
                  </>
                )}
                <DefaultButton
                  text="Pular"
                  onClick={() => { setBtnPreview(null); setBtnDescription(''); setBtnError(undefined); }}
                  title="Pular configuração de botões por IA"
                />
              </Stack>
            </Stack>
          )}
        </>
      )}
    </Stack>
  );
};
