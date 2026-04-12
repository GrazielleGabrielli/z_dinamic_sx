import * as React from 'react';
import { useState } from 'react';
import {
  Modal,
  Stack,
  Text,
  TextField,
  PrimaryButton,
  DefaultButton,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
  Icon,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../../services/shared/types';
import type { IAIStepConfig, IAIStructureOutput } from '../../../../../services/ai/AIConfigService';
import { generateFormStructure } from '../../../../../services/ai/AIConfigService';

export interface IAIStructureModalProps {
  isOpen: boolean;
  listTitle: string;
  meta: IFieldMetadata[];
  openAiApiKey: string;
  onApply: (steps: IAIStepConfig[]) => void;
  onDismiss: () => void;
}

export const AIStructureModal: React.FC<IAIStructureModalProps> = ({
  isOpen,
  listTitle,
  meta,
  openAiApiKey,
  onApply,
  onDismiss,
}) => {
  const [description, setDescription] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const [preview, setPreview] = useState<IAIStructureOutput | null>(null);

  const handleDismiss = (): void => {
    if (loading) return;
    setDescription('');
    setError(undefined);
    setPreview(null);
    onDismiss();
  };

  const handleGenerate = async (): Promise<void> => {
    const desc = description.trim();
    if (!desc) {
      setError('Descreva o sistema antes de gerar.');
      return;
    }
    if (!openAiApiKey) {
      setError('Configure a chave OpenAI nas propriedades da web part antes de usar esta função.');
      return;
    }
    setError(undefined);
    setPreview(null);
    setLoading(true);
    try {
      const result = await generateFormStructure(openAiApiKey, {
        description: desc,
        listTitle,
        fields: meta,
      });
      setPreview(result);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  };

  const handleApply = (): void => {
    if (!preview) return;
    onApply(preview.steps);
    setDescription('');
    setError(undefined);
    setPreview(null);
  };

  const handleRetry = (): void => {
    setPreview(null);
    setError(undefined);
  };

  const visibleSteps = preview?.steps.filter(
    (s) => s.id !== 'ocultos' && s.id !== 'fixos'
  ) ?? [];

  return (
    <Modal
      isOpen={isOpen}
      onDismiss={handleDismiss}
      isBlocking={loading}
      styles={{ main: { width: 560, maxWidth: '95vw', padding: 24 } }}
    >
      <Stack tokens={{ childrenGap: 16 }}>
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          <Icon iconName="Robot" styles={{ root: { fontSize: 20, color: '#0078d4' } }} />
          <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
            Criar estrutura com IA
          </Text>
        </Stack>

        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Descreva brevemente o sistema. A IA irá organizar os campos da lista{' '}
          <strong>"{listTitle}"</strong> ({meta.length} campo(s)) em etapas lógicas.
        </Text>

        {!preview && (
          <TextField
            label="Descrição do sistema"
            placeholder="Ex: Formulário de solicitação de férias onde o colaborador informa o período e o tipo de ausência para aprovação do gestor."
            multiline
            rows={4}
            value={description}
            onChange={(_, v) => setDescription(v ?? '')}
            disabled={loading}
            required
          />
        )}

        {error && (
          <MessageBar messageBarType={MessageBarType.error} onDismiss={() => setError(undefined)}>
            {error}
          </MessageBar>
        )}

        {loading && (
          <Stack horizontalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { padding: '16px 0' } }}>
            <Spinner size={SpinnerSize.large} />
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Gerando estrutura...
            </Text>
          </Stack>
        )}

        {preview && !loading && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
              Estrutura gerada — {visibleSteps.length} etapa(s)
            </Text>
            {visibleSteps.map((step) => (
              <Stack
                key={step.id}
                styles={{
                  root: {
                    border: '1px solid #edebe9',
                    borderRadius: 4,
                    padding: '8px 12px',
                    background: '#faf9f8',
                  },
                }}
                tokens={{ childrenGap: 4 }}
              >
                <Text styles={{ root: { fontWeight: 600 } }}>{step.title}</Text>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  {step.fieldNames.length === 0
                    ? 'Sem campos'
                    : step.fieldNames.map((n) => {
                        const m = meta.find((f) => f.InternalName === n);
                        return m ? `${m.Title} (${n})` : n;
                      }).join(' · ')}
                </Text>
              </Stack>
            ))}
          </Stack>
        )}

        <Stack horizontal tokens={{ childrenGap: 8 }} horizontalAlign="end">
          {!preview && !loading && (
            <>
              <PrimaryButton
                text="Gerar estrutura"
                onClick={() => void handleGenerate()}
                disabled={!description.trim()}
              />
              <DefaultButton text="Cancelar" onClick={handleDismiss} />
            </>
          )}
          {preview && !loading && (
            <>
              <PrimaryButton text="Aplicar" onClick={handleApply} />
              <DefaultButton text="Tentar novamente" onClick={handleRetry} />
              <DefaultButton text="Cancelar" onClick={handleDismiss} />
            </>
          )}
        </Stack>
      </Stack>
    </Modal>
  );
};
