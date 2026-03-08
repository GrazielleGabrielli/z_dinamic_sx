import * as React from 'react';
import { Stack, Text } from '@fluentui/react';
import { TViewMode } from '../../../core/config/types';
import { IWizardFormState } from '../types';

interface IStep2Props {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
}

interface IModeCard {
  key: TViewMode;
  title: string;
  description: string;
  enabled: boolean;
}

const MODE_CARDS: IModeCard[] = [
  {
    key: 'list',
    title: 'Modo Lista',
    description: 'Visualização em tabela com dashboard e paginação server-side.',
    enabled: true,
  },
  {
    key: 'projectManagement',
    title: 'Gestão de Projetos',
    description: 'Kanban, linha do tempo e gestão de tarefas.',
    enabled: false,
  },
  {
    key: 'formManager',
    title: 'Formulário + Gestor',
    description: 'Criação e gestão de formulários customizados.',
    enabled: false,
  },
];

const cardBase: React.CSSProperties = {
  padding: '16px 20px',
  borderRadius: 8,
  border: '2px solid',
  transition: 'all 0.15s ease',
  userSelect: 'none',
};

export const Step2Mode: React.FC<IStep2Props> = ({ form, onChange }) => {
  const handleSelect = (key: TViewMode, enabled: boolean): void => {
    if (!enabled) return;
    onChange({ mode: key });
  };

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          Modo da webpart
        </Text>
        <Text variant="medium" styles={{ root: { color: '#605e5c', marginTop: 4, display: 'block' } }}>
          Escolha como os dados serão apresentados.
        </Text>
      </Stack.Item>

      <Stack tokens={{ childrenGap: 10 }}>
        {MODE_CARDS.map((modeCard) => {
          const isSelected = form.mode === modeCard.key;
          return (
            <div
              key={modeCard.key}
              onClick={() => handleSelect(modeCard.key, modeCard.enabled)}
              style={{
                ...cardBase,
                borderColor: isSelected ? '#0078d4' : modeCard.enabled ? '#c8c6c4' : '#edebe9',
                background: isSelected ? '#f3f9ff' : modeCard.enabled ? '#fff' : '#faf9f8',
                opacity: modeCard.enabled ? 1 : 0.65,
                cursor: modeCard.enabled ? 'pointer' : 'not-allowed',
              }}
            >
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Stack tokens={{ childrenGap: 4 }}>
                  <Text
                    variant="large"
                    styles={{
                      root: { fontWeight: 600, color: modeCard.enabled ? '#201f1e' : '#a19f9d' },
                    }}
                  >
                    {modeCard.title}
                  </Text>
                  <Text
                    variant="small"
                    styles={{ root: { color: modeCard.enabled ? '#605e5c' : '#a19f9d' } }}
                  >
                    {modeCard.description}
                  </Text>
                </Stack>

                {!modeCard.enabled && (
                  <span
                    style={{
                      fontSize: 11,
                      fontWeight: 600,
                      color: '#605e5c',
                      background: '#edebe9',
                      padding: '2px 8px',
                      borderRadius: 12,
                      whiteSpace: 'nowrap',
                    }}
                  >
                    Em breve
                  </span>
                )}

                {isSelected && (
                  <span
                    style={{
                      fontSize: 11,
                      fontWeight: 600,
                      color: '#0078d4',
                      background: '#deecf9',
                      padding: '2px 8px',
                      borderRadius: 12,
                    }}
                  >
                    Selecionado
                  </span>
                )}
              </Stack>
            </div>
          );
        })}
      </Stack>
    </Stack>
  );
};
