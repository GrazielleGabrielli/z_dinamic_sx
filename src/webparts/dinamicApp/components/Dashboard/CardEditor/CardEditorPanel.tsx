import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  IconButton,
  PrimaryButton,
  DefaultButton,
  Separator,
} from '@fluentui/react';
import { IDashboardCardConfig } from '../../../core/config/types';
import { generateDefaultCards } from '../../../core/config/utils';
import { CardForm } from './CardForm';

interface ICardEditorPanelProps {
  isOpen: boolean;
  listTitle: string;
  cards: IDashboardCardConfig[];
  cardsCount: number;
  onSave: (cards: IDashboardCardConfig[]) => void;
  onDismiss: () => void;
}

type TPanelView = 'list' | 'form';

const AGGREGATE_LABEL: Record<string, string> = {
  count: 'contagem',
  sum: 'soma',
};

function initLocalCards(
  cards: IDashboardCardConfig[],
  cardsCount: number
): IDashboardCardConfig[] {
  return cards.length > 0 ? [...cards] : generateDefaultCards(cardsCount);
}

export const CardEditorPanel: React.FC<ICardEditorPanelProps> = ({
  isOpen,
  listTitle,
  cards,
  cardsCount,
  onSave,
  onDismiss,
}) => {
  const [localCards, setLocalCards] = useState<IDashboardCardConfig[]>(() =>
    initLocalCards(cards, cardsCount)
  );
  const [view, setView] = useState<TPanelView>('list');
  const [editingIndex, setEditingIndex] = useState<number | undefined>(undefined);

  useEffect(() => {
    if (isOpen) {
      setLocalCards(initLocalCards(cards, cardsCount));
      setView('list');
      setEditingIndex(undefined);
    }
  }, [isOpen]);

  const handleEdit = (index: number): void => {
    setEditingIndex(index);
    setView('form');
  };

  const handleAdd = (): void => {
    setEditingIndex(undefined);
    setView('form');
  };

  const handleDelete = (index: number): void => {
    setLocalCards((prev) => prev.filter((_, i) => i !== index));
  };

  const handleConfirmCard = (card: IDashboardCardConfig): void => {
    if (editingIndex !== undefined) {
      setLocalCards((prev) => prev.map((c, i) => (i === editingIndex ? card : c)));
    } else {
      setLocalCards((prev) => [...prev, card]);
    }
    setView('list');
  };

  const handleSave = (): void => {
    onSave(localCards);
  };

  const getEditingCard = (): IDashboardCardConfig | undefined => {
    if (editingIndex !== undefined && editingIndex < localCards.length) {
      return localCards[editingIndex];
    }
    return undefined;
  };

  const panelHeader =
    view === 'list'
      ? 'Editar cards do dashboard'
      : editingIndex !== undefined
      ? 'Editar card'
      : 'Novo card';

  const renderListView = (): React.ReactElement => (
    <Stack tokens={{ childrenGap: 0 }}>
      {localCards.length === 0 && (
        <Text variant="small" styles={{ root: { color: '#a19f9d', padding: '16px 0' } }}>
          Nenhum card configurado ainda.
        </Text>
      )}

      {localCards.map((card, index) => (
        <React.Fragment key={card.id}>
          <div
            style={{
              padding: '14px 0',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center',
            }}
          >
            <Stack tokens={{ childrenGap: 2 }}>
              <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
                {card.title}
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {AGGREGATE_LABEL[card.aggregate] ?? card.aggregate}
                {card.field !== undefined ? ` · campo: ${card.field}` : ''}
                {card.filter !== undefined
                  ? ` · filtro: ${card.filter.field} ${card.filter.operator} "${card.filter.value}"`
                  : ''}
              </Text>
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 2 }}>
              <IconButton
                iconProps={{ iconName: 'Edit' }}
                title="Editar"
                ariaLabel="Editar card"
                onClick={() => handleEdit(index)}
                styles={{ root: { color: '#0078d4' } }}
              />
              <IconButton
                iconProps={{ iconName: 'Delete' }}
                title="Remover"
                ariaLabel="Remover card"
                onClick={() => handleDelete(index)}
                styles={{ root: { color: '#d13438' } }}
              />
            </Stack>
          </div>
          {index < localCards.length - 1 && (
            <Separator styles={{ root: { padding: 0 } }} />
          )}
        </React.Fragment>
      ))}

      <div style={{ marginTop: 20 }}>
        <DefaultButton
          iconProps={{ iconName: 'Add' }}
          text="Adicionar card"
          onClick={handleAdd}
        />
      </div>
    </Stack>
  );

  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.medium}
      styles={{ main: { width: '85vw', maxWidth: '85vw' } }}
      headerText={panelHeader}
      closeButtonAriaLabel="Fechar"
      isFooterAtBottom={true}
      onRenderFooterContent={() => (
        <Stack
          horizontal
          tokens={{ childrenGap: 8 }}
          styles={{ root: { paddingBottom: 16 } }}
        >
          {view === 'list' && (
            <>
              <PrimaryButton text="Salvar" onClick={handleSave} />
              <DefaultButton text="Cancelar" onClick={onDismiss} />
            </>
          )}
        </Stack>
      )}
    >
      <div style={{ paddingTop: 16 }}>
        {view === 'list' && renderListView()}
        {view === 'form' && (
          <CardForm
            listTitle={listTitle}
            card={getEditingCard()}
            onConfirm={handleConfirmCard}
            onBack={() => setView('list')}
          />
        )}
      </div>
    </Panel>
  );
};
