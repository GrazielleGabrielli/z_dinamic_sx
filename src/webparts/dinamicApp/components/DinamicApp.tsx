import * as React from 'react';
import { useState } from 'react';
import { Text, Stack, Separator, ActionButton } from '@fluentui/react';
import type { IDinamicAppProps } from './IDinamicAppProps';
import { parseConfig } from '../core/config/validators';
import { IDashboardCardConfig, IDynamicViewConfig } from '../core/config/types';
import { ConfigWizard } from './Wizard/ConfigWizard';
import { DashboardView } from './Dashboard/DashboardView';
import { CardEditorPanel } from './Dashboard/CardEditor/CardEditorPanel';

const DinamicApp: React.FC<IDinamicAppProps> = ({ configJson, siteUrl, onSaveConfig }) => {
  const [isEditingWebPart, setIsEditingWebPart] = useState(false);
  const [isEditingCards, setIsEditingCards] = useState(false);

  const config = parseConfig(configJson ?? undefined);

  const handleWizardComplete = (newConfig: IDynamicViewConfig): void => {
    onSaveConfig(newConfig);
    setIsEditingWebPart(false);
  };

  const handleSaveCards = (cards: IDashboardCardConfig[]): void => {
    if (!config) return;
    onSaveConfig({
      ...config,
      dashboard: {
        ...config.dashboard,
        cards,
        cardsCount: cards.length,
      },
    });
    setIsEditingCards(false);
  };

  if (config === undefined || isEditingWebPart) {
    return (
      <ConfigWizard
        siteUrl={siteUrl}
        onComplete={handleWizardComplete}
        initialValues={config}
        onCancel={config !== undefined ? () => setIsEditingWebPart(false) : undefined}
      />
    );
  }

  const showDashboard = config.dashboard.enabled && config.dashboard.cardsCount > 0;

  return (
    <>
      {/* Toolbar */}
      <div
        style={{
          display: 'flex',
          justifyContent: 'flex-end',
          padding: '6px 16px 0',
          borderBottom: '1px solid #f3f2f1',
        }}
      >
        <ActionButton
          iconProps={{ iconName: 'Settings' }}
          onClick={() => setIsEditingWebPart(true)}
          styles={{ root: { color: '#605e5c', fontSize: 12 } }}
        >
          Editar configuração
        </ActionButton>
      </div>

      <Stack styles={{ root: { padding: '20px 24px 0' } }}>
        {showDashboard && (
          <>
            <DashboardView
              config={config.dashboard}
              dataSource={config.dataSource}
              onEditCards={() => setIsEditingCards(true)}
            />
            <Separator />
          </>
        )}

        <Stack tokens={{ childrenGap: 6 }} styles={{ root: { padding: '16px 0' } }}>
          <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
            {config.dataSource.title}
          </Text>
          <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
            Modo: {config.mode} · Origem: {config.dataSource.kind}
          </Text>
          <Text variant="small" styles={{ root: { color: '#c8c6c4' } }}>
            Listagem e paginação serão implementadas na próxima etapa.
          </Text>
        </Stack>
      </Stack>

      <CardEditorPanel
        isOpen={isEditingCards}
        listTitle={config.dataSource.title}
        cards={config.dashboard.cards}
        cardsCount={config.dashboard.cardsCount}
        onSave={handleSaveCards}
        onDismiss={() => setIsEditingCards(false)}
      />
    </>
  );
};

export default DinamicApp;
