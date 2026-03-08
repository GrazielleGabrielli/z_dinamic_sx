import * as React from 'react';
import { Stack, Text, Toggle, Dropdown, IDropdownOption } from '@fluentui/react';
import { IWizardFormState } from '../types';

interface IStep3Props {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
}

const CARDS_COUNT_OPTIONS: IDropdownOption[] = [1, 2, 3, 4, 5].map((n) => ({
  key: n,
  text: `${n} card${n > 1 ? 's' : ''}`,
}));

export const Step3Dashboard: React.FC<IStep3Props> = ({ form, onChange }) => {
  const handleToggle = (_: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    onChange({ dashboardEnabled: !!checked });
  };

  const handleCardsCount = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    if (!opt) return;
    onChange({ cardsCount: opt.key as number });
  };

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          Dashboard
        </Text>
        <Text variant="medium" styles={{ root: { color: '#605e5c', marginTop: 4, display: 'block' } }}>
          O dashboard exibe métricas e totalizadores no topo da webpart.
        </Text>
      </Stack.Item>

      <Toggle
        label="Habilitar dashboard"
        checked={form.dashboardEnabled}
        onChange={handleToggle}
        onText="Sim"
        offText="Não"
      />

      {form.dashboardEnabled && (
        <Dropdown
          label="Quantidade de cards"
          options={CARDS_COUNT_OPTIONS}
          selectedKey={form.cardsCount}
          onChange={handleCardsCount}
          styles={{ root: { maxWidth: 220 } }}
        />
      )}

      {form.dashboardEnabled && (
        <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
          O conteúdo de cada card será configurado após a conclusão do wizard.
        </Text>
      )}
    </Stack>
  );
};
