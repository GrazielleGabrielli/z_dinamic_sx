import * as React from 'react';
import { Stack, Text, Toggle, Dropdown, IDropdownOption, ChoiceGroup, IChoiceGroupOption } from '@fluentui/react';
import { IWizardFormState } from '../types';
import { TChartType } from '../../../core/config/types';
import { ChartTypeCard } from '../../Dashboard/ChartTypeCard';

interface IStep3Props {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
}

const CARDS_COUNT_OPTIONS: IDropdownOption[] = [1, 2, 3, 4, 5].map((n) => ({
  key: n,
  text: `${n} card${n > 1 ? 's' : ''}`,
}));

const DASHBOARD_TYPE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'cards', text: 'Cards' },
  { key: 'charts', text: 'Gráficos' },
];

const CHART_TYPES: TChartType[] = ['bar', 'line', 'area', 'pie', 'donut'];

export const Step3Dashboard: React.FC<IStep3Props> = ({ form, onChange }) => {
  const handleToggle = (_: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    onChange({ dashboardEnabled: !!checked });
  };

  const handleDashboardType = (_: React.FormEvent<HTMLElement | HTMLInputElement> | undefined, opt?: IChoiceGroupOption): void => {
    if (!opt) return;
    onChange({ dashboardType: opt.key as 'cards' | 'charts' });
  };

  const handleCardsCount = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    if (!opt) return;
    onChange({ cardsCount: opt.key as number });
  };

  const handleChartType = (type: TChartType): void => {
    onChange({ chartType: type });
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
        <ChoiceGroup
          label="Tipo de dashboard"
          options={DASHBOARD_TYPE_OPTIONS}
          selectedKey={form.dashboardType}
          onChange={handleDashboardType}
          styles={{ flexContainer: { display: 'flex', gap: 16 } }}
        />
      )}

      {form.dashboardEnabled && form.dashboardType === 'cards' && (
        <Dropdown
          label="Quantidade de cards"
          options={CARDS_COUNT_OPTIONS}
          selectedKey={form.cardsCount}
          onChange={handleCardsCount}
          styles={{ root: { maxWidth: 220 } }}
        />
      )}

      {form.dashboardEnabled && form.dashboardType === 'cards' && (
        <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
          O conteúdo de cada card será configurado após a conclusão do wizard.
        </Text>
      )}

      {form.dashboardEnabled && form.dashboardType === 'charts' && (
        <Stack tokens={{ childrenGap: 10 }}>
          <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
            Escolha o tipo de gráfico
          </Text>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 10 }}>
            {CHART_TYPES.map((type) => (
              <ChartTypeCard
                key={type}
                type={type}
                selected={form.chartType === type}
                onClick={handleChartType}
              />
            ))}
          </div>
          <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
            A configuração dos dados do gráfico será feita após a conclusão do wizard.
          </Text>
        </Stack>
      )}
    </Stack>
  );
};
