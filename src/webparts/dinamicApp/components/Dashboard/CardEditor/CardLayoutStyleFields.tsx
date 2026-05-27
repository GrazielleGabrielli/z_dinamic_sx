import * as React from 'react';
import { Stack, Text, Dropdown, IDropdownOption, Toggle } from '@fluentui/react';
import type {
  IDashboardCardLayoutStyle,
  TBorderRadius,
  TCardVariant,
  TPadding,
  TShadow,
} from '../../../core/config/types';
import {
  CARD_BORDER_RADIUS_OPTIONS,
  CARD_PADDING_OPTIONS,
  CARD_SHADOW_OPTIONS,
  CARD_VARIANT_OPTIONS,
} from './dashboardCardStyleOptions';

export interface ICardLayoutStyleFieldsProps {
  value: IDashboardCardLayoutStyle;
  onChange: (next: IDashboardCardLayoutStyle) => void;
}

export const CardLayoutStyleFields: React.FC<ICardLayoutStyleFieldsProps> = ({ value, onChange }) => {
  const patch = (partial: Partial<IDashboardCardLayoutStyle>): void => {
    onChange({ ...value, ...partial });
  };

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Aplica-se a todos os cards deste dashboard. Cores e ícone continuam por card.
      </Text>
      <Dropdown
        label="Variante"
        options={CARD_VARIANT_OPTIONS}
        selectedKey={value.variant}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ variant: opt.key as TCardVariant });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Dropdown
        label="Borda (cantos)"
        options={CARD_BORDER_RADIUS_OPTIONS}
        selectedKey={value.borderRadius}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ borderRadius: opt.key as TBorderRadius });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Dropdown
        label="Preenchimento"
        options={CARD_PADDING_OPTIONS}
        selectedKey={value.padding}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ padding: opt.key as TPadding });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Dropdown
        label="Sombra"
        options={CARD_SHADOW_OPTIONS}
        selectedKey={value.shadow}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ shadow: opt.key as TShadow });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Toggle
        label="Exibir borda"
        checked={value.border}
        onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) => patch({ border: !!checked })}
      />
    </Stack>
  );
};
