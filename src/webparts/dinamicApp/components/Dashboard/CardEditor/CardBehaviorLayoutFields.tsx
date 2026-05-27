import * as React from 'react';
import { Stack, Dropdown, IDropdownOption, Toggle } from '@fluentui/react';
import type {
  IDashboardCardLayoutStyle,
  TAlign,
  TIconPosition,
  TLoadingStyle,
} from '../../../core/config/types';
import {
  CARD_ALIGN_OPTIONS,
  CARD_ICON_POSITION_OPTIONS,
  CARD_LOADING_STYLE_OPTIONS,
} from './dashboardCardStyleOptions';

export interface ICardBehaviorLayoutFieldsProps {
  value: IDashboardCardLayoutStyle;
  onChange: (next: IDashboardCardLayoutStyle) => void;
}

export const CardBehaviorLayoutFields: React.FC<ICardBehaviorLayoutFieldsProps> = ({
  value,
  onChange,
}) => {
  const patch = (partial: Partial<IDashboardCardLayoutStyle>): void => {
    onChange({ ...value, ...partial });
  };

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Dropdown
        label="Alinhamento"
        options={CARD_ALIGN_OPTIONS}
        selectedKey={value.align}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ align: opt.key as TAlign });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Toggle
        label="Exibir subtítulo"
        checked={value.showSubtitle}
        onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
          patch({ showSubtitle: !!checked })
        }
      />
      <Toggle
        label="Exibir valor"
        checked={value.showValue}
        onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
          patch({ showValue: !!checked })
        }
      />
      <Toggle
        label="Exibir ícone"
        checked={value.showIcon}
        onChange={(_: React.MouseEvent<HTMLElement>, checked?: boolean) =>
          patch({ showIcon: !!checked })
        }
      />
      {value.showIcon && (
        <Dropdown
          label="Posição do ícone"
          options={CARD_ICON_POSITION_OPTIONS}
          selectedKey={value.iconPosition}
          onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
            if (opt) patch({ iconPosition: opt.key as TIconPosition });
          }}
          styles={{ root: { maxWidth: 220 } }}
        />
      )}
      <Dropdown
        label="Estilo de carregamento"
        options={CARD_LOADING_STYLE_OPTIONS}
        selectedKey={value.loadingStyle}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ loadingStyle: opt.key as TLoadingStyle });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
    </Stack>
  );
};
