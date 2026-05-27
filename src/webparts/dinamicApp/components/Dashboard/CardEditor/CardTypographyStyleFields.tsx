import * as React from 'react';
import { Stack, Dropdown, IDropdownOption } from '@fluentui/react';
import type {
  IDashboardCardLayoutStyle,
  TFontWeight,
  TSubtitleSize,
  TTitleSize,
  TValueSize,
} from '../../../core/config/types';
import {
  CARD_FONT_WEIGHT_OPTIONS,
  CARD_SUBTITLE_SIZE_OPTIONS,
  CARD_TITLE_SIZE_OPTIONS,
  CARD_VALUE_SIZE_OPTIONS,
} from './dashboardCardStyleOptions';

export interface ICardTypographyStyleFieldsProps {
  value: IDashboardCardLayoutStyle;
  onChange: (next: IDashboardCardLayoutStyle) => void;
}

export const CardTypographyStyleFields: React.FC<ICardTypographyStyleFieldsProps> = ({
  value,
  onChange,
}) => {
  const patch = (partial: Partial<IDashboardCardLayoutStyle>): void => {
    onChange({ ...value, ...partial });
  };

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Dropdown
        label="Tamanho do título"
        options={CARD_TITLE_SIZE_OPTIONS}
        selectedKey={value.titleSize}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ titleSize: opt.key as TTitleSize });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Dropdown
        label="Tamanho do subtítulo"
        options={CARD_SUBTITLE_SIZE_OPTIONS}
        selectedKey={value.subtitleSize}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ subtitleSize: opt.key as TSubtitleSize });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Dropdown
        label="Tamanho do valor"
        options={CARD_VALUE_SIZE_OPTIONS}
        selectedKey={value.valueSize}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ valueSize: opt.key as TValueSize });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Dropdown
        label="Peso do título"
        options={CARD_FONT_WEIGHT_OPTIONS}
        selectedKey={value.titleWeight}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ titleWeight: opt.key as TFontWeight });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
      <Dropdown
        label="Peso do valor"
        options={CARD_FONT_WEIGHT_OPTIONS}
        selectedKey={value.valueWeight}
        onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
          if (opt) patch({ valueWeight: opt.key as TFontWeight });
        }}
        styles={{ root: { maxWidth: 220 } }}
      />
    </Stack>
  );
};
