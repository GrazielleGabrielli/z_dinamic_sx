import * as React from 'react';
import { useMemo } from 'react';
import { Dropdown, IDropdownOption, Stack, useTheme, type ITheme } from '@fluentui/react';
import type { TFormCustomButtonPaletteSlot } from '../../core/config/types/formManager';
import {
  FORM_CUSTOM_BUTTON_SLOT_LABELS,
  FORM_CUSTOM_BUTTON_THEME_SLOTS,
  paletteBgFromSlot,
} from '../../core/formManager/formCustomButtonTheme';

export function ThemePaletteSlotDropdownRow(props: { option: IDropdownOption; theme: ITheme }): JSX.Element {
  const { option, theme } = props;
  const p = theme.palette;
  const isOutline = String(option.key) === 'outline';
  const swatch = (option.data as { swatch?: string } | undefined)?.swatch;
  return (
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { minHeight: 26 } }}>
      <div
        title={swatch}
        style={{
          width: 14,
          height: 14,
          borderRadius: 2,
          flexShrink: 0,
          boxSizing: 'border-box',
          border: `1px solid ${p.neutralQuaternaryAlt}`,
          backgroundColor: isOutline ? p.white : swatch ?? p.neutralLight,
        }}
      />
      <span>{option.text}</span>
    </Stack>
  );
}

export function useThemePaletteSlotDropdownOptions(): IDropdownOption[] {
  const theme = useTheme();
  return useMemo((): IDropdownOption[] => {
    const outline: IDropdownOption = { key: 'outline', text: FORM_CUSTOM_BUTTON_SLOT_LABELS.outline };
    const colorOpts: IDropdownOption[] = FORM_CUSTOM_BUTTON_THEME_SLOTS.map((slot) => ({
      key: slot,
      text: FORM_CUSTOM_BUTTON_SLOT_LABELS[slot],
      data: { swatch: paletteBgFromSlot(theme, slot) },
    }));
    return [outline, ...colorOpts];
  }, [theme]);
}

export const ThemePaletteSlotDropdown: React.FC<{
  label: string;
  selectedKey: TFormCustomButtonPaletteSlot;
  onChange: (slot: TFormCustomButtonPaletteSlot) => void;
}> = ({ label, selectedKey, onChange }) => {
  const theme = useTheme();
  const options = useThemePaletteSlotDropdownOptions();
  return (
    <Dropdown
      label={label}
      options={options}
      selectedKey={selectedKey}
      onRenderOption={(option) =>
        option ? <ThemePaletteSlotDropdownRow option={option} theme={theme} /> : null
      }
      onRenderTitle={(selected) =>
        selected && selected.length > 0 ? (
          <ThemePaletteSlotDropdownRow option={selected[0]} theme={theme} />
        ) : null
      }
      onChange={(_, o) => {
        if (!o) return;
        onChange(String(o.key) as TFormCustomButtonPaletteSlot);
      }}
    />
  );
};
