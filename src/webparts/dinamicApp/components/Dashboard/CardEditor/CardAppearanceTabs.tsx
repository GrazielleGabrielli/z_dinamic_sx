import * as React from 'react';
import { Pivot, PivotItem, Stack } from '@fluentui/react';
import type { IDashboardCardLayoutStyle } from '../../../core/config/types';
import { CardLayoutStyleFields } from './CardLayoutStyleFields';
import { CardTypographyStyleFields } from './CardTypographyStyleFields';
import { CardBehaviorLayoutFields } from './CardBehaviorLayoutFields';

export interface ICardAppearanceTabsProps {
  value: IDashboardCardLayoutStyle;
  onChange: (next: IDashboardCardLayoutStyle) => void;
}

export const CardAppearanceTabs: React.FC<ICardAppearanceTabsProps> = ({ value, onChange }) => (
  <Pivot styles={{ root: { marginTop: 4 } }}>
    <PivotItem headerText="Estilo" itemKey="estilo">
      <Stack styles={{ root: { paddingTop: 12 } }}>
        <CardLayoutStyleFields value={value} onChange={onChange} />
      </Stack>
    </PivotItem>
    <PivotItem headerText="Tipografia" itemKey="tipografia">
      <Stack styles={{ root: { paddingTop: 12 } }}>
        <CardTypographyStyleFields value={value} onChange={onChange} />
      </Stack>
    </PivotItem>
    <PivotItem headerText="Layout" itemKey="layout">
      <Stack styles={{ root: { paddingTop: 12 } }}>
        <CardBehaviorLayoutFields value={value} onChange={onChange} />
      </Stack>
    </PivotItem>
  </Pivot>
);
