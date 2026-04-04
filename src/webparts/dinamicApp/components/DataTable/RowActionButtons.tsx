import * as React from 'react';
import { Stack, IconButton } from '@fluentui/react';
import type { IListRowActionConfig } from '../../core/config/types';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { resolveListRowActionUrl, isSafeListRowNavigationUrl } from '../../core/table/utils/resolveListRowActionUrl';
import { listRowActionIconName } from './listRowActionUi';

export interface IRowActionButtonsProps {
  actions: IListRowActionConfig[];
  item: Record<string, unknown>;
  dynamicContext: IDynamicContext;
}

export const RowActionButtons: React.FC<IRowActionButtonsProps> = ({ actions, item, dynamicContext }) => {
  if (!actions.length) return null;

  return (
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 2 }} styles={{ root: { flexWrap: 'wrap' } }}>
      {actions.map((a) => {
        const href = resolveListRowActionUrl(a.urlTemplate, item, dynamicContext);
        if (!href || !isSafeListRowNavigationUrl(href)) return null;
        const icon = listRowActionIconName(a.iconPreset, a.customIconName);
        const newTab = a.openInNewTab === true;
        return (
          <IconButton
            key={a.id}
            iconProps={{ iconName: icon }}
            title={a.title}
            ariaLabel={a.title}
            href={href}
            target={newTab ? '_blank' : undefined}
            rel={newTab ? 'noopener noreferrer' : undefined}
            onClick={(ev) => { ev.stopPropagation(); }}
            styles={{ root: { width: 28, height: 28 }, icon: { fontSize: 14 } }}
          />
        );
      })}
    </Stack>
  );
};
