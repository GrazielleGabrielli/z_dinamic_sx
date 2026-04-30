import type { IListRowActionConfig } from './types';
import type { IFormCustomButtonConfig } from './types/formManager';

export const DEFAULT_FORM_MANAGER_NAV_BUTTONS: IFormCustomButtonConfig[] = [
  {
    id: 'dinamicSx_nav_fechar',
    label: 'Fechar',
    operation: 'redirect',
    behavior: 'actionsOnly',
    redirectUrlTemplate: '[siteurl]/sitepages/pagina_view.aspx',
    actions: [],
  },
  {
    id: 'dinamicSx_nav_editar',
    label: 'Editar',
    operation: 'redirect',
    behavior: 'actionsOnly',
    redirectUrlTemplate: '[siteurl]/sitepages/pagina_view.aspx?Form=Edit&FormID={{FormID}}',
    actions: [],
    modes: ['view'],
  },
  {
    id: 'dinamicSx_nav_ver',
    label: 'Ver',
    operation: 'redirect',
    behavior: 'actionsOnly',
    redirectUrlTemplate: '[siteurl]/sitepages/pagina_view.aspx?Form=Disp&FormID={{FormID}}',
    actions: [],
    modes: ['view'],
  },
];

export const DEFAULT_LIST_VIEW_ROW_NAV_ACTIONS: IListRowActionConfig[] = [
  {
    id: 'dinamicSx_row_fechar',
    title: 'Fechar',
    iconPreset: 'link',
    urlTemplate: '[siteurl]/sitepages/pagina_view.aspx',
    scope: 'icon',
  },
  {
    id: 'dinamicSx_row_editar',
    title: 'Editar',
    iconPreset: 'edit',
    urlTemplate: '[siteurl]/sitepages/pagina_view.aspx?Form=Edit&FormID={{ID}}',
    scope: 'icon',
  },
  {
    id: 'dinamicSx_row_ver',
    title: 'Ver',
    iconPreset: 'view',
    urlTemplate: '[siteurl]/sitepages/pagina_view.aspx?Form=Disp&FormID={{ID}}',
    scope: 'icon',
  },
];
