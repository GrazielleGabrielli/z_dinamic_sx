import type { IListRowActionConfig } from './types';
import type { IFormCustomButtonConfig } from './types/formManager';

export const DEFAULT_FORM_MANAGER_NAV_BUTTONS: IFormCustomButtonConfig[] = [
  {
    id: 'dinamicSx_nav_editar',
    label: 'Editar',
    operation: 'redirect',
    behavior: 'actionsOnly',
    themePaletteSlot: 'themePrimary',
    customColorHex: '#0078D4',
    redirectUrlTemplate: '?Form=Edit&FormID={{FormID}}',
    actions: [],
    modes: ['view'],
  },
  {
    id: 'dinamicSx_nav_salvar',
    label: 'Salvar',
    appearance: 'primary',
    operation: 'add',
    behavior: 'actionsOnly',
    themePaletteSlot: 'themePrimary',
    customColorHex: '#107C10',
    actions: [],
    modes: ['create'],
  },
  {
    id: 'dinamicSx_nav_fechar',
    label: 'Fechar',
    operation: 'redirect',
    behavior: 'actionsOnly',
    themePaletteSlot: 'outline',
    customColorHex: '#605E5C',
    redirectUrlTemplate: '[siteurl]/sitepages/configure-o-destino-do-botao-fechar.aspx',
    actions: [],
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
