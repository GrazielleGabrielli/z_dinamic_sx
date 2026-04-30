import {
  IDynamicViewConfig,
  IDashboardCardConfig,
  TAggregateType,
  IFormManagerConfig,
  TViewMode,
} from '../types';
import {
  FORM_FIXOS_STEP_ID,
  FORM_OCULTOS_STEP_ID,
  type IFormLinkedChildFormConfig,
} from '../types/formManager';
import { getDefaultDashboardCardStyle } from '../../dashboard/utils';
import {
  DEFAULT_FORM_MANAGER_NAV_BUTTONS,
  DEFAULT_LIST_VIEW_ROW_NAV_ACTIONS,
} from '../defaultFormNavButtons';

export function getDefaultFormManagerConfig(): IFormManagerConfig {
  return {
    sections: [
      { id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', visible: true },
      { id: FORM_FIXOS_STEP_ID, title: 'Fixos', visible: true },
      { id: 'main', title: 'Geral', visible: true },
    ],
    fields: [],
    rules: [],
    steps: [
      { id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] },
      { id: FORM_FIXOS_STEP_ID, title: 'Fixos', fieldNames: [] },
      { id: 'main', title: 'Geral', fieldNames: [] },
    ],
    stepLayout: 'segmented',
    linkedChildForms: [],
    customButtons: DEFAULT_FORM_MANAGER_NAV_BUTTONS.map((b) => ({ ...b, actions: [...b.actions] })),
  };
}

export function newLinkedChildFormConfig(id: string): IFormLinkedChildFormConfig {
  return {
    id,
    listTitle: '',
    parentLookupFieldInternalName: '',
    sections: [
      { id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', visible: true },
      { id: FORM_FIXOS_STEP_ID, title: 'Fixos', visible: true },
      { id: 'main', title: 'Geral', visible: true },
    ],
    fields: [],
    rules: [],
    steps: [
      { id: FORM_OCULTOS_STEP_ID, title: 'Ocultos', fieldNames: [] },
      { id: FORM_FIXOS_STEP_ID, title: 'Fixos', fieldNames: [] },
      { id: 'main', title: 'Geral', fieldNames: [] },
    ],
  };
}

export function getDefaultConfigForMode(mode: TViewMode): IDynamicViewConfig {
  const base = getDefaultConfig();
  return {
    ...base,
    mode,
    ...(mode === 'formManager' ? { formManager: getDefaultFormManagerConfig() } : {}),
  };
}

export function getDefaultConfig(): IDynamicViewConfig {
  return {
    dataSource: {
      kind: 'list',
      title: '',
    },
    mode: 'list',
    dashboard: {
      enabled: false,
      dashboardType: 'cards',
      cardsCount: 0,
      cards: [],
      chartType: 'bar',
    },
    pagination: {
      enabled: true,
      pageSize: 10,
      pageSizeOptions: [5, 10, 20, 50, 100],
      layout: 'buttons',
    },
    listView: {
      columns: [{ field: 'Title' }],
      filters: [],
      sort: null,
      viewModes: [
        { id: 'all', label: 'Todas', filters: [] },
        { id: 'mine', label: 'Minhas', filters: [{ field: 'Author/Id', operator: 'eq', value: '[Me]' }] },
      ],
      activeViewModeId: 'all',
      listCardViewEnabled: false,
      listRowActions: DEFAULT_LIST_VIEW_ROW_NAV_ACTIONS.map((a) => ({ ...a })),
    },
    projectManagement: {
      columns: [],
    },
  };
}

export function isConfigured(config: IDynamicViewConfig): boolean {
  return config.dataSource.title.trim().length > 0;
}

export function generateDefaultCards(count: number): IDashboardCardConfig[] {
  const defaultStyle = getDefaultDashboardCardStyle();
  const cards: IDashboardCardConfig[] = [];
  for (let i = 0; i < count; i++) {
    cards.push({
      id: `card_${i + 1}`,
      title: `Card ${i + 1}`,
      subtitle: '',
      aggregate: 'count' as TAggregateType,
      emptyValueText: 'Nenhum item',
      errorText: 'Erro ao carregar',
      loadingText: 'Carregando...',
      style: { ...defaultStyle },
    });
  }
  return cards;
}
