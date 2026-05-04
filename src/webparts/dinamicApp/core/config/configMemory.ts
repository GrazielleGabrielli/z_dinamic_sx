import type {
  IDynamicViewConfig,
  IDataSourceConfig,
  IModeConfigSnapshot,
  IConfigMemory,
  TViewMode,
} from './types';
import type { IWizardFormState } from '../../components/Wizard/types';
import { configToWizardState } from '../../components/Wizard/types';
import { getDefaultConfig, getDefaultFormManagerConfig } from './utils';
import { buildConfig } from './builders';

export function sourceKey(ds: IDataSourceConfig): string {
  const w = (ds.webServerRelativeUrl ?? '').trim();
  return `${w}::${ds.kind}::${ds.title.trim()}`;
}

export function cloneJson<T>(x: T): T {
  return JSON.parse(JSON.stringify(x)) as T;
}

export function emptyConfigMemory(): IConfigMemory {
  return { bySource: {} };
}

export function captureFullSnapshot(c: IDynamicViewConfig): IModeConfigSnapshot {
  return {
    listView: cloneJson(c.listView),
    ...(c.projectManagement !== undefined && { projectManagement: cloneJson(c.projectManagement) }),
    ...(c.formManager !== undefined && { formManager: cloneJson(c.formManager) }),
    dashboard: cloneJson(c.dashboard),
    pagination: cloneJson(c.pagination),
    ...(c.listPageLayout !== undefined && { listPageLayout: cloneJson(c.listPageLayout) }),
    ...(c.pdfTemplate !== undefined && { pdfTemplate: cloneJson(c.pdfTemplate) }),
    ...(c.tableConfig !== undefined && { tableConfig: cloneJson(c.tableConfig) }),
  };
}

export function applySnapshotToConfig(
  base: IDynamicViewConfig,
  snap: IModeConfigSnapshot | undefined
): IDynamicViewConfig {
  if (!snap) return base;
  return {
    ...base,
    ...(snap.listView !== undefined && { listView: snap.listView }),
    ...(snap.projectManagement !== undefined && { projectManagement: snap.projectManagement }),
    ...(snap.formManager !== undefined && { formManager: snap.formManager }),
    ...(snap.dashboard !== undefined && { dashboard: snap.dashboard }),
    ...(snap.pagination !== undefined && { pagination: snap.pagination }),
    ...(snap.listPageLayout !== undefined && { listPageLayout: snap.listPageLayout }),
    ...(snap.pdfTemplate !== undefined && { pdfTemplate: snap.pdfTemplate }),
    ...(snap.tableConfig !== undefined && { tableConfig: snap.tableConfig }),
  };
}

export function upsertConfigMemoryForListSource(
  base: IDynamicViewConfig,
  childDataSource: IDataSourceConfig,
  partial: Partial<Pick<IModeConfigSnapshot, 'listView' | 'pagination' | 'tableConfig' | 'pdfTemplate'>>
): IDynamicViewConfig {
  const memory = base.configMemory ? cloneJson(base.configMemory) : emptyConfigMemory();
  if (!memory.bySource) memory.bySource = {};
  const key = sourceKey(childDataSource);
  const prevEntry = memory.bySource[key] ?? {};
  const prevListSnap: Partial<IModeConfigSnapshot> = { ...(prevEntry.list ?? {}) };
  memory.bySource[key] = {
    ...prevEntry,
    list: { ...prevListSnap, ...partial } as IModeConfigSnapshot,
  };
  return { ...base, configMemory: memory };
}

function mergeWizardFormIntoConfig(
  prev: IDynamicViewConfig,
  mergedForm: IWizardFormState,
  nextDs: IDataSourceConfig,
  nextMode: TViewMode
): IDynamicViewConfig {
  const built = buildConfig({
    dataSource: nextDs,
    mode: nextMode,
    dashboard: {
      enabled: mergedForm.dashboardEnabled,
      dashboardType: mergedForm.dashboardType,
      cardsCount: mergedForm.cardsCount,
      cards: prev.dashboard.cards,
      chartType: mergedForm.chartType,
      chartSeries: prev.dashboard.chartSeries ?? [],
    },
    pagination: {
      enabled: mergedForm.paginationEnabled,
      pageSize: mergedForm.pageSize,
      pageSizeOptions: mergedForm.pageSizeOptions,
    },
    listView: {
      ...prev.listView,
      viewModes: mergedForm.viewModes,
      activeViewModeId: mergedForm.activeViewModeId,
      ...(mergedForm.viewModePicker === 'tabs' ? { viewModePicker: 'tabs' as const } : { viewModePicker: undefined }),
      ...(mergedForm.viewModeDefaultRules && mergedForm.viewModeDefaultRules.length > 0
        ? { viewModeDefaultRules: mergedForm.viewModeDefaultRules }
        : { viewModeDefaultRules: undefined }),
    },
    projectManagement: prev.projectManagement,
    formManager:
      nextMode === 'formManager'
        ? { ...(prev.formManager ?? getDefaultFormManagerConfig()), stepLayout: mergedForm.formStepLayout }
        : prev.formManager,
  });
  return {
    ...built,
    ...(prev.listPageLayout !== undefined && { listPageLayout: prev.listPageLayout }),
    ...(prev.pdfTemplate !== undefined && { pdfTemplate: prev.pdfTemplate }),
    ...(prev.tableConfig !== undefined && { tableConfig: prev.tableConfig }),
  };
}

export function applyWizardPartial(
  prev: IDynamicViewConfig,
  partial: Partial<IWizardFormState>
): IDynamicViewConfig {
  const mergedForm: IWizardFormState = {
    ...configToWizardState(prev),
    ...partial,
  };
  const wTrim = (mergedForm.dataSourceWebServerRelativeUrl ?? '').trim();
  const nextDs: IDataSourceConfig = {
    kind: mergedForm.kind,
    title: mergedForm.title,
    ...(wTrim ? { webServerRelativeUrl: wTrim } : {}),
  };
  const nextMode = mergedForm.mode;
  const prevKey = sourceKey(prev.dataSource);
  const nextKey = sourceKey(nextDs);
  const normWeb = (v?: string): string => (v ?? '').trim();
  const transition =
    (partial.mode !== undefined && partial.mode !== prev.mode) ||
    (partial.title !== undefined && partial.title.trim() !== prev.dataSource.title.trim()) ||
    (partial.kind !== undefined && partial.kind !== prev.dataSource.kind) ||
    (partial.dataSourceWebServerRelativeUrl !== undefined &&
      normWeb(partial.dataSourceWebServerRelativeUrl) !== normWeb(prev.dataSource.webServerRelativeUrl));

  let memory: IConfigMemory = prev.configMemory ? cloneJson(prev.configMemory) : emptyConfigMemory();
  if (!memory.bySource) memory.bySource = {};

  if (transition) {
    if (!memory.bySource[prevKey]) memory.bySource[prevKey] = {};
    const byMode = memory.bySource[prevKey] as Record<string, IModeConfigSnapshot>;
    byMode[prev.mode] = captureFullSnapshot(prev);

    let working = mergeWizardFormIntoConfig(prev, mergedForm, nextDs, nextMode);
    const snap = memory.bySource[nextKey]?.[nextMode];
    if (snap !== undefined) {
      working = applySnapshotToConfig(working, snap);
    } else if (nextKey !== prevKey) {
      const defaults = getDefaultConfig();
      working = {
        ...working,
        dataSource: nextDs,
        mode: nextMode,
        listView: defaults.listView,
        projectManagement: defaults.projectManagement,
        listPageLayout: undefined,
        pdfTemplate: undefined,
        tableConfig: undefined,
        formManager: undefined,
        dashboard: {
          ...defaults.dashboard,
          enabled: mergedForm.dashboardEnabled,
          dashboardType: mergedForm.dashboardType,
          cardsCount: mergedForm.cardsCount,
          chartType: mergedForm.chartType,
        },
      };
    }
    return { ...working, dataSource: nextDs, mode: nextMode, configMemory: memory };
  }

  const next = mergeWizardFormIntoConfig(prev, mergedForm, nextDs, nextMode);
  return { ...next, dataSource: nextDs, mode: nextMode, configMemory: memory };
}
