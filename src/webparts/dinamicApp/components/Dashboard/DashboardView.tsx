import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import { Stack, ActionButton, MessageBar, MessageBarType, Dropdown, IDropdownOption } from '@fluentui/react';
import { FieldsService, UsersService } from '../../../../services';
import {
  IChartSeriesConfig,
  IDashboardCardConfig,
  IDashboardConfig,
  IDataSourceConfig,
  IDynamicViewConfig,
  IListPageBlock,
  IListPageSection,
} from '../../core/config/types';
import { generateDefaultCards } from '../../core/config/utils';
import { DashboardEngine } from '../../core/dashboard/DashboardEngine';
import { IDashboardCardResult } from '../../core/dashboard/types';
import { buildDynamicContext, parseQueryString } from '../../core/dynamicTokens';
import {
  buildLinkedViewModeOData,
  collectCompatibleListBlocksForDashboard,
  resolveLinkedListBlockIdForDashboard,
} from '../../core/listPage/linkedViewModeOData';
import { DashboardCard } from './DashboardCard';
import { ChartView } from './ChartView';
import type { IDynamicContext } from '../../core/dynamicTokens/types';

interface IDashboardViewProps {
  dashboardBlockId: string;
  config: IDashboardConfig;
  dataSource: IDataSourceConfig;
  refreshKey?: number;
  onEditCards?: (blockId: string) => void;
  onEditSeries?: (blockId: string) => void;
  onSwitchToCharts?: (blockId: string) => void;
  onCardClick?: (card: IDashboardCardConfig, blockId: string) => void;
  selectedCardId?: string | null;
  onSeriesClick?: (series: IChartSeriesConfig, blockId: string) => void;
  selectedSeriesId?: string | null;
  /** Quando há filtros do dashboard aplicados na listagem. */
  dashboardAppliesListFilter?: boolean;
  /** Limpa todos os filtros ativos (dashboard + tabela). */
  onClearFilters?: () => void;
  /** Lista de página: modo ativo por bloco + config para conjugar filtros OData do modo da tabela. */
  listPairing?: {
    rootConfig: IDynamicViewConfig;
    sections: IListPageSection[];
    dashboardBlock: IListPageBlock;
    rootDashboard: IDashboardConfig;
    activeViewModeByBlockId: Record<string, string>;
  };
  /** Permite gravar vínculo bloco dashboard ↔ bloco lista (mesma fonte OData). */
  onLinkedTableChange?: (pairedListBlockId: string | undefined) => void;
}

export const DashboardView: React.FC<IDashboardViewProps> = ({
  dashboardBlockId,
  config,
  dataSource,
  refreshKey = 0,
  onEditCards,
  onEditSeries,
  onSwitchToCharts,
  onCardClick,
  selectedCardId,
  onSeriesClick,
  selectedSeriesId,
  dashboardAppliesListFilter,
  onClearFilters,
  listPairing,
  onLinkedTableChange,
}) => {
  if (config.dashboardType === 'charts') {
    return (
      <ChartView
        config={config}
        dataSource={dataSource}
        refreshKey={refreshKey}
        onEditSeries={onEditSeries !== undefined ? () => onEditSeries(dashboardBlockId) : undefined}
        onSeriesClick={
          onSeriesClick !== undefined ? (s) => onSeriesClick(s, dashboardBlockId) : undefined
        }
        selectedSeriesId={selectedSeriesId}
        showListFilterHint={dashboardAppliesListFilter === true}
        onClearFilters={onClearFilters}
        listPairing={listPairing}
        onLinkedTableChange={onLinkedTableChange}
      />
    );
  }

  const engine = React.useMemo(() => new DashboardEngine(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);

  const [results, setResults] = useState<IDashboardCardResult[]>(() =>
    engine.buildLoadingResults(config)
  );
  const [globalError, setGlobalError] = useState<string | undefined>(undefined);
  const [fieldMetadata, setFieldMetadata] = useState<Awaited<ReturnType<FieldsService['getVisibleFields']>> | undefined>(undefined);
  const [dynamicContext, setDynamicContext] = useState<IDynamicContext | undefined>(undefined);

  const linkedResolved = listPairing
    ? resolveLinkedListBlockIdForDashboard(listPairing.dashboardBlock, listPairing.rootDashboard)
    : undefined;

  const linkableTables = useMemo(() => {
    if (!listPairing) return [];
    return collectCompatibleListBlocksForDashboard(
      listPairing.rootConfig,
      listPairing.sections,
      listPairing.dashboardBlock
    );
  }, [listPairing]);

  const linkDropdownOptions: IDropdownOption[] = useMemo(() => {
    const opts: IDropdownOption[] = [{ key: '__none__', text: 'Nenhuma (só filtros do dashboard)' }];
    for (let i = 0; i < linkableTables.length; i++) {
      const t = linkableTables[i];
      opts.push({ key: t.id, text: t.label });
    }
    return opts;
  }, [linkableTables]);

  const mergedViewModeOData = useMemo(() => {
    if (!listPairing || !linkedResolved || fieldMetadata === undefined) return undefined;
    return buildLinkedViewModeOData(
      listPairing.rootConfig,
      listPairing.sections,
      dataSource,
      linkedResolved,
      listPairing.activeViewModeByBlockId,
      dynamicContext,
      fieldMetadata
    );
  }, [
    listPairing,
    listPairing?.activeViewModeByBlockId,
    linkedResolved,
    dataSource,
    fieldMetadata,
    dynamicContext,
  ]);

  useEffect(() => {
    const us = new UsersService();
    us.getCurrentUser()
      .then((user) =>
        setDynamicContext(
          buildDynamicContext({
            currentUser: {
              id: user.Id,
              title: user.Title,
              name: user.Title,
              email: user.Email,
              loginName: user.LoginName,
            },
            query:
              typeof window !== 'undefined' && window.location
                ? parseQueryString(window.location.search)
                : undefined,
            now: new Date(),
          })
        )
      )
      .catch(() => setDynamicContext(buildDynamicContext({ now: new Date() })));
  }, []);

  useEffect(() => {
    if (!dataSource.title.trim()) return;
    setFieldMetadata(undefined);
    const lw = dataSource.webServerRelativeUrl?.trim() || undefined;
    fieldsService
      .getVisibleFields(dataSource.title, lw)
      .then(setFieldMetadata)
      .catch(() => setFieldMetadata([]));
  }, [dataSource.title, dataSource.webServerRelativeUrl]);

  useEffect(() => {
    setResults(engine.buildLoadingResults(config));
    setGlobalError(undefined);

    const run = fieldMetadata
      ? engine.computeAll(config, dataSource, fieldMetadata, dynamicContext, mergedViewModeOData)
      : fieldMetadata === undefined
        ? Promise.resolve(engine.buildLoadingResults(config))
        : engine.computeAll(config, dataSource, [], dynamicContext, mergedViewModeOData);

    run
      .then((computed) => setResults(computed))
      .catch((err: Error) => setGlobalError(`Erro ao carregar dashboard: ${err.message}`));
  }, [config, dataSource, fieldMetadata, refreshKey, dynamicContext, mergedViewModeOData, engine]);

  const cardsWithDefaults = useMemo(
    () => (config.cards.length > 0 ? config.cards : generateDefaultCards(config.cardsCount)),
    [config.cards, config.cardsCount]
  );
  const getCardConfig = (id: string) => {
    for (let i = 0; i < cardsWithDefaults.length; i++) {
      if (cardsWithDefaults[i].id === id) return cardsWithDefaults[i];
    }
    return undefined;
  };

  if (results.length === 0) return null;

  return (
    <div style={{ marginBottom: 24 }}>
      <Stack
        horizontal
        horizontalAlign="end"
        verticalAlign="center"
        styles={{ root: { marginBottom: 12 } }}
      >
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 4 }} wrap>
          {linkableTables.length > 0 && onLinkedTableChange !== undefined && (
            <Dropdown
              label="Combinar com modo da tabela"
              options={linkDropdownOptions}
              selectedKey={linkedResolved ?? '__none__'}
              onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
                const k = opt ? String(opt.key) : '__none__';
                onLinkedTableChange(k === '__none__' ? undefined : k);
              }}
              styles={{ root: { minWidth: 260, maxWidth: 320 } }}
            />
          )}
          {onSwitchToCharts !== undefined && (
            <ActionButton
              iconProps={{ iconName: 'BarChartVertical' }}
              onClick={() => onSwitchToCharts(dashboardBlockId)}
              styles={{ root: { height: 28, color: '#0078d4' } }}
            >
              Gráficos
            </ActionButton>
          )}
          {onEditCards !== undefined && (
            <ActionButton
              iconProps={{ iconName: 'Edit' }}
              onClick={() => onEditCards(dashboardBlockId)}
              styles={{ root: { height: 28, color: '#0078d4' } }}
            >
              Editar cards
            </ActionButton>
          )}
        </Stack>
      </Stack>

      {globalError !== undefined && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={false}
          styles={{ root: { marginBottom: 12 } }}
        >
          {globalError}
        </MessageBar>
      )}

      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 16 }}>
        {results.map((result) => {
          const cfg = getCardConfig(result.id);
          return (
            <DashboardCard
              key={result.id}
              result={result}
              cardConfig={cfg}
              selected={selectedCardId === result.id}
              onActivate={onCardClick && cfg ? () => onCardClick(cfg, dashboardBlockId) : undefined}
            />
          );
        })}
      </div>
      {dashboardAppliesListFilter === true && onClearFilters && (
        <ActionButton
          iconProps={{ iconName: 'ClearFilter' }}
          onClick={onClearFilters}
          styles={{ root: { color: '#a4262c', height: 28, marginTop: 4 } }}
        >
          Remover Filtros
        </ActionButton>
      )}
    </div>
  );
};
