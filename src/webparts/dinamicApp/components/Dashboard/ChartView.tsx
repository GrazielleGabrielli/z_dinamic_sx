import * as React from 'react';
import { useRef, useEffect, useState, useMemo } from 'react';
import * as echarts from 'echarts';
import type { EChartsOption } from 'echarts';
import { Stack, Text, ActionButton, Spinner, SpinnerSize, MessageBar, MessageBarType } from '@fluentui/react';
import { FieldsService } from '../../../../services';
import { IDashboardConfig, IDataSourceConfig, TChartType } from '../../core/config/types';
import { IChartSeriesResult } from '../../core/dashboard/types';
import { DashboardEngine } from '../../core/dashboard/DashboardEngine';

const DEFAULT_COLORS = ['#0078d4', '#2b88d8', '#71afe5', '#00b294', '#ffaa44', '#d13438', '#8764b8', '#038387'];

function buildOption(chartType: TChartType, series: IChartSeriesResult[]): EChartsOption {
  const labels = series.map((s) => s.label);
  const values = series.map((s) => s.value);
  const colors = series.map((s, i) => s.color ?? DEFAULT_COLORS[i % DEFAULT_COLORS.length]);

  if (chartType === 'pie' || chartType === 'donut') {
    return {
      tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
      legend: { orient: 'horizontal', bottom: 0, data: labels },
      series: [{
        type: 'pie',
        radius: chartType === 'donut' ? ['40%', '70%'] : '65%',
        center: ['50%', '45%'],
        data: series.map((s, i) => ({
          name: s.label,
          value: s.value,
          itemStyle: { color: s.color ?? DEFAULT_COLORS[i % DEFAULT_COLORS.length] },
        })),
        label: { show: true, formatter: '{b}: {c}' },
        emphasis: { itemStyle: { shadowBlur: 10, shadowOffsetX: 0, shadowColor: 'rgba(0,0,0,0.3)' } },
      }],
    };
  }

  const isArea = chartType === 'area';

  return {
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
    grid: { left: 48, right: 24, top: 24, bottom: 36 },
    xAxis: {
      type: 'category',
      data: labels,
      axisLabel: { interval: 0, rotate: labels.length > 6 ? 30 : 0 },
    },
    yAxis: { type: 'value', minInterval: 1 },
    series: [{
      type: chartType === 'bar' ? 'bar' : 'line',
      data: values.map((v, i) => ({ value: v, itemStyle: { color: colors[i] } })),
      smooth: chartType !== 'bar',
      areaStyle: isArea ? { opacity: 0.3 } : undefined,
      lineStyle: chartType !== 'bar' ? { color: colors[0], width: 2 } : undefined,
      itemStyle: chartType === 'bar' ? undefined : { color: colors[0] },
      barMaxWidth: 60,
    }],
  };
}

interface IChartState {
  results: IChartSeriesResult[];
  loading: boolean;
  error: string | undefined;
}

interface IChartViewProps {
  config: IDashboardConfig;
  dataSource: IDataSourceConfig;
  onEditSeries: () => void;
}

export const ChartView: React.FC<IChartViewProps> = ({ config, dataSource, onEditSeries }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const chartRef = useRef<echarts.ECharts | null>(null);
  const engine = useMemo(() => new DashboardEngine(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);

  const [chartState, setChartState] = useState<IChartState>({
    results: [],
    loading: false,
    error: undefined,
  });
  const [fieldMetadata, setFieldMetadata] = useState<Awaited<ReturnType<FieldsService['getVisibleFields']>> | undefined>(undefined);

  useEffect(() => {
    if (!dataSource.title.trim()) return;
    setFieldMetadata(undefined);
    fieldsService.getVisibleFields(dataSource.title).then(setFieldMetadata).catch(() => setFieldMetadata([]));
  }, [dataSource.title]);

  useEffect(() => {
    const series = config.chartSeries ?? [];
    if (series.length === 0) {
      setChartState({ results: [], loading: false, error: undefined });
      return;
    }
    if (fieldMetadata === undefined) {
      setChartState((s) => ({ ...s, loading: true, error: undefined }));
      return;
    }
    setChartState((s) => ({ ...s, loading: true, error: undefined }));
    engine
      .computeAllSeries(config, dataSource, fieldMetadata)
      .then((results) => setChartState({ results, loading: false, error: undefined }))
      .catch((err: Error) => setChartState({ results: [], loading: false, error: `Erro ao carregar dados: ${err.message}` }));
  }, [config, dataSource, fieldMetadata]);

  // container sempre montado no DOM — só visibility muda
  // assim containerRef.current nunca é null quando o effect de init roda
  useEffect(() => {
    if (!containerRef.current || chartState.results.length === 0) return;

    if (!chartRef.current || chartRef.current.isDisposed()) {
      chartRef.current = echarts.init(containerRef.current, null, { renderer: 'canvas' });
    }

    const chartType = config.chartType ?? 'bar';
    chartRef.current.setOption(buildOption(chartType, chartState.results), true);
    chartRef.current.resize();
  }, [chartState.results, config.chartType]);

  useEffect(() => {
    return () => {
      chartRef.current?.dispose();
      chartRef.current = null;
    };
  }, []);

  const hasSeries = (config.chartSeries ?? []).length > 0;
  const showChart = hasSeries && !chartState.loading && chartState.results.length > 0;

  return (
    <div style={{ marginBottom: 24 }}>
      <Stack
        horizontal
        horizontalAlign="space-between"
        verticalAlign="center"
        styles={{ root: { marginBottom: 12 } }}
      >
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
          Dashboard
        </Text>
        <ActionButton
          iconProps={{ iconName: 'Edit' }}
          onClick={onEditSeries}
          styles={{ root: { height: 28, color: '#0078d4' } }}
        >
          Editar séries
        </ActionButton>
      </Stack>

      {chartState.error !== undefined && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false} styles={{ root: { marginBottom: 12 } }}>
          {chartState.error}
        </MessageBar>
      )}

      {!hasSeries && (
        <div
          style={{
            border: '2px dashed #edebe9',
            borderRadius: 8,
            padding: '32px 24px',
            textAlign: 'center',
            background: '#faf9f8',
          }}
        >
          <Text variant="medium" styles={{ root: { color: '#a19f9d', display: 'block', marginBottom: 12 } }}>
            Nenhuma série configurada ainda.
          </Text>
          <ActionButton iconProps={{ iconName: 'Add' }} onClick={onEditSeries} styles={{ root: { color: '#0078d4' } }}>
            Adicionar série
          </ActionButton>
        </div>
      )}

      {hasSeries && chartState.loading && (
        <Stack horizontalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { padding: '32px 0' } }}>
          <Spinner size={SpinnerSize.medium} />
          <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>Carregando dados...</Text>
        </Stack>
      )}

      {/* container sempre no DOM quando hasSeries — display:none evita o problema de containerRef=null */}
      <div
        style={{
          background: '#fff',
          border: '1px solid #edebe9',
          borderRadius: 8,
          padding: '16px',
          display: showChart ? 'block' : 'none',
        }}
      >
        <div ref={containerRef} style={{ width: '100%', height: 320 }} />
      </div>
    </div>
  );
};
