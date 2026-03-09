import * as React from 'react';
import { useRef, useEffect } from 'react';
import * as echarts from 'echarts';
import type { EChartsOption } from 'echarts';
import { TChartType } from '../../core/config/types';

const CHART_LABELS: Record<TChartType, string> = {
  bar: 'Barras',
  line: 'Linha',
  area: 'Área',
  pie: 'Pizza',
  donut: 'Rosca',
};

function getOption(type: TChartType): EChartsOption {
  const primary = '#0078d4';
  const secondary = '#2b88d8';
  const tertiary = '#71afe5';

  switch (type) {
    case 'bar':
      return {
        animation: false,
        grid: { left: 4, right: 4, top: 6, bottom: 4 },
        xAxis: { type: 'category', data: ['A', 'B', 'C', 'D'], show: false },
        yAxis: { type: 'value', show: false },
        series: [{ type: 'bar', data: [4, 7, 3, 6], itemStyle: { color: primary }, barMaxWidth: 14 }],
      };
    case 'line':
      return {
        animation: false,
        grid: { left: 4, right: 4, top: 6, bottom: 4 },
        xAxis: { type: 'category', data: ['A', 'B', 'C', 'D', 'E'], show: false },
        yAxis: { type: 'value', show: false },
        series: [{ type: 'line', data: [3, 7, 2, 8, 5], smooth: true, itemStyle: { color: primary }, lineStyle: { color: primary } }],
      };
    case 'area':
      return {
        animation: false,
        grid: { left: 4, right: 4, top: 6, bottom: 4 },
        xAxis: { type: 'category', data: ['A', 'B', 'C', 'D', 'E'], show: false },
        yAxis: { type: 'value', show: false },
        series: [{
          type: 'line',
          data: [3, 7, 2, 8, 5],
          smooth: true,
          areaStyle: { color: `${primary}33` },
          itemStyle: { color: primary },
          lineStyle: { color: primary },
        }],
      };
    case 'pie':
      return {
        animation: false,
        series: [{
          type: 'pie',
          radius: '75%',
          center: ['50%', '50%'],
          data: [
            { value: 35, name: 'A', itemStyle: { color: primary } },
            { value: 40, name: 'B', itemStyle: { color: secondary } },
            { value: 25, name: 'C', itemStyle: { color: tertiary } },
          ],
          label: { show: false },
          emphasis: { disabled: true },
        }],
      };
    case 'donut':
      return {
        animation: false,
        series: [{
          type: 'pie',
          radius: ['38%', '72%'],
          center: ['50%', '50%'],
          data: [
            { value: 35, name: 'A', itemStyle: { color: primary } },
            { value: 40, name: 'B', itemStyle: { color: secondary } },
            { value: 25, name: 'C', itemStyle: { color: tertiary } },
          ],
          label: { show: false },
          emphasis: { disabled: true },
        }],
      };
  }
}

interface IChartTypeCardProps {
  type: TChartType;
  selected: boolean;
  onClick: (type: TChartType) => void;
}

export const ChartTypeCard: React.FC<IChartTypeCardProps> = ({ type, selected, onClick }) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const chartRef = useRef<echarts.ECharts | null>(null);

  useEffect(() => {
    if (!containerRef.current) return;
    chartRef.current = echarts.init(containerRef.current, null, { renderer: 'canvas' });
    chartRef.current.setOption(getOption(type));
    return () => {
      chartRef.current?.dispose();
      chartRef.current = null;
    };
  }, [type]);

  return (
    <div
      onClick={() => onClick(type)}
      style={{
        width: 100,
        border: selected ? '2px solid #0078d4' : '2px solid #edebe9',
        borderRadius: 8,
        padding: '8px 4px 6px',
        cursor: 'pointer',
        background: selected ? '#eff6fc' : '#fff',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        userSelect: 'none',
        transition: 'border-color 0.15s, background 0.15s',
      }}
    >
      <div ref={containerRef} style={{ width: 88, height: 66 }} />
      <span
        style={{
          fontSize: 12,
          color: selected ? '#0078d4' : '#605e5c',
          fontWeight: selected ? 600 : 400,
          marginTop: 6,
        }}
      >
        {CHART_LABELS[type]}
      </span>
    </div>
  );
};
