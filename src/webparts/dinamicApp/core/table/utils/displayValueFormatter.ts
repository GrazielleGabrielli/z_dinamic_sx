import type { ITableColumnConfig, ITableColumnFormattingConfig } from '../types';
import { DEFAULT_BOOLEAN_TRUE, DEFAULT_BOOLEAN_FALSE } from '../constants/tableDefaults';

function formatNumber(value: number, formatting?: ITableColumnFormattingConfig): string {
  const decimals = formatting?.numberDecimals;
  if (decimals !== undefined) {
    return value.toFixed(decimals);
  }
  return String(value);
}

function formatCurrency(value: number, formatting?: ITableColumnFormattingConfig): string {
  const code = formatting?.currencyCode ?? 'BRL';
  const decimals = formatting?.numberDecimals ?? 2;
  try {
    return new Intl.NumberFormat('pt-BR', {
      style: 'currency',
      currency: code,
      minimumFractionDigits: decimals,
      maximumFractionDigits: decimals,
    }).format(value);
  } catch {
    return `${formatNumber(value, { ...formatting, numberDecimals: decimals })}`;
  }
}

function formatDate(value: unknown, formatting?: ITableColumnFormattingConfig): string {
  if (value == null) return '';
  const str = String(value);
  const date = new Date(str);
  if (isNaN(date.getTime())) return str;
  const format = formatting?.dateFormat;
  if (format) {
    const d = date.getDate();
    const m = date.getMonth() + 1;
    const y = date.getFullYear();
    const pad = (n: number) => (n < 10 ? '0' + n : String(n));
    return format
      .replace('dd', pad(d))
      .replace('MM', pad(m))
      .replace('yyyy', String(y));
  }
  return date.toLocaleDateString('pt-BR');
}

function formatBoolean(value: unknown, formatting?: ITableColumnFormattingConfig): string {
  const b = Boolean(value);
  return b ? (formatting?.trueText ?? DEFAULT_BOOLEAN_TRUE) : (formatting?.falseText ?? DEFAULT_BOOLEAN_FALSE);
}

export function formatDisplayValue(
  rawValue: unknown,
  column: ITableColumnConfig
): string {
  const formatting = column.formatting;
  const prefix = formatting?.prefix ?? '';
  const suffix = formatting?.suffix ?? '';

  if (rawValue == null || rawValue === '') return '';

  switch (column.fieldType) {
    case 'number':
      return prefix + formatNumber(Number(rawValue), formatting) + suffix;
    case 'currency':
      return prefix + formatCurrency(Number(rawValue), formatting) + suffix;
    case 'date':
      return prefix + formatDate(rawValue, formatting) + suffix;
    case 'boolean':
      return prefix + formatBoolean(rawValue, formatting) + suffix;
    default:
      return prefix + String(rawValue) + suffix;
  }
}
