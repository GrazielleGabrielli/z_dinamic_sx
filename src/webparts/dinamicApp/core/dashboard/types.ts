import { IDashboardCardStyleConfig } from '../config/types';

export type TCardStatus = 'loading' | 'ready' | 'error';

export interface IDashboardCardResult {
  id: string;
  title: string;
  subtitle?: string;
  aggregate: string;
  value: number | undefined;
  status: TCardStatus;
  error?: string;
  style?: IDashboardCardStyleConfig;
  emptyValueText?: string;
  errorText?: string;
  loadingText?: string;
}
