import type { IDynamicContext } from '../dynamicTokens/types';
import type { TFormRule, TFormManagerFormMode, TFormSubmitKind } from '../config/types/formManager';
import {
  collectApplicableValidateDateRules,
  evaluateValidateDateRulesForField,
  mergeValidateDateCalendarBounds,
  startOfDay,
} from './formRuleEngine';

export interface IValidateDateCalendarProps {
  minDate?: Date;
  maxDate?: Date;
  restrictedDates?: Date[];
}

const CALENDAR_RESTRICTED_SCAN_CAP = 800;

function enumerateDaysCapped(start: Date, end: Date, maxDays: number): Date[] {
  const out: Date[] = [];
  const first = startOfDay(start);
  const last = startOfDay(end);
  for (let n = 0; n < maxDays; n++) {
    const d = new Date(first.getFullYear(), first.getMonth(), first.getDate() + n);
    if (d > last) break;
    out.push(d);
  }
  return out;
}

function getRestrictedDatesScanRange(
  mergedMin: Date | undefined,
  mergedMax: Date | undefined,
  now: Date
): { start: Date; end: Date } {
  const pad = 366 * 2;
  const t0 = startOfDay(now);
  if (mergedMin && mergedMax) {
    return { start: mergedMin, end: mergedMax };
  }
  if (mergedMin && !mergedMax) {
    const end = new Date(mergedMin.getTime());
    end.setDate(end.getDate() + pad);
    return { start: mergedMin, end: startOfDay(end) };
  }
  if (!mergedMin && mergedMax) {
    const start = new Date(mergedMax.getTime());
    start.setDate(start.getDate() - pad);
    return { start: startOfDay(start), end: mergedMax };
  }
  const start = new Date(t0.getTime());
  start.setDate(start.getDate() - 366);
  const end = new Date(t0.getTime());
  end.setDate(end.getDate() + pad);
  return { start: startOfDay(start), end: startOfDay(end) };
}

export function buildValidateDateCalendarProps(
  rules: readonly TFormRule[],
  field: string,
  values: Record<string, unknown>,
  params: {
    formMode: TFormManagerFormMode;
    submitKind: TFormSubmitKind | undefined;
    userGroupTitles: string[];
    dynamicContext: IDynamicContext;
    fieldVisible: (name: string) => boolean;
    now?: Date;
  }
): IValidateDateCalendarProps {
  const applicable = collectApplicableValidateDateRules(rules, field, values, params);
  if (applicable.length === 0) return {};

  const ts = params.now ?? new Date();
  const bounds = mergeValidateDateCalendarBounds(applicable, values, params.dynamicContext, ts);
  const range = getRestrictedDatesScanRange(bounds.minDate, bounds.maxDate, ts);
  const days = enumerateDaysCapped(range.start, range.end, CALENDAR_RESTRICTED_SCAN_CAP);

  const restrictedDates: Date[] = [];
  for (let i = 0; i < days.length; i++) {
    const day = days[i];
    const iso = day.toISOString();
    const nextValues = { ...values, [field]: iso };
    const msg = evaluateValidateDateRulesForField(rules, field, nextValues, { ...params, now: ts });
    if (msg) restrictedDates.push(startOfDay(day));
  }

  const out: IValidateDateCalendarProps = {};
  if (bounds.minDate) out.minDate = bounds.minDate;
  if (bounds.maxDate) out.maxDate = bounds.maxDate;
  if (restrictedDates.length > 0) out.restrictedDates = restrictedDates;
  return out;
}
