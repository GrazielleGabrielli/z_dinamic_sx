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
const OUTSIDE_BOUNDS_RESTRICT_DAYS = 500;

function addDaysLocal(base: Date, deltaDays: number): Date {
  const x = startOfDay(base);
  x.setDate(x.getDate() + deltaDays);
  return x;
}

function dateKeyLocal(d: Date): string {
  const s = startOfDay(d);
  return `${s.getFullYear()}-${s.getMonth()}-${s.getDate()}`;
}

function mergeUniqueRestrictedDates(a: readonly Date[], b: readonly Date[]): Date[] {
  const seen = new Set<string>();
  const out: Date[] = [];
  const push = (d: Date): void => {
    const sd = startOfDay(d);
    const k = dateKeyLocal(sd);
    if (seen.has(k)) return;
    seen.add(k);
    out.push(sd);
  };
  for (let i = 0; i < a.length; i++) push(a[i]);
  for (let i = 0; i < b.length; i++) push(b[i]);
  return out;
}

function restrictedDatesBeforeMin(minDate: Date, count: number): Date[] {
  const out: Date[] = [];
  let d = addDaysLocal(minDate, -1);
  for (let i = 0; i < count; i++) {
    out.push(d);
    d = addDaysLocal(d, -1);
  }
  return out;
}

function restrictedDatesAfterMax(maxDate: Date, count: number): Date[] {
  const out: Date[] = [];
  let d = addDaysLocal(maxDate, 1);
  for (let i = 0; i < count; i++) {
    out.push(d);
    d = addDaysLocal(d, 1);
  }
  return out;
}

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

  const outsideRange: Date[] = [];
  if (bounds.minDate) {
    outsideRange.push(...restrictedDatesBeforeMin(bounds.minDate, OUTSIDE_BOUNDS_RESTRICT_DAYS));
  }
  if (bounds.maxDate) {
    outsideRange.push(...restrictedDatesAfterMax(bounds.maxDate, OUTSIDE_BOUNDS_RESTRICT_DAYS));
  }

  const out: IValidateDateCalendarProps = {};
  if (bounds.minDate) out.minDate = bounds.minDate;
  if (bounds.maxDate) out.maxDate = bounds.maxDate;
  const mergedRestricted =
    restrictedDates.length > 0 || outsideRange.length > 0
      ? mergeUniqueRestrictedDates(restrictedDates, outsideRange)
      : [];
  if (mergedRestricted.length > 0) out.restrictedDates = mergedRestricted;
  return out;
}
