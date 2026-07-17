/**
 * Helpers de data para tokens [today], [startOfMonth], etc.
 * Usa referência now; formatação consistente para filtros OData (ISO date ou datetime).
 * Futuro: timezone, formatos configuráveis.
 */

function toLocalDate(d: Date): Date {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

export function getToday(now: Date): Date {
  return toLocalDate(now);
}

export function getNow(now: Date): Date {
  return new Date(now.getTime());
}

export function getTomorrow(now: Date): Date {
  const t = toLocalDate(now);
  t.setDate(t.getDate() + 1);
  return t;
}

export function getYesterday(now: Date): Date {
  const t = toLocalDate(now);
  t.setDate(t.getDate() - 1);
  return t;
}

export function getStartOfMonth(now: Date): Date {
  return new Date(now.getFullYear(), now.getMonth(), 1);
}

export function getEndOfMonth(now: Date): Date {
  return new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59, 999);
}

export function getStartOfYear(now: Date): Date {
  return new Date(now.getFullYear(), 0, 1);
}

export function getEndOfYear(now: Date): Date {
  return new Date(now.getFullYear(), 11, 31, 23, 59, 59, 999);
}

/**
 * Retorno padrão para tokens de data em filtros: string ISO (date ou datetime).
 * [today] → "2025-03-09", [now] → "2025-03-09T12:00:00.000Z" (ex.).
 */
function pad2(n: number): string {
  const s = String(n);
  return s.length >= 2 ? s : '0' + s;
}

export function toIsoDateString(d: Date): string {
  const y = d.getFullYear();
  const m = pad2(d.getMonth() + 1);
  const day = pad2(d.getDate());
  return y + '-' + m + '-' + day;
}

export function toIsoDateTimeString(d: Date): string {
  return d.toISOString();
}

/**
 * Dia civil local a partir de valor de formulário/SharePoint.
 * `YYYY-MM-DD` e DateOnly (`…T00:00:00Z`) usam o calendário da string —
 * evita o shift de `new Date('YYYY-MM-DD')` (UTC) em fusos negativos.
 */
export function parseCalendarDateValue(value: unknown): Date | undefined {
  if (value === null || value === undefined || value === '') return undefined;
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return undefined;
    return toLocalDate(value);
  }
  const t = String(value).trim();
  if (!t) return undefined;

  const br = /^(\d{1,2})[-/](\d{1,2})[-/](\d{4})$/.exec(t);
  if (br) {
    const day = parseInt(br[1], 10);
    const month = parseInt(br[2], 10) - 1;
    const year = parseInt(br[3], 10);
    const dt = new Date(year, month, day);
    if (dt.getFullYear() !== year || dt.getMonth() !== month || dt.getDate() !== day) return undefined;
    return dt;
  }

  const isoPrefix = /^(\d{4})-(\d{2})-(\d{2})/.exec(t);
  if (isoPrefix) {
    const pure = /^\d{4}-\d{2}-\d{2}$/.test(t);
    const midnightUtc = /^\d{4}-\d{2}-\d{2}T00:00:00(\.\d+)?(Z|[+-]00:00)?$/i.test(t);
    if (pure || midnightUtc) {
      const y = parseInt(isoPrefix[1], 10);
      const mo = parseInt(isoPrefix[2], 10) - 1;
      const d = parseInt(isoPrefix[3], 10);
      const dt = new Date(y, mo, d);
      if (dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== d) return undefined;
      return dt;
    }
    const instant = new Date(t);
    if (isNaN(instant.getTime())) return undefined;
    return toLocalDate(instant);
  }

  const instant = new Date(t);
  if (isNaN(instant.getTime())) return undefined;
  return toLocalDate(instant);
}
