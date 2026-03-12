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
