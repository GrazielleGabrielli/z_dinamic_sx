import { QUERY_TOKEN_PREFIX } from '../constants';

/**
 * Verifica se value é uma string no formato [token] (case-insensitive).
 * Não considera [query:status] como "apenas token" sem chave; use isQueryToken para query.
 */
export function isDynamicToken(value: unknown): value is string {
  if (typeof value !== 'string') return false;
  const s = value.trim();
  return s.length >= 3 && s.charAt(0) === '[' && s.charAt(s.length - 1) === ']';
}

/**
 * Extrai o conteúdo entre colchetes ou null se não for token válido.
 * Ex.: "[me]" → "me", "[query:status]" → "query:status".
 */
export function extractDynamicToken(value: string): string | null {
  if (typeof value !== 'string') return null;
  const s = value.trim();
  if (s.length < 3 || s.charAt(0) !== '[' || s.charAt(s.length - 1) !== ']') return null;
  return s.slice(1, s.length - 1).trim() || null;
}

/**
 * Verifica se o token extraído é do tipo [query:key].
 */
export function isQueryToken(value: string): boolean {
  const token = extractDynamicToken(value);
  return token !== null && token.toLowerCase().indexOf(QUERY_TOKEN_PREFIX) === 0;
}

/**
 * Para [query:status] retorna "status". Para token não-query retorna null.
 */
export function extractQueryKey(value: string): string | null {
  const token = extractDynamicToken(value);
  if (token === null) return null;
  const lower = token.toLowerCase();
  if (lower.indexOf(QUERY_TOKEN_PREFIX) !== 0) return null;
  const key = token.slice(QUERY_TOKEN_PREFIX.length).trim();
  return key || null;
}

/**
 * Normaliza token para comparação (minúsculas).
 */
export function normalizeToken(raw: string): string {
  return raw.trim().toLowerCase();
}
