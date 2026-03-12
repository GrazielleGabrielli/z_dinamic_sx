import { extractQueryKey } from './tokenUtils';

/**
 * Resolve [query:key] usando context.query.
 * Retorna undefined se a chave não existir (fallback consistente).
 */
export function resolveQueryToken(
  value: string,
  query: Record<string, string> | undefined
): string | undefined {
  const key = extractQueryKey(value);
  if (key === null || !query) return undefined;
  const v = query[key];
  return v !== undefined ? v : undefined;
}
