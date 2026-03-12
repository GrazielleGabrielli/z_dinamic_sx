/**
 * Tokens dinâmicos suportados. Útil para UI, validação e documentação.
 * Formato no JSON: "[token]" ou "[query:chave]".
 */

export const USER_TOKENS = [
  'me',
  'myId',
  'myName',
  'myEmail',
  'myLogin',
  'myDepartment',
  'myJobTitle',
] as const;

export const DATE_TOKENS = [
  'today',
  'now',
  'tomorrow',
  'yesterday',
  'startOfMonth',
  'endOfMonth',
  'startOfYear',
  'endOfYear',
] as const;

export const QUERY_TOKEN_PREFIX = 'query:';

export const SITE_TOKENS = ['siteTitle', 'siteUrl', 'listTitle'] as const;

export const SPECIAL_TOKENS = ['empty', 'null', 'true', 'false'] as const;

export const ALL_STATIC_TOKENS: readonly string[] = [
  ...USER_TOKENS,
  ...DATE_TOKENS,
  ...SITE_TOKENS,
  ...SPECIAL_TOKENS,
];

/** Regex: [token] com token em minúsculas, ou [query:key] */
export const TOKEN_PATTERN = /^\[([a-z]+(?::[a-z0-9_-]+)?)\]$/i;
