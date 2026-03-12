/**
 * Contexto disponível na resolução de tokens dinâmicos.
 * Montado a partir de WebPartContext, URL e config (ex.: buildDynamicContext).
 */
export interface IDynamicContext {
  currentUser?: {
    id?: number;
    title?: string;
    name?: string;
    email?: string;
    loginName?: string;
    department?: string;
    jobTitle?: string;
  };
  site?: {
    title?: string;
    url?: string;
  };
  list?: {
    title?: string;
  };
  /** Parâmetros da query string (ex.: ?status=Pendente → { status: 'Pendente' }) */
  query?: Record<string, string>;
  /** Data/hora de referência para tokens de data. Se ausente, usa new Date(). */
  now?: Date;
}

/** Tokens de usuário atual. Resolução: undefined se currentUser ausente. */
export type TUserDynamicToken =
  | 'me'
  | 'myId'
  | 'myName'
  | 'myEmail'
  | 'myLogin'
  | 'myDepartment'
  | 'myJobTitle';

/** Tokens de data. Usam context.now ou new Date(). Retorno: string ISO ou Date conforme uso. */
export type TDateDynamicToken =
  | 'today'
  | 'now'
  | 'tomorrow'
  | 'yesterday'
  | 'startOfMonth'
  | 'endOfMonth'
  | 'startOfYear'
  | 'endOfYear';

/** Token [query:key]. Resolução: context.query?.[key] ?? undefined. */
export type TQueryDynamicToken = 'query';

/** Tokens de site/lista. Resolução: undefined se site/list ausente. */
export type TSiteDynamicToken = 'siteTitle' | 'siteUrl' | 'listTitle';

/** Tokens especiais/booleanos. Sempre resolvem. */
export type TSpecialDynamicToken = 'empty' | 'null' | 'true' | 'false';

export type TDynamicTokenKind =
  | TUserDynamicToken
  | TDateDynamicToken
  | TSiteDynamicToken
  | TSpecialDynamicToken
  | 'query';
