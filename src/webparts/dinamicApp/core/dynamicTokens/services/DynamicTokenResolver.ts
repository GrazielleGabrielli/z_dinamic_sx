import type { IDynamicContext } from '../types';
import { isDynamicToken, extractDynamicToken, normalizeToken } from '../utils/tokenUtils';

/*
 * Exemplos de uso (com context preenchido):
 *   resolveValue('[me]', context)           → 23
 *   resolveValue('[myEmail]', context)      → 'user@empresa.com'
 *   resolveValue('[today]', context)        → '2025-03-09'
 *   resolveValue('[query:status]', context) → 'Pendente' (se ?status=Pendente na URL)
 *   resolveValue('[null]', context)         → null
 *   resolveValue('[false]', context)        → false
 * Token não resolvido (ex.: [me] sem currentUser) → undefined.
 */
import { resolveQueryToken } from '../utils/queryTokenUtils';
import {
  getToday,
  getNow,
  getTomorrow,
  getYesterday,
  getStartOfMonth,
  getEndOfMonth,
  getStartOfYear,
  getEndOfYear,
  toIsoDateString,
  toIsoDateTimeString,
} from '../utils/dateTokenUtils';
import { QUERY_TOKEN_PREFIX } from '../constants';

/**
 * Resolução centralizada de tokens dinâmicos.
 * Fallback: token não resolvido ou contexto ausente → undefined.
 */
export class DynamicTokenResolver {
  /**
   * Resolve um valor arbitrário: se for string [token], resolve; senão devolve o próprio valor.
   * Ex.: resolveValue('[me]', context) → 23
   *      resolveValue('[myEmail]', context) → 'user@empresa.com'
   *      resolveValue('[query:status]', context) → 'Pendente'
   *      resolveValue('[null]', context) → null
   */
  resolveValue(value: unknown, context: IDynamicContext): unknown {
    if (value === null || value === undefined) return value;
    if (typeof value === 'string' && isDynamicToken(value)) {
      return this.resolveStringToken(value, context);
    }
    return value;
  }

  /**
   * Resolve uma string que deve ser um token [token] ou [query:key].
   * Retorna undefined se o token não puder ser resolvido (ex.: [me] sem currentUser).
   */
  resolveStringToken(token: string, context: IDynamicContext): unknown {
    const raw = extractDynamicToken(token);
    if (raw === null) return undefined;
    const t = normalizeToken(raw);

    if (t.indexOf(QUERY_TOKEN_PREFIX) === 0) {
      return resolveQueryToken(token, context.query);
    }

    switch (t) {
      case 'me':
      case 'myid':
        return context.currentUser?.id !== undefined ? context.currentUser.id : undefined;
      case 'myname':
        return context.currentUser?.title ?? context.currentUser?.name ?? undefined;
      case 'myemail':
        return context.currentUser?.email ?? undefined;
      case 'mylogin':
        return context.currentUser?.loginName ?? undefined;
      case 'mydepartment':
        return context.currentUser?.department ?? undefined;
      case 'myjobtitle':
        return context.currentUser?.jobTitle ?? undefined;
      case 'sitetitle':
        return context.site?.title ?? undefined;
      case 'siteurl':
        return context.site?.url ?? undefined;
      case 'listtitle':
        return context.list?.title ?? undefined;
      case 'empty':
        return '';
      case 'null':
        return null;
      case 'true':
        return true;
      case 'false':
        return false;
      default:
        return this.resolveDateToken(t, context);
    }
  }

  private resolveDateToken(t: string, context: IDynamicContext): string | undefined {
    const now = context.now ?? new Date();
    let d: Date;
    switch (t) {
      case 'today':
        d = getToday(now);
        return toIsoDateString(d);
      case 'now':
        return toIsoDateTimeString(getNow(now));
      case 'tomorrow':
        d = getTomorrow(now);
        return toIsoDateString(d);
      case 'yesterday':
        d = getYesterday(now);
        return toIsoDateString(d);
      case 'startofmonth':
        d = getStartOfMonth(now);
        return toIsoDateString(d);
      case 'endofmonth':
        d = getEndOfMonth(now);
        return toIsoDateString(d);
      case 'startofyear':
        d = getStartOfYear(now);
        return toIsoDateString(d);
      case 'endofyear':
        d = getEndOfYear(now);
        return toIsoDateString(d);
      default:
        return undefined;
    }
  }

  /**
   * Resolve um valor tipado para uso em filtro (número, string, boolean, null).
   * Se for token e resolver para undefined, retorna undefined (chamador pode omitir o segmento).
   */
  resolveFilterValue(value: unknown, context: IDynamicContext): string | number | boolean | null | undefined {
    const resolved = this.resolveValue(value, context);
    if (resolved === undefined) return undefined;
    if (resolved === null) return null;
    if (typeof resolved === 'string' || typeof resolved === 'number' || typeof resolved === 'boolean') {
      return resolved;
    }
    if (resolved instanceof Date) return toIsoDateTimeString(resolved);
    return String(resolved);
  }

  /**
   * Percorre input (objeto ou array) e resolve todos os valores que forem tokens.
   * Objetos e arrays aninhados são processados recursivamente.
   * Ex.: { field: 'ResponsavelId', operator: 'eq', value: '[me]' } → { ..., value: 23 }
   */
  resolveObjectTokens<T>(input: T, context: IDynamicContext): T {
    if (input === null || input === undefined) return input;
    if (typeof input === 'string') {
      const resolved = this.resolveValue(input, context);
      return (resolved !== undefined ? resolved : input) as T;
    }
    if (typeof input === 'number' || typeof input === 'boolean') return input;
    if (Array.isArray(input)) {
      const arr = input.map((item) => this.resolveObjectTokens(item, context));
      return arr as unknown as T;
    }
    if (typeof input === 'object') {
      const out: Record<string, unknown> = {};
      const obj = input as Record<string, unknown>;
      for (const key of Object.keys(obj)) {
        out[key] = this.resolveObjectTokens(obj[key], context);
      }
      return out as unknown as T;
    }
    return input;
  }
}

const defaultResolver = new DynamicTokenResolver();

export function resolveValue(value: unknown, context: IDynamicContext): unknown {
  return defaultResolver.resolveValue(value, context);
}

export function resolveStringToken(token: string, context: IDynamicContext): unknown {
  return defaultResolver.resolveStringToken(token, context);
}

export function resolveFilterValue(value: unknown, context: IDynamicContext): string | number | boolean | null | undefined {
  return defaultResolver.resolveFilterValue(value, context);
}

export function resolveObjectTokens<T>(input: T, context: IDynamicContext): T {
  return defaultResolver.resolveObjectTokens(input, context);
}
