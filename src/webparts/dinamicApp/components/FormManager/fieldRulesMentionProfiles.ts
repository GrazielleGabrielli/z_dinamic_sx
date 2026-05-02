import type { IDropdownOption } from '@fluentui/react';
import type { FieldMappedType, IFieldMetadata } from '../../../../services/shared/types';

export const DEFAULT_VALUE_MENTION_DATE_TOKENS: { literal: string; hint: string }[] = [
  { literal: '[today]', hint: 'Data de hoje (ISO)' },
  { literal: '[now]', hint: 'Data e hora atuais (ISO)' },
  { literal: '[tomorrow]', hint: 'Dia seguinte (ISO)' },
  { literal: '[yesterday]', hint: 'Dia anterior (ISO)' },
  { literal: '[startOfMonth]', hint: 'Primeiro dia do mês corrente' },
  { literal: '[endOfMonth]', hint: 'Último dia do mês corrente' },
  { literal: '[startOfYear]', hint: 'Primeiro dia do ano corrente' },
  { literal: '[endOfYear]', hint: 'Último dia do ano corrente' },
];

export const DEFAULT_VALUE_USER_CONTEXT_TOKENS: { literal: string; hint: string }[] = [
  { literal: '[me]', hint: 'Id numérico do utilizador atual' },
  { literal: '[myId]', hint: 'Igual a [me]' },
  { literal: '[myName]', hint: 'Nome do utilizador' },
  { literal: '[myEmail]', hint: 'E-mail do utilizador' },
  { literal: '[myLogin]', hint: 'Nome de início de sessão' },
];

export const DEFAULT_VALUE_LITERAL_TOKENS: { literal: string; hint: string }[] = [
  { literal: '[empty]', hint: 'Texto vazio' },
  { literal: '[null]', hint: 'Valor nulo' },
  { literal: '[true]', hint: 'Booleano verdadeiro' },
  { literal: '[false]', hint: 'Booleano falso' },
];

export const DEFAULT_VALUE_QUERY_TOKENS: { literal: string; hint: string }[] = [
  { literal: '[query:nome]', hint: 'Valor do parâmetro ?nome= na URL da página' },
];

export const DATE_DEFAULT_MENTION_SUFFIX_PRESETS: { insert: string; primary: string; secondary: string }[] = [
  { insert: ' + 1', primary: '+ 1 dia', secondary: 'Após token ou {{campo}} de data' },
  { insert: ' + 7', primary: '+ 7 dias', secondary: '' },
  { insert: ' + 14', primary: '+ 14 dias', secondary: '' },
  { insert: ' + 30', primary: '+ 30 dias', secondary: '' },
];

/** Expressão em coluna lookup/user: só identificadores numéricos de contexto. */
export const SET_COMPUTED_LOOKUP_EXPR_CONTEXT_TOKENS: { literal: string; hint: string }[] = [
  { literal: '[me]', hint: 'Id numérico do utilizador atual' },
  { literal: '[myId]', hint: 'Igual a [me]' },
];

export const SET_COMPUTED_TOKEN_GROUPS = {
  userProfile: [
    { literal: '[me]', hint: 'Id numérico do utilizador atual' },
    { literal: '[myId]', hint: 'Igual a [me]' },
    { literal: '[myName]', hint: 'Nome do utilizador' },
    { literal: '[myEmail]', hint: 'E-mail do utilizador' },
    { literal: '[myLogin]', hint: 'Nome de início de sessão' },
    { literal: '[myDepartment]', hint: 'Departamento (se disponível)' },
    { literal: '[myJobTitle]', hint: 'Cargo (se disponível)' },
  ],
  siteList: [
    { literal: '[siteTitle]', hint: 'Título do site' },
    { literal: '[siteUrl]', hint: 'URL do site' },
    { literal: '[listTitle]', hint: 'Título da lista' },
  ],
  date: [
    { literal: '[today]', hint: 'Data de hoje (ISO)' },
    { literal: '[now]', hint: 'Data e hora atuais (ISO)' },
    { literal: '[tomorrow]', hint: 'Dia seguinte (ISO)' },
    { literal: '[yesterday]', hint: 'Dia anterior (ISO)' },
    { literal: '[startOfMonth]', hint: 'Primeiro dia do mês corrente' },
    { literal: '[endOfMonth]', hint: 'Último dia do mês corrente' },
    { literal: '[startOfYear]', hint: 'Primeiro dia do ano corrente' },
    { literal: '[endOfYear]', hint: 'Último dia do ano corrente' },
  ],
  literals: [
    { literal: '[empty]', hint: 'Texto vazio' },
    { literal: '[null]', hint: 'Valor nulo' },
    { literal: '[true]', hint: 'Booleano verdadeiro' },
    { literal: '[false]', hint: 'Booleano falso' },
  ],
  query: [{ literal: '[query:nome]', hint: 'Valor do parâmetro ?nome= na URL da página' }],
} as const;

export type TSetComputedTokenPick = {
  userProfile: boolean;
  siteList: boolean;
  date: boolean;
  literals: boolean;
  query: boolean;
  /** Só [me]/[myId]; ignora os outros flags de token. */
  lookupFieldExprOnly?: boolean;
};

export function mergeSetComputedContextTokens(pick: TSetComputedTokenPick): { literal: string; hint: string }[] {
  if (pick.lookupFieldExprOnly) {
    return SET_COMPUTED_LOOKUP_EXPR_CONTEXT_TOKENS.slice();
  }
  const out: { literal: string; hint: string }[] = [];
  if (pick.userProfile) out.push(...SET_COMPUTED_TOKEN_GROUPS.userProfile);
  if (pick.siteList) out.push(...SET_COMPUTED_TOKEN_GROUPS.siteList);
  if (pick.date) out.push(...SET_COMPUTED_TOKEN_GROUPS.date);
  if (pick.literals) out.push(...SET_COMPUTED_TOKEN_GROUPS.literals);
  if (pick.query) out.push(...SET_COMPUTED_TOKEN_GROUPS.query);
  return out;
}

export type TDefaultValueMentionParts = {
  dateTokens: boolean;
  userContextTokens: boolean;
  literalTokens: boolean;
  queryTemplate: boolean;
  dateSuffixesAndDateRefs: boolean;
  lookupPaths: boolean;
  numericRefs: boolean;
};

export function defaultValueMentionPartsNumericOnly(): TDefaultValueMentionParts {
  return {
    dateTokens: false,
    userContextTokens: false,
    literalTokens: false,
    queryTemplate: false,
    dateSuffixesAndDateRefs: false,
    lookupPaths: false,
    numericRefs: true,
  };
}

export function defaultValueMentionParts(mt: FieldMappedType): TDefaultValueMentionParts {
  switch (mt) {
    case 'text':
    case 'multiline':
    case 'url':
    case 'choice':
    case 'multichoice':
    case 'taxonomy':
    case 'taxonomymulti':
      return {
        dateTokens: true,
        userContextTokens: true,
        literalTokens: true,
        queryTemplate: true,
        dateSuffixesAndDateRefs: true,
        lookupPaths: true,
        numericRefs: true,
      };
    case 'datetime':
      return {
        dateTokens: true,
        userContextTokens: false,
        literalTokens: false,
        queryTemplate: false,
        dateSuffixesAndDateRefs: true,
        lookupPaths: false,
        numericRefs: true,
      };
    case 'number':
    case 'currency':
      return {
        dateTokens: false,
        userContextTokens: false,
        literalTokens: true,
        queryTemplate: true,
        dateSuffixesAndDateRefs: false,
        lookupPaths: false,
        numericRefs: true,
      };
    case 'boolean':
      return {
        dateTokens: false,
        userContextTokens: false,
        literalTokens: true,
        queryTemplate: true,
        dateSuffixesAndDateRefs: false,
        lookupPaths: false,
        numericRefs: false,
      };
    case 'lookup':
    case 'lookupmulti':
    case 'user':
    case 'usermulti':
      return {
        dateTokens: false,
        userContextTokens: true,
        literalTokens: true,
        queryTemplate: true,
        dateSuffixesAndDateRefs: false,
        lookupPaths: true,
        numericRefs: true,
      };
    default:
      return {
        dateTokens: true,
        userContextTokens: true,
        literalTokens: true,
        queryTemplate: true,
        dateSuffixesAndDateRefs: true,
        lookupPaths: true,
        numericRefs: true,
      };
  }
}

/** Referências {{Campo}} em expressão “completa”; exclui calculada. */
export const REF_NON_CALCULATED: ReadonlySet<FieldMappedType> = new Set<FieldMappedType>([
  'text',
  'multiline',
  'choice',
  'multichoice',
  'number',
  'currency',
  'boolean',
  'datetime',
  'lookup',
  'lookupmulti',
  'user',
  'usermulti',
  'url',
  'taxonomy',
  'taxonomymulti',
]);

const REF_DATETIME_ONLY: ReadonlySet<FieldMappedType> = new Set<FieldMappedType>(['datetime']);

const REF_NUMERIC: ReadonlySet<FieldMappedType> = new Set<FieldMappedType>(['number', 'currency']);

const REF_LOOKUP_EXPR_SIMPLE_FIELDS: ReadonlySet<FieldMappedType> = new Set<FieldMappedType>([
  'lookup',
  'lookupmulti',
  'user',
  'usermulti',
]);

export type TSetComputedMentionParts = {
  tokens: TSetComputedTokenPick;
  allowedRefMappedTypes: ReadonlySet<FieldMappedType>;
  includeAttfolders: boolean;
  includeLookupPaths: boolean;
  includeNumericAux: boolean;
  expressionHelpVariant: 'full' | 'datetime' | 'numeric' | 'boolean' | 'lookup';
};

export function setComputedMentionParts(mt: FieldMappedType): TSetComputedMentionParts {
  switch (mt) {
    case 'text':
    case 'multiline':
    case 'url':
    case 'choice':
    case 'multichoice':
    case 'taxonomy':
    case 'taxonomymulti':
      return {
        tokens: {
          userProfile: true,
          siteList: true,
          date: true,
          literals: true,
          query: true,
        },
        allowedRefMappedTypes: REF_NON_CALCULATED,
        includeAttfolders: true,
        includeLookupPaths: true,
        includeNumericAux: true,
        expressionHelpVariant: 'full',
      };
    case 'datetime':
      return {
        tokens: {
          userProfile: false,
          siteList: false,
          date: true,
          literals: true,
          query: true,
        },
        allowedRefMappedTypes: REF_DATETIME_ONLY,
        includeAttfolders: false,
        includeLookupPaths: false,
        includeNumericAux: true,
        expressionHelpVariant: 'datetime',
      };
    case 'number':
    case 'currency':
      return {
        tokens: {
          userProfile: false,
          siteList: false,
          date: false,
          literals: true,
          query: true,
        },
        allowedRefMappedTypes: REF_NUMERIC,
        includeAttfolders: false,
        includeLookupPaths: false,
        includeNumericAux: true,
        expressionHelpVariant: 'numeric',
      };
    case 'boolean':
      return {
        tokens: {
          userProfile: false,
          siteList: false,
          date: false,
          literals: true,
          query: true,
        },
        allowedRefMappedTypes: REF_NON_CALCULATED,
        includeAttfolders: false,
        includeLookupPaths: false,
        includeNumericAux: false,
        expressionHelpVariant: 'boolean',
      };
    case 'lookup':
    case 'lookupmulti':
    case 'user':
    case 'usermulti':
      return {
        tokens: {
          userProfile: false,
          siteList: false,
          date: false,
          literals: false,
          query: false,
          lookupFieldExprOnly: true,
        },
        allowedRefMappedTypes: REF_LOOKUP_EXPR_SIMPLE_FIELDS,
        includeAttfolders: false,
        includeLookupPaths: true,
        includeNumericAux: true,
        expressionHelpVariant: 'lookup',
      };
    default:
      return {
        tokens: {
          userProfile: true,
          siteList: true,
          date: true,
          literals: true,
          query: true,
        },
        allowedRefMappedTypes: REF_NON_CALCULATED,
        includeAttfolders: true,
        includeLookupPaths: true,
        includeNumericAux: true,
        expressionHelpVariant: 'full',
      };
  }
}

export function filterFieldOptionsByMappedTypes(
  fieldOptions: IDropdownOption[],
  listFieldMetadata: IFieldMetadata[] | undefined,
  allowed: ReadonlySet<FieldMappedType>,
  excludeInternalName: string
): IDropdownOption[] {
  if (!listFieldMetadata?.length) {
    return fieldOptions.filter((o) => String(o.key) !== excludeInternalName);
  }
  const byInternal = new Map(listFieldMetadata.map((m) => [m.InternalName, m]));
  return fieldOptions.filter((o) => {
    const k = String(o.key);
    if (k === excludeInternalName) return false;
    const meta = byInternal.get(k);
    if (!meta) return true;
    const mt = meta.MappedType ?? 'unknown';
    if (mt === 'unknown') return true;
    return allowed.has(mt as FieldMappedType);
  });
}
