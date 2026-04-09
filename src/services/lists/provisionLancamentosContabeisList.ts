import { getSP } from '../core/sp';
import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import {
  DateTimeFieldFormatType,
  DateTimeFieldFriendlyFormatType,
  FieldUserSelectionMode,
  UrlFieldFormatType,
} from '@pnp/sp/fields';

export const LANCAMENTOS_CONTABEIS_LIST_TITLE = 'LancamentosContabeis';
export const NATUREZAS_OPERACAO_LIST_TITLE = 'NaturezasOperacao';

export const LANCAMENTOS_CONTABEIS_PROVISIONED_FIELD_INTERNAL_NAMES = [
  'solicitante',
  'anoFiscal',
  'mesReferencia',
  'razaoContabil',
  'tipoDocumentoSap',
  'valorLancamento',
  'explicacao',
  'aprovadorContabil',
  'validador',
  'dataAprovacao',
  'numeroDocumentoSap',
  'anexoValidadorUrl',
  'naturezaOperacao',
  'statusAprovacao',
  'statusLancamento',
  'statusFluxo',
] as const;

const LIST_DESC =
  'Lançamentos contábeis (SAP, aprovação). Anexos de suporte: anexos do item. Anexo validador: URL no campo indicado (.xlsx).';

const NATUREZAS_DESC = 'Itens para o lookup «Natureza operação» em LancamentosContabeis.';

const MESES_PT = [
  'Janeiro',
  'Fevereiro',
  'Março',
  'Abril',
  'Maio',
  'Junho',
  'Julho',
  'Agosto',
  'Setembro',
  'Outubro',
  'Novembro',
  'Dezembro',
];

function assertCleanInternalName(name: string): void {
  if (!/^[a-z][a-zA-Z0-9]*$/.test(name)) {
    throw new Error(`Nome interno inválido (apenas ASCII, começar por minúscula): ${name}`);
  }
}

function isAlreadyExistsError(e: unknown): boolean {
  const s = String(e);
  return (
    /already\s*exist/i.test(s) ||
    /duplicate/i.test(s) ||
    /0x800706d3/i.test(s) ||
    /field.*exist/i.test(s)
  );
}

async function fieldExists(
  list: ReturnType<SPFI['web']['lists']['getByTitle']>,
  internalName: string
): Promise<boolean> {
  try {
    await list.fields.getByInternalNameOrTitle(internalName)();
    return true;
  } catch {
    return false;
  }
}

async function setFieldDisplayAndDescription(
  list: ReturnType<SPFI['web']['lists']['getByTitle']>,
  internalName: string,
  displayTitle: string,
  description: string
): Promise<void> {
  await list.fields.getByInternalNameOrTitle(internalName).update({
    Title: displayTitle,
    Description: description,
  });
}

export interface IProvisionLancamentosContabeisResult {
  success: boolean;
  messages: string[];
  error?: string;
}

export async function provisionLancamentosContabeisList(): Promise<IProvisionLancamentosContabeisResult> {
  const messages: string[] = [];
  const sp = getSP();

  try {
    await sp.web.lists.ensure(NATUREZAS_OPERACAO_LIST_TITLE, NATUREZAS_DESC, 100, false);
    messages.push(`Lista «${NATUREZAS_OPERACAO_LIST_TITLE}» garantida.`);

    await sp.web.lists.ensure(LANCAMENTOS_CONTABEIS_LIST_TITLE, LIST_DESC, 100, false, {
      EnableAttachments: true,
    });
    messages.push(`Lista «${LANCAMENTOS_CONTABEIS_LIST_TITLE}» garantida.`);

    const natList = sp.web.lists.getByTitle(NATUREZAS_OPERACAO_LIST_TITLE);
    const natMeta = (await natList.select('Id')()) as { Id?: string };
    const lookupListId = String(natMeta.Id ?? '').replace(/[{}]/g, '');

    const mainList = sp.web.lists.getByTitle(LANCAMENTOS_CONTABEIS_LIST_TITLE);
    const mainMeta = (await mainList.select('EnableAttachments')()) as { EnableAttachments?: boolean };
    if (mainMeta.EnableAttachments !== true) {
      await mainList.update({ EnableAttachments: true });
      messages.push('Anexos ativados na lista principal.');
    }

    const ensureUser = async (
      internal: string,
      display: string,
      desc: string,
      mode: FieldUserSelectionMode
    ): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addUser(internal, {
          SelectionMode: mode,
        });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    const ensureText = async (internal: string, display: string, desc: string): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addText(internal, { MaxLength: 255 });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    const ensureNumber = async (internal: string, display: string, desc: string): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addNumber(internal, {
          MinimumValue: 2000,
          MaximumValue: 2100,
        });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    const ensureCurrency = async (internal: string, display: string, desc: string): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addCurrency(internal, {
          CurrencyLocaleId: 1046,
        });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    const ensureNote = async (internal: string, display: string, desc: string): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addMultilineText(internal, {
          RichText: false,
          NumberOfLines: 6,
          AppendOnly: false,
        });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    const ensureUrl = async (internal: string, display: string, desc: string): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addUrl(internal, {
          DisplayFormat: UrlFieldFormatType.Hyperlink,
        });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    const ensureDateTime = async (internal: string, display: string, desc: string): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addDateTime(internal, {
          DisplayFormat: DateTimeFieldFormatType.DateTime,
          FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Unspecified,
        });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    const ensureChoice = async (
      internal: string,
      display: string,
      desc: string,
      choices: string[],
      fillIn: boolean,
      defaultVal?: string
    ): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addChoice(internal, {
          Choices: choices,
          FillInChoice: fillIn,
        });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
        if (defaultVal !== undefined) {
          try {
            await mainList.fields.getByInternalNameOrTitle(internal).update({ DefaultValue: defaultVal });
          } catch {
            /* ignore */
          }
        }
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    const ensureLookup = async (internal: string, display: string, desc: string): Promise<void> => {
      assertCleanInternalName(internal);
      if (await fieldExists(mainList, internal)) return;
      try {
        await mainList.fields.addLookup(internal, {
          LookupListId: lookupListId,
          LookupFieldName: 'Title',
        });
        await setFieldDisplayAndDescription(mainList, internal, display, desc);
      } catch (e) {
        if (!isAlreadyExistsError(e)) throw e;
      }
    };

    await ensureUser(
      'solicitante',
      'Solicitante',
      'Requisitante; predefinido com o utilizador atual no fluxo.',
      FieldUserSelectionMode.PeopleOnly
    );

    await ensureNumber('anoFiscal', 'Ano fiscal', 'Ano fiscal (ex.: 2026).');

    await ensureChoice(
      'mesReferencia',
      'Mês referência',
      'Mês de referência.',
      [...MESES_PT],
      true,
      undefined
    );

    await ensureText('razaoContabil', 'Razão contábil', 'Razão contábil.');
    await ensureText('tipoDocumentoSap', 'Tipo documento SAP', 'Tipo de documento SAP.');
    await ensureCurrency('valorLancamento', 'Valor lançamento', 'Valor do lançamento.');
    await ensureNote(
      'explicacao',
      'Explicação',
      'Comentário do requisitante com detalhes do lançamento.'
    );

    await ensureUser(
      'aprovadorContabil',
      'Aprovador contábil',
      'Membro do grupo Contabilidade (selecionar no site).',
      FieldUserSelectionMode.PeopleAndGroups
    );

    await ensureUser(
      'validador',
      'Validador',
      'Membro do grupo Validador (selecionar no site).',
      FieldUserSelectionMode.PeopleAndGroups
    );

    await ensureDateTime(
      'dataAprovacao',
      'Data de aprovação',
      'Preenchida quando o lançamento for aprovado.'
    );

    await ensureText('numeroDocumentoSap', 'Número documento SAP', 'Número do documento gerado no SAP (texto livre).');
    await ensureUrl(
      'anexoValidadorUrl',
      'Anexo validador (URL)',
      'URL para o ficheiro Excel do validador (.xlsx).'
    );

    await ensureLookup(
      'naturezaOperacao',
      'Natureza operação',
      `Natureza da operação (lista ${NATUREZAS_OPERACAO_LIST_TITLE}).`
    );

    await ensureChoice(
      'statusAprovacao',
      'Status aprovação',
      'Automático no fluxo.',
      ['Pendente', 'Aprovado', 'Reprovado'],
      false,
      'Pendente'
    );

    await ensureChoice(
      'statusLancamento',
      'Status lançamento',
      'Automático no fluxo.',
      ['Rascunho', 'Em validação', 'Validado', 'Invalidado', 'Lançado'],
      false,
      'Rascunho'
    );

    await ensureChoice(
      'statusFluxo',
      'Status ID',
      'Automático no fluxo.',
      ['Em aberto', 'Concluído', 'Cancelado'],
      false,
      'Em aberto'
    );

    messages.push(
      'Campos garantidos: nome interno = primeiro parâmetro add* (ASCII); título visível aplicado em seguida.'
    );
    return { success: true, messages };
  } catch (e) {
    const error = e instanceof Error ? e.message : String(e);
    return { success: false, messages, error };
  }
}

function isFieldMissingOrGoneError(e: unknown): boolean {
  const s = String(e);
  return (
    /404/i.test(s) ||
    /not\s*found/i.test(s) ||
    /does\s*not\s*exist/i.test(s) ||
    /cannot\s*get\s*field/i.test(s) ||
    /field\s*not\s*found/i.test(s) ||
    /0x80070057/i.test(s)
  );
}

export async function deleteLancamentosContabeisProvisionedFields(): Promise<IProvisionLancamentosContabeisResult> {
  const messages: string[] = [];
  const sp = getSP();

  try {
    const list = sp.web.lists.getByTitle(LANCAMENTOS_CONTABEIS_LIST_TITLE);

    for (const internalName of LANCAMENTOS_CONTABEIS_PROVISIONED_FIELD_INTERNAL_NAMES) {
      try {
        await list.fields.getByInternalNameOrTitle(internalName).delete();
        messages.push(`Removido: ${internalName}`);
      } catch (e) {
        if (isFieldMissingOrGoneError(e)) {
          messages.push(`Já não existia: ${internalName}`);
        } else {
          const msg = e instanceof Error ? e.message : String(e);
          return {
            success: false,
            messages,
            error: `Falha ao remover «${internalName}»: ${msg}`,
          };
        }
      }
    }

    messages.push('Concluído (Title, anexos e colunas de sistema não são removidos).');
    return { success: true, messages };
  } catch (e) {
    const error = e instanceof Error ? e.message : String(e);
    if (/does not exist|not found|cannot find/i.test(error)) {
      return {
        success: false,
        messages,
        error: `Lista «${LANCAMENTOS_CONTABEIS_LIST_TITLE}» não encontrada neste site.`,
      };
    }
    return { success: false, messages, error };
  }
}
