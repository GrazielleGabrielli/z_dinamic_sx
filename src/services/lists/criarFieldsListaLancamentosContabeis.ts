import type { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import { DateTimeFieldFormatType } from '@pnp/sp/fields';

export const LISTA_LANCAMENTOS_CONTABEIS = 'LancamentosContabeis';

/** Lista de destino do lookup «Natureza Operação». Criada automaticamente se não existir. */
export const LISTA_NATUREZA_OPERACAO_LOOKUP = 'NaturezasOperacao';

const DESC_LISTA_NATUREZA =
  'Catálogo para o campo Lookup «NaturezaOperacao» em LancamentosContabeis. Adicione um item por natureza.';

const MESES_CHOICE = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
] as const;

export interface ICriarFieldsListaLancamentosContabeisResult {
  success: boolean;
  /** Nomes internos dos campos criados nesta execução. */
  criados: string[];
  /** Nomes internos que já existiam (não duplicados). */
  jaExistiam: string[];
  /** Avisos (ex.: lista de lookup criada, descrição do Title atualizada). */
  avisos: string[];
  error?: string;
}

function escapeXml(s: string): string {
  return s
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function assertCleanInternalName(name: string): void {
  if (!/^[A-Za-z][a-zA-Z0-9]*$/.test(name)) {
    throw new Error(`Nome interno inválido: ${name}`);
  }
}

function fieldNameAttrs(internalName: string, displayName: string): string {
  assertCleanInternalName(internalName);
  return `Name="${internalName}" StaticName="${internalName}" DisplayName="${escapeXml(displayName)}"`;
}

function descAttr(text: string): string {
  return `Description="${escapeXml(text)}"`;
}

function listGuidBraced(id: string | undefined): string {
  if (id === undefined || String(id).trim() === '') {
    throw new Error('Id da lista de lookup em falta.');
  }
  const s = String(id).replace(/[{}]/g, '');
  return `{${s}}`;
}

function mesesChoiceXml(): string {
  return MESES_CHOICE.map((m) => `<CHOICE>${escapeXml(m)}</CHOICE>`).join('');
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

async function criarCampoSeNaoExistir(
  list: ReturnType<SPFI['web']['lists']['getByTitle']>,
  internalName: string,
  schemaXml: string,
  criados: string[],
  jaExistiam: string[]
): Promise<void> {
  try {
    await list.fields.getByInternalNameOrTitle(internalName)();
    jaExistiam.push(internalName);
    return;
  } catch {
    /* não existe */
  }
  try {
    await list.fields.createFieldAsXml({ SchemaXml: schemaXml });
    criados.push(internalName);
  } catch (e) {
    if (isAlreadyExistsError(e)) {
      jaExistiam.push(internalName);
      return;
    }
    throw e;
  }
}

/**
 * Cria (se ainda não existirem) os campos da lista `LancamentosContabeis`.
 *
 * @param sp Instância PnPjs (`SPFI`), p.ex. `getSP()` após `getSP(context)` no WebPart.
 *
 * **AnexoValidador / AnexoSuporte:** listas SharePoint não têm coluna nativa de ficheiro por campo.
 * Aqui ficam como texto de várias linhas (URLs ou referências). Anexos genéricos do item continuam nos anexos da lista.
 *
 * **NaturezaOperacao:** lookup para a lista `NaturezasOperacao` (garantida com `ensure` antes do campo).
 */
export async function criarFieldsListaLancamentosContabeis(
  sp: SPFI
): Promise<ICriarFieldsListaLancamentosContabeisResult> {
  const criados: string[] = [];
  const jaExistiam: string[] = [];
  const avisos: string[] = [];

  try {
    const list = sp.web.lists.getByTitle(LISTA_LANCAMENTOS_CONTABEIS);
    await list.select('Id')();

    const meta = (await list.select('EnableAttachments')()) as { EnableAttachments?: boolean };
    if (meta.EnableAttachments !== true) {
      await list.update({ EnableAttachments: true });
      avisos.push('Anexos da lista foram ativados.');
    }

    jaExistiam.push('Title');
    try {
      await list.fields.getByInternalNameOrTitle('Title').update({
        Description: 'Nome do lançamento.',
      });
      avisos.push('Coluna Title: descrição atualizada.');
    } catch {
      /* ignore */
    }

    const natEnsure = await sp.web.lists.ensure(LISTA_NATUREZA_OPERACAO_LOOKUP, DESC_LISTA_NATUREZA, 100, false);
    if (natEnsure.created) {
      avisos.push(`Lista «${LISTA_NATUREZA_OPERACAO_LOOKUP}» criada para o lookup NaturezaOperacao.`);
    }

    const natList = sp.web.lists.getByTitle(LISTA_NATUREZA_OPERACAO_LOOKUP);
    const natMeta = (await natList.select('Id')()) as { Id?: string };
    const lookupListToken = listGuidBraced(natMeta.Id);

    await criarCampoSeNaoExistir(
      list,
      'Requester',
      `<Field Type="User" ${fieldNameAttrs('Requester', 'Requester')} UserSelectionMode="0" List="UserInfo" Required="FALSE" ${descAttr(
        'Requisitante; predefinido com o utilizador atual.'
      )}><Default>[me]</Default></Field>`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'FiscalYear',
      `<Field Type="Number" ${fieldNameAttrs('FiscalYear', 'Fiscal Year')} Min="2000" Max="2100" Percentage="FALSE" ${descAttr(
        'Ano fiscal (ex.: 2026).'
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'Month',
      `<Field Type="Choice" ${fieldNameAttrs('Month', 'Month')} FillInChoice="TRUE" ${descAttr(
        'Mês de referência.'
      )}><CHOICES>${mesesChoiceXml()}</CHOICES></Field>`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'Ledger',
      `<Field Type="Text" ${fieldNameAttrs('Ledger', 'Ledger')} MaxLength="255" Required="FALSE" ${descAttr(
        'Razão contábil / ledger.'
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'DocType',
      `<Field Type="Text" ${fieldNameAttrs('DocType', 'Doc Type')} MaxLength="255" Required="FALSE" ${descAttr(
        'Tipo de documento SAP.'
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'Amount',
      `<Field Type="Currency" ${fieldNameAttrs('Amount', 'Amount')} Decimals="2" LCID="1046" Required="FALSE" ${descAttr(
        'Valor do lançamento.'
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'Explicacao',
      `<Field Type="Note" ${fieldNameAttrs('Explicacao', 'Explicação')} NumLines="6" RichText="FALSE" AppendOnly="FALSE" Required="FALSE" ${descAttr(
        'Comentário do requisitante.'
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'AprovadorContabil',
      `<Field Type="User" ${fieldNameAttrs('AprovadorContabil', 'Aprovador Contábil')} UserSelectionMode="0" List="UserInfo" Required="FALSE" ${descAttr(
        'Aprovador (utilizador único).'
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'Validador',
      `<Field Type="User" ${fieldNameAttrs('Validador', 'Validador')} UserSelectionMode="0" List="UserInfo" Required="FALSE" ${descAttr(
        'Validador (utilizador único).'
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'DataAprovacao',
      `<Field Type="DateTime" ${fieldNameAttrs('DataAprovacao', 'Data de Aprovação')} DisplayFormat="${DateTimeFieldFormatType.DateTime}" FriendlyDisplayFormat="0" Required="FALSE" ${descAttr(
        'Data/hora da aprovação.'
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'DocNumberSAP',
      `<Field Type="Text" ${fieldNameAttrs('DocNumberSAP', 'Doc Number (SAP)')} MaxLength="255" Required="FALSE" ${descAttr(
        'Número do documento no SAP.'
      )} />`,
      criados,
      jaExistiam
    );

    /**
     * AnexoValidador: não existe coluna de ficheiro dedicada por campo em lista clássica.
     * Texto longo para URL(s) ou referência ao ficheiro; validação de .xlsx no formulário/app.
     */
    await criarCampoSeNaoExistir(
      list,
      'AnexoValidador',
      `<Field Type="Note" ${fieldNameAttrs('AnexoValidador', 'Anexo Validador')} NumLines="4" RichText="FALSE" AppendOnly="FALSE" Required="FALSE" ${descAttr(
        'Referência/URL do ficheiro validador (ex.: .xlsx). Ver anexos do item ou biblioteca se precisar de ficheiro real.'
      )} />`,
      criados,
      jaExistiam
    );

    /**
     * AnexoSuporte: idem; use anexos do item para vários ficheiros ou guarde URLs aqui.
     */
    await criarCampoSeNaoExistir(
      list,
      'AnexoSuporte',
      `<Field Type="Note" ${fieldNameAttrs('AnexoSuporte', 'Anexo Suporte')} NumLines="8" RichText="FALSE" AppendOnly="FALSE" Required="FALSE" ${descAttr(
        'URLs ou notas sobre documentos de suporte (contrato, PDF, etc.).'
      )} />`,
      criados,
      jaExistiam
    );

    /**
     * NaturezaOperacao: Lookup para `LISTA_NATUREZA_OPERACAO_LOOKUP` (Id obtido acima).
     * Para apontar para outra lista, altere a chamada `ensure` / o título em `LISTA_NATUREZA_OPERACAO_LOOKUP`.
     */
    await criarCampoSeNaoExistir(
      list,
      'NaturezaOperacao',
      `<Field Type="Lookup" ${fieldNameAttrs('NaturezaOperacao', 'Natureza Operação')} List="${lookupListToken}" ShowField="Title" Required="FALSE" ${descAttr(
        `Lookup para a lista «${LISTA_NATUREZA_OPERACAO_LOOKUP}».`
      )} />`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'StatusAprovacao',
      `<Field Type="Choice" ${fieldNameAttrs('StatusAprovacao', 'Status Aprovação')} Required="FALSE" ${descAttr(
        'Estado da aprovação.'
      )}><CHOICES><CHOICE>Pendente</CHOICE><CHOICE>Aprovado</CHOICE><CHOICE>Reprovado</CHOICE></CHOICES><Default>Pendente</Default></Field>`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'StatusLancamento',
      `<Field Type="Choice" ${fieldNameAttrs('StatusLancamento', 'Status Lançamento')} Required="FALSE" ${descAttr(
        'Estado do lançamento.'
      )}><CHOICES><CHOICE>Rascunho</CHOICE><CHOICE>Em validação</CHOICE><CHOICE>Validado</CHOICE><CHOICE>Invalidado</CHOICE><CHOICE>Lançado</CHOICE></CHOICES><Default>Rascunho</Default></Field>`,
      criados,
      jaExistiam
    );

    await criarCampoSeNaoExistir(
      list,
      'StatusID',
      `<Field Type="Choice" ${fieldNameAttrs('StatusID', 'Status ID')} Required="FALSE" ${descAttr(
        'Estado geral do pedido.'
      )}><CHOICES><CHOICE>Em aberto</CHOICE><CHOICE>Concluído</CHOICE><CHOICE>Cancelado</CHOICE></CHOICES><Default>Em aberto</Default></Field>`,
      criados,
      jaExistiam
    );

    try {
      const req = await list.fields.getByInternalNameOrTitle('Requester')();
      if (!req.ReadOnlyField) {
        await list.fields.getByInternalNameOrTitle('Requester').update({ ReadOnlyField: true });
        avisos.push('Requester definido como só leitura na lista.');
      }
    } catch {
      /* ignore */
    }

    return { success: true, criados, jaExistiam, avisos };
  } catch (e) {
    const error = e instanceof Error ? e.message : String(e);
    return {
      success: false,
      criados,
      jaExistiam,
      avisos,
      error,
    };
  }
}
