import { getSP } from '../../../services/core/sp';
import type { AutomacaoCampanhaFormData } from './automacaoCampanhaTypes';
import { datetimeLocalToIso } from './automacaoCampanhaUtils';

const AUTOMACAO_CAMPANHA_LIST_TITLE = 'AutomacaoCampanhas';

export async function createAutomacaoCampanhaItem(data: AutomacaoCampanhaFormData): Promise<number> {
  const sp = getSP();

  if (!sp) {
    throw new Error('Contexto do SharePoint nao inicializado.');
  }

  const inicioIso = datetimeLocalToIso(data.Inicio);
  const fimIso = datetimeLocalToIso(data.Fim);

  const payload: Record<string, unknown> = {
    Title: data.Title.trim(),
    descricao_campanha: data.descricao_campanha.trim(),
    texto_regra: data.texto_regra.trim(),
    Tipo_campanha: data.Tipo_campanha.trim(),
    EnviarPara: data.EnviarPara.trim()
  };

  if (inicioIso) {
    payload.Inicio = inicioIso;
  }

  if (fimIso) {
    payload.Fim = fimIso;
  }

  try {
    const result = await sp.web.lists.getByTitle(AUTOMACAO_CAMPANHA_LIST_TITLE).items.add(payload);
    const r = result as { data?: { Id?: number; ID?: number }; Id?: number; ID?: number };
    const blob = r.data !== undefined && r.data !== null ? r.data : r;
    const id = blob.Id ?? blob.ID;
    if (id === undefined) {
      throw new Error('Resposta sem ID ao criar item em AutomacaoCampanhas.');
    }
    return Number(id);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Falha ao criar item em AutomacaoCampanhas.';
    throw new Error(message);
  }
}
