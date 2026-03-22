import type { AutomacaoCampanhaFormData } from './automacaoCampanhaTypes';

export const TIPO_CAMPANHA_OPCOES = ['Incentivo', 'Comparativo', 'Ranking', 'Elegibilidade'];
export const ENVIAR_PARA_OPCOES = ['Vendedor', 'Loja', 'Gerente'];

const pad2 = (n: number): string => (n < 10 ? `0${n}` : String(n));

const toDatetimeLocalValue = (value: unknown): string => {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  if (typeof value === 'string') {
    const d = new Date(value);
    if (!isNaN(d.getTime())) {
      return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
    }
  }

  if (typeof value === 'number' && !isNaN(value)) {
    const d = new Date(value);
    if (!isNaN(d.getTime())) {
      return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}T${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
    }
  }

  return '';
};

const getRowString = (row: Record<string, unknown>, keys: string[]): string => {
  const rowKeys = Object.keys(row);
  for (let k = 0; k < keys.length; k++) {
    const keyLower = keys[k].toLowerCase();
    for (let r = 0; r < rowKeys.length; r++) {
      const rk = rowKeys[r];
      if (rk.toLowerCase() === keyLower) {
        const val = row[rk];
        if (val !== null && val !== undefined) {
          return String(val);
        }
      }
    }
  }
  return '';
};

export const buildAutomacaoCampanhaInitialForm = (
  textoConsulta: string,
  respostaPBI: string
): AutomacaoCampanhaFormData => {
  const base: AutomacaoCampanhaFormData = {
    Title: '',
    TextoConsulta: textoConsulta,
    descricao_campanha: '',
    texto_regra: textoConsulta,
    Inicio: '',
    Fim: '',
    Tipo_campanha: '',
    EnviarPara: ''
  };

  const trimmed = respostaPBI.trim();
  if (!trimmed) {
    return base;
  }

  try {
    const parsed = JSON.parse(trimmed) as unknown;
    if (!Array.isArray(parsed) || parsed.length === 0) {
      return base;
    }

    const first = parsed[0];
    if (first === null || typeof first !== 'object') {
      return base;
    }

    const row = first as Record<string, unknown>;

    return {
      Title: getRowString(row, ['Title', 'Título', 'Titulo', 'titulo', 'Nome', 'nome']),
      TextoConsulta: textoConsulta,
      descricao_campanha: getRowString(row, ['descricao_campanha', 'Descricao_campanha', 'descricao', 'Descricao']),
      texto_regra: textoConsulta,
      Inicio: toDatetimeLocalValue(row.Inicio ?? row.inicio ?? row.DataInicio),
      Fim: toDatetimeLocalValue(row.Fim ?? row.fim ?? row.DataFim),
      Tipo_campanha: getRowString(row, ['Tipo_campanha', 'tipo_campanha', 'TipoCampanha']),
      EnviarPara: getRowString(row, ['EnviarPara', 'enviarPara', 'Enviar_para'])
    };
  } catch {
    return base;
  }
};

export const datetimeLocalToIso = (value: string): string | undefined => {
  if (!value.trim()) {
    return undefined;
  }
  const d = new Date(value);
  if (isNaN(d.getTime())) {
    return undefined;
  }
  return d.toISOString();
};
