export interface AutomacaoCampanhaFormData {
  Title: string;
  TextoConsulta: string;
  descricao_campanha: string;
  texto_regra: string;
  Inicio: string;
  Fim: string;
  Tipo_campanha: string;
  EnviarPara: string;
}

export interface AutomacaoCampanhaFormErrors {
  Title?: string;
  TextoConsulta?: string;
  descricao_campanha?: string;
  texto_regra?: string;
  Inicio?: string;
  Fim?: string;
  Tipo_campanha?: string;
  EnviarPara?: string;
}

export const EMPTY_AUTOMACAO_CAMPANHA_FORM: AutomacaoCampanhaFormData = {
  Title: '',
  TextoConsulta: '',
  descricao_campanha: '',
  texto_regra: '',
  Inicio: '',
  Fim: '',
  Tipo_campanha: '',
  EnviarPara: ''
};
