export interface ValidationTemplateItem {
  Id: number;
  ID?: number;
  Title: string;
  TextoConsulta: string;
  Status: string;
  RespostaPBI: string;
  Modified?: string;
  Created?: string;
  GUID?: string;
}

export interface ValidationPhaseOneFormData {
  textoConsulta: string;
}

export interface ValidationPhaseOneErrors {
  textoConsulta?: string;
}

export interface CreateValidationTemplateItemInput {
  Title: string;
  TextoConsulta: string;
  Status: string;
}
