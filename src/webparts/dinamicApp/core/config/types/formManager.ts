export type TFormManagerFormMode = 'create' | 'edit' | 'view';

export type TFormConditionOp =
  | 'eq'
  | 'ne'
  | 'gt'
  | 'ge'
  | 'lt'
  | 'le'
  | 'contains'
  | 'startsWith'
  | 'endsWith'
  | 'isEmpty'
  | 'isFilled'
  | 'isTrue'
  | 'isFalse';

export type TFormCompareKind = 'literal' | 'field' | 'token';

export interface IFormCompareRef {
  kind: TFormCompareKind;
  value: string;
}

export type TFormConditionNode =
  | { kind: 'all'; children: TFormConditionNode[] }
  | { kind: 'any'; children: TFormConditionNode[] }
  | { kind: 'leaf'; field: string; op: TFormConditionOp; compare?: IFormCompareRef };

export type TFormRuleTargetKind = 'field' | 'section';

export type TFormVisibilityIntent = 'show' | 'hide';

export type TFormSubmitKind = 'draft' | 'submit';

export interface IFormRuleBase {
  id: string;
  enabled?: boolean;
  when?: TFormConditionNode;
  modes?: TFormManagerFormMode[];
  /** Só aplica se o usuário estiver em algum destes grupos (título). Vazio = todos. */
  groupTitles?: string[];
  /** Regras com `fullSubmitOnly` não rodam em rascunho. */
  tags?: string[];
}

export interface IFormRuleSetVisibility extends IFormRuleBase {
  action: 'setVisibility';
  targetKind: TFormRuleTargetKind;
  targetId: string;
  visibility: TFormVisibilityIntent;
}

export interface IFormRuleSetRequired extends IFormRuleBase {
  action: 'setRequired';
  field: string;
  required: boolean;
}

export interface IFormRuleSetDisabled extends IFormRuleBase {
  action: 'setDisabled';
  field: string;
  disabled: boolean;
}

export interface IFormRuleSetReadOnly extends IFormRuleBase {
  action: 'setReadOnly';
  field: string;
  readOnly: boolean;
}

export interface IFormRuleClearFields extends IFormRuleBase {
  action: 'clearFields';
  fields: string[];
  /** Quando este campo muda de valor, limpa `fields` (ignora `when` para o disparo). */
  triggerField?: string;
}

export interface IFormRuleSetDefault extends IFormRuleBase {
  action: 'setDefault';
  field: string;
  value: string;
}

export interface IFormRuleValidateValue extends IFormRuleBase {
  action: 'validateValue';
  field: string;
  minNumber?: number;
  maxNumber?: number;
  minLength?: number;
  maxLength?: number;
  pattern?: string;
  patternMessage?: string;
  allowList?: string[];
  denyList?: string[];
  message?: string;
}

export interface IFormRuleValidateDate extends IFormRuleBase {
  action: 'validateDate';
  field: string;
  minIso?: string;
  maxIso?: string;
  minDaysFromToday?: number;
  maxDaysFromToday?: number;
  blockWeekends?: boolean;
  blockedIsoDates?: string[];
  gteField?: string;
  lteField?: string;
  gtField?: string;
  ltField?: string;
  message?: string;
}

export interface IFormRuleAtLeastOne extends IFormRuleBase {
  action: 'atLeastOne';
  fields: string[];
  message?: string;
}

export interface IFormRuleMultiMinMax extends IFormRuleBase {
  action: 'multiMinMax';
  field: string;
  min?: number;
  max?: number;
  message?: string;
}

export interface IFormRuleShowMessage extends IFormRuleBase {
  action: 'showMessage';
  variant: 'info' | 'warning' | 'error';
  text: string;
}

export interface IFormRuleFilterLookup extends IFormRuleBase {
  action: 'filterLookupOptions';
  field: string;
  parentField: string;
  /** OData filter com `{parent}` substituído pelo Id do pai */
  odataFilterTemplate: string;
}

export interface IFormRuleSetComputed extends IFormRuleBase {
  action: 'setComputed';
  field: string;
  /** Expressão segura: números, + - * / ( ), nomes de campo, STR_concat(a,b), DAYS(a,b) */
  expression: string;
}

export interface IFormRuleProfileField extends IFormRuleBase {
  action: 'profileVisibility' | 'profileEditable' | 'profileRequired';
  field: string;
  groupTitles: string[];
  /** visibility: true = visível só para grupos; editable/required: true = só esses grupos */
  allow: boolean;
}

export interface IFormRuleAuthorField extends IFormRuleBase {
  action: 'authorFieldAccess';
  field: string;
}

export interface IFormRuleAttachment extends IFormRuleBase {
  action: 'attachmentRules';
  minCount?: number;
  maxCount?: number;
  maxBytesPerFile?: number;
  allowedMimeTypes?: string[];
  requiredWhen?: TFormConditionNode;
  message?: string;
}

export interface IFormRuleAsyncUniqueness extends IFormRuleBase {
  action: 'asyncUniqueness';
  field: string;
  listTitle?: string;
  message?: string;
}

export interface IFormRuleAsyncCountLimit extends IFormRuleBase {
  action: 'asyncCountLimit';
  listTitle?: string;
  filterTemplate: string;
  maxCount: number;
  message?: string;
}

export interface IFormRuleEffectiveSection extends IFormRuleBase {
  action: 'setEffectiveSection';
  field: string;
  sectionId: string;
}

export type TFormRule =
  | IFormRuleSetVisibility
  | IFormRuleSetRequired
  | IFormRuleSetDisabled
  | IFormRuleSetReadOnly
  | IFormRuleClearFields
  | IFormRuleSetDefault
  | IFormRuleValidateValue
  | IFormRuleValidateDate
  | IFormRuleAtLeastOne
  | IFormRuleMultiMinMax
  | IFormRuleShowMessage
  | IFormRuleFilterLookup
  | IFormRuleSetComputed
  | IFormRuleProfileField
  | IFormRuleAuthorField
  | IFormRuleAttachment
  | IFormRuleAsyncUniqueness
  | IFormRuleAsyncCountLimit
  | IFormRuleEffectiveSection;

/** Nome interno reservado para anexos ao item (não é coluna SharePoint). */
export const FORM_ATTACHMENTS_FIELD_INTERNAL = '__formAttachments';

export interface IFormFieldConfig {
  internalName: string;
  label?: string;
  helpText?: string;
  placeholder?: string;
  sectionId?: string;
  visible?: boolean;
  required?: boolean;
  disabled?: boolean;
  readOnly?: boolean;
  width?: 'full' | 'half';
  /** Campos neste grupo abrem em painel/modal */
  modalGroupId?: string;
  /** Seção efetiva quando condição (avaliada no motor com prefixo de regra dedicada) */
  effectiveSectionId?: string;
}

export interface IFormSectionConfig {
  id: string;
  title: string;
  visible?: boolean;
  collapsed?: boolean;
}

export interface IFormStepConfig {
  id: string;
  title: string;
  fieldNames: string[];
}

/** Navegação visual entre etapas no formulário (várias etapas). */
export type TFormStepLayoutKind = 'rail' | 'segmented' | 'timeline' | 'cards';

/** Estilo dos botões «Etapa anterior» / «Próxima etapa» no rodapé (independe do layout do passador de etapas). */
export type TFormStepNavButtonsKind = 'fluent' | 'pills' | 'dots' | 'icons' | 'links';

export type TFormCustomButtonBehavior = 'actionsOnly' | 'draft' | 'submit' | 'close';

/** Operação principal do botão personalizado (além das ações em cadeia). */
export type TFormCustomButtonOperation = 'legacy' | 'redirect' | 'add' | 'update' | 'delete';

export interface IFormButtonActionShowFields {
  kind: 'showFields';
  fields: string[];
}

export interface IFormButtonActionHideFields {
  kind: 'hideFields';
  fields: string[];
}

export interface IFormButtonActionSetFieldValue {
  kind: 'setFieldValue';
  field: string;
  /** Texto fixo ou expressão `str:{{Campo}}` (mesma sintaxe de setComputed texto) */
  valueTemplate: string;
}

export interface IFormButtonActionJoinFields {
  kind: 'joinFields';
  targetField: string;
  sourceFields: string[];
  separator: string;
}

export type TFormButtonAction =
  | IFormButtonActionShowFields
  | IFormButtonActionHideFields
  | IFormButtonActionSetFieldValue
  | IFormButtonActionJoinFields;

export interface IFormCustomButtonConfig {
  id: string;
  label: string;
  appearance?: 'primary' | 'default';
  behavior?: TFormCustomButtonBehavior;
  /** Omitido ou legacy: usa apenas `behavior` + ações. */
  operation?: TFormCustomButtonOperation;
  /** Só para operation === 'redirect'. Placeholders {{NomeInterno}} e especiais {{FormID}}, {{Form}} (modo: Display|Edit|New). */
  redirectUrlTemplate?: string;
  /** operation === 'delete': mostrar em modo ver (Disp). Omitido = true. */
  deleteShowInView?: boolean;
  /** operation === 'delete': mostrar em modo editar. Omitido = true. */
  deleteShowInEdit?: boolean;
  modes?: TFormManagerFormMode[];
  actions: TFormButtonAction[];
}

export interface IFormManagerConfig {
  sections: IFormSectionConfig[];
  fields: IFormFieldConfig[];
  rules: TFormRule[];
  steps?: IFormStepConfig[];
  /** Colunas da grade gestor (usa mesma origem que listView se vazio) */
  managerColumnFields?: string[];
  /** Ajuda dinâmica por campo quando condição */
  dynamicHelp?: { field: string; when: TFormConditionNode; helpText: string }[];
  /** Botões com ações ao clicar (mostrar/ocultar campos, valores, juntar campos) */
  customButtons?: IFormCustomButtonConfig[];
  /** Apresentação das etapas quando há mais de uma */
  stepLayout?: TFormStepLayoutKind;
  /** Estilo dos botões anterior/próximo etapa no rodapé */
  stepNavButtons?: TFormStepNavButtonsKind;
  /** Se true, mostra Enviar, Rascunho e Fechar além dos botões personalizados. */
  showDefaultFormButtons?: boolean;
}
