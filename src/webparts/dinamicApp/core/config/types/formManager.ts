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
  /**
   * Números: `{{Campo}}`, operadores + - * / ( ).
   * Texto: prefixo `str:` com `{{Campo}}` e tokens dinâmicos entre colchetes, ex. `[me]`, `[myEmail]`, `[today]`, `[query:chave]`.
   * Só token: `[myName]` (valor único).
   * Pasta de anexos (biblioteca): `attfolder:nodeId` com id do nó configurado na árvore de pastas em Anexos.
   */
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
  /** Extensões permitidas, sem ponto, minúsculas (ex.: pdf, docx). Vazio / omitido = qualquer extensão. */
  allowedFileExtensions?: string[];
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

/** Id fixo da etapa «Ocultos»: campos só no payload / metadados, sem UI no formulário. */
export const FORM_OCULTOS_STEP_ID = 'ocultos';

/** Id sintético do botão de histórico integrado (config na aba Lista de logs). */
export const FORM_BUILTIN_HISTORY_BUTTON_ID = '__builtin_history';

/** Como apresentar o botão de histórico de versões no formulário. */
export type TFormHistoryButtonKind = 'text' | 'icon' | 'iconAndText';

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
  /** Modos em que a etapa entra no passador. Omitido ou vazio = Criar, Editar e Ver. */
  showInFormModes?: TFormManagerFormMode[];
}

/** Navegação visual entre etapas no formulário (várias etapas). */
export type TFormStepLayoutKind =
  | 'rail'
  | 'segmented'
  | 'timeline'
  | 'cards'
  | 'breadcrumb'
  | 'underline'
  | 'outline'
  | 'compact'
  | 'steps'
  | 'minimal';

/** Estilo dos botões «Etapa anterior» / «Próxima etapa» no rodapé (independe do layout do passador de etapas). */
export type TFormStepNavButtonsKind =
  | 'fluent'
  | 'pills'
  | 'dots'
  | 'icons'
  | 'links'
  | 'split'
  | 'stacked'
  | 'ghost'
  | 'toolbar'
  | 'compact';

/** Indicador ao carregar campos da lista ou item na vista formulário. */
export type TFormDataLoadingUiKind =
  | 'spinner'
  | 'spinnerLarge'
  | 'shimmer'
  | 'progress'
  | 'cardShimmer';

/** Indicador ao gravar (botões personalizados). */
export type TFormSubmitLoadingUiKind =
  | 'overlay'
  | 'topProgress'
  | 'formShimmer'
  | 'belowButtons'
  | 'infoBar';

/** Onde gravar ficheiros escolhidos no controlo «Anexos ao item». */
export type TFormAttachmentStorageKind = 'itemAttachments' | 'documentLibrary';

/** Nó da árvore de pastas (níveis 2+) abaixo da pasta com o ID do item na biblioteca. */
export interface IAttachmentLibraryFolderTreeNode {
  id: string;
  /** Texto fixo ou modelo com placeholders `{{Title}}`, `{{NomeInterno}}`, etc. */
  nameTemplate: string;
  children?: IAttachmentLibraryFolderTreeNode[];
  /** Pasta onde os ficheiros são gravados (um único nó na árvore). */
  uploadTarget?: boolean;
  /**
   * Etapa em que o input de ficheiros desta pasta aparece (guardado como array com um único id; mesmo layout global).
   * Configurações antigas com vários ids são reduzidas ao primeiro.
   * Se nenhum nó tiver isto preenchido, mantém-se um único uploader com `uploadTarget`.
   */
  showUploaderInStepIds?: string[];
  /** Modos de formulário em que o input desta pasta pode aparecer. Omitido ou vazio = todos. */
  showUploaderModes?: TFormManagerFormMode[];
  /** Só utilizadores nestes grupos (título) veem o input; omitido ou vazio = qualquer utilizador. */
  showUploaderGroupTitles?: string[];
  /** Condição sobre valores dos campos (mesma árvore que nas regras do formulário). Omitida = sem filtro por dados. */
  showUploaderWhen?: TFormConditionNode;
  /** Mínimo de ficheiros nesta pasta (já na biblioteca + pendentes). Omitido = sem mínimo. */
  minAttachmentCount?: number;
  /** Máximo de ficheiros nesta pasta (já na biblioteca + pendentes). Omitido = sem máximo. */
  maxAttachmentCount?: number;
}

/** Destino em biblioteca: upload de ficheiros com lookup à lista principal do formulário. */
export interface IFormManagerAttachmentLibraryConfig {
  /** Título da biblioteca de documentos no site. */
  libraryTitle?: string;
  /**
   * Campo Lookup (simples) na biblioteca que aponta para a **lista principal** do formulário.
   * Na gravação define-se `{Campo}Id` com o id do item na lista principal.
   */
  sourceListLookupFieldInternalName?: string;
  /**
   * Pastas **abaixo** da pasta de nível 1 (nome = ID do item): pode haver várias ao mesmo nível;
   * cada uma pode ter filhos e irmãos. `uploadTarget` define onde o upload grava.
   */
  folderTree?: IAttachmentLibraryFolderTreeNode[];
  /**
   * Lista linear legada (um único ramo vertical). Migrada para `folderTree` ao gravar/ler.
   * @deprecated Usar `folderTree`.
   */
  folderPathSegments?: string[];
}

/** Vista do controlo de anexos quando o campo «Anexos ao item» está no formulário. */
export type TFormAttachmentUploadLayoutKind =
  | 'default'
  | 'dropzone'
  | 'card'
  | 'ribbon'
  | 'compact';

/** Como cada ficheiro escolhido aparece na lista (pré-visualização). Omitido = nameAndSize. */
export type TFormAttachmentFilePreviewKind =
  | 'nameOnly'
  | 'nameAndSize'
  | 'iconAndName'
  | 'thumbnailAndName'
  | 'thumbnailLarge';

export type TFormCustomButtonBehavior = 'actionsOnly' | 'draft' | 'submit' | 'close';

/** Clique no botão de histórico integrado: openOnly = só painel; resto = como botões personalizados + abrir histórico. */
export type TFormHistoryIntegratedClickBehavior = 'openOnly' | TFormCustomButtonBehavior;

/** Largura do bloco do formulário na vista (aba Estrutura). */
export type TFormRootWidthMode = 'full' | 'percent';

/** Posição horizontal do bloco do formulário na área disponível. */
export type TFormRootHorizontalAlign = 'start' | 'center' | 'end';

/** Onde o painel de histórico de auditoria abre (aba Lista de logs). */
export type TFormHistoryPresentationKind = 'panel' | 'modal' | 'collapse';

/** Estilo da lista de registos dentro do painel de histórico (aba Componentes). */
export type TFormHistoryLayoutKind = 'list' | 'timeline' | 'cards' | 'compact';

/** Operação principal do botão personalizado (além das ações em cadeia). */
export type TFormCustomButtonOperation =
  | 'legacy'
  | 'redirect'
  | 'add'
  | 'update'
  | 'delete'
  | 'history';

export interface IFormButtonActionShowFields {
  kind: 'showFields';
  fields: string[];
  /**
   * Etapa (id de `steps`, não «Ocultos») onde os campos devem aparecer quando só estão em Ocultos.
   * Com várias etapas visíveis, deve ser definido para esses campos serem renderizados.
   */
  displayOnStepId?: string;
  /** Se definido, a ação só corre quando a condição for verdadeira (valores já mesclados das ações anteriores). */
  when?: TFormConditionNode;
}

export interface IFormButtonActionHideFields {
  kind: 'hideFields';
  fields: string[];
  when?: TFormConditionNode;
}

export interface IFormButtonActionSetFieldValue {
  kind: 'setFieldValue';
  field: string;
  /** Texto fixo, `str:…`, token `[me]`, ou `attfolder:idDoNo` (pasta Anexos), como em setComputed. */
  valueTemplate: string;
  when?: TFormConditionNode;
}

export interface IFormButtonActionJoinFields {
  kind: 'joinFields';
  targetField: string;
  /**
   * Texto livre com `{{NomeInterno}}` substituído pelos valores atuais.
   * Se, após trim, não for vazio, tem prioridade sobre `sourceFields` + `separator`.
   */
  valueTemplate?: string;
  /** Ordem dos valores quando o modelo está vazio (modo legado). Mantida na UI para ordenar e inserir placeholders. */
  sourceFields: string[];
  /** Usado só quando `valueTemplate` está vazio (modo legado). */
  separator: string;
  when?: TFormConditionNode;
}

export type TFormButtonAction =
  | IFormButtonActionShowFields
  | IFormButtonActionHideFields
  | IFormButtonActionSetFieldValue
  | IFormButtonActionJoinFields;

export interface IFormCustomButtonConfig {
  id: string;
  label: string;
  /** Só usado em operation === 'history': texto curto (subtítulo / ajuda). */
  shortDescription?: string;
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
  /** false: botão ignorado na UI. */
  enabled?: boolean;
  /** Visibilidade condicional (omitido = sem filtro por dados). */
  when?: TFormConditionNode;
  /** Títulos de grupos SharePoint; vazio/omitido = qualquer usuário. */
  groupTitles?: string[];
  /**
   * Se true, o botão só aparece quando todos os campos obrigatórios visíveis estão preenchidos
   * (regras + obrigatório na lista; anexos se obrigatórios). Cumulativo com grupos e condição «when».
   */
  showOnlyWhenAllRequiredFilled?: boolean;
  /** Loading ao gravar; omitido usa `defaultSubmitLoadingKind` do gestor. */
  submitLoadingKind?: TFormSubmitLoadingUiKind;
  actions: TFormButtonAction[];
}

export interface IFormStepNavigationConfig {
  requireFilledRequiredToAdvance?: boolean;
  fullValidationOnAdvance?: boolean;
  allowBackWithoutValidation?: boolean;
}

/** Registo de auditoria: lista de destino e texto por botão (HTML). */
export interface IFormManagerActionLogConfig {
  /** Quando true, o runtime pode gravar entradas na lista configurada. */
  captureEnabled?: boolean;
  /** Título da lista SharePoint (não biblioteca) onde gravar logs. */
  listTitle?: string;
  /**
   * Nome interno do campo **várias linhas** na lista de log onde se grava o texto da ação (metadata).
   * Obrigatório para `captureEnabled`; só colunas multilinha são oferecidas na UI.
   */
  actionFieldInternalName?: string;
  /**
   * Nome interno do campo **Lookup** (simples) na lista de log que aponta para a **lista principal** do formulário.
   * Na gravação define-se `{Campo}Id` com o id do item na lista principal.
   */
  sourceListLookupFieldInternalName?: string;
  /** HTML (editor rich) por id de `customButtons`. */
  descriptionsHtmlByButtonId?: Record<string, string>;
}

export interface IFormManagerConfig {
  sections: IFormSectionConfig[];
  fields: IFormFieldConfig[];
  rules: TFormRule[];
  steps?: IFormStepConfig[];
  /**
   * Largura do formulário na vista. Omitido = legacy (≈720px, alinhado ao início).
   * `full` = 100% da área; `percent` = largura em % com `formRootWidthPercent`.
   */
  formRootWidthMode?: TFormRootWidthMode;
  /** 1–100. Usado com `formRootWidthMode === 'percent'`. Omitido = 100. */
  formRootWidthPercent?: number;
  /** Alinhamento do bloco do formulário. Omitido = start (legacy). */
  formRootHorizontalAlign?: TFormRootHorizontalAlign;
  stepNavigation?: IFormStepNavigationConfig;
  /** Colunas da grade gestor (usa mesma origem que listView se vazio) */
  managerColumnFields?: string[];
  /** Ajuda dinâmica por campo quando condição */
  dynamicHelp?: { field: string; when: TFormConditionNode; helpText: string }[];
  /** Botões com ações ao clicar (mostrar/ocultar campos, valores, juntar campos) */
  customButtons?: IFormCustomButtonConfig[];
  /** Lista e textos para registo de auditoria por botão. */
  actionLog?: IFormManagerActionLogConfig;
  /** Apresentação das etapas quando há mais de uma */
  stepLayout?: TFormStepLayoutKind;
  /** Estilo dos botões anterior/próximo etapa no rodapé */
  stepNavButtons?: TFormStepNavButtonsKind;
  /** Indicador ao carregar dados do formulário (lista / item). */
  formDataLoadingKind?: TFormDataLoadingUiKind;
  /** Padrão de loading ao gravar quando o botão não define override. */
  defaultSubmitLoadingKind?: TFormSubmitLoadingUiKind;
  /** Se true, mostra o botão de histórico de auditoria (registos da lista de log filtrados pelo lookup ao item). */
  historyEnabled?: boolean;
  /** Onde abrir o painel de histórico de auditoria (padrão: painel lateral). */
  historyPresentationKind?: TFormHistoryPresentationKind;
  /** Aspeto dos registos no painel: lista, linha do tempo, cartões ou compacto (padrão: list). */
  historyLayoutKind?: TFormHistoryLayoutKind;
  /** Apresentação do botão de histórico integrado. Omitido = só texto. */
  historyButtonKind?: TFormHistoryButtonKind;
  /** Texto do botão (ou tooltip se só ícone). Padrão: «Histórico». */
  historyButtonLabel?: string;
  /** Nome do ícone Fluent (ex.: History). Usado em `icon` e `iconAndText`. */
  historyButtonIcon?: string;
  /** Subtítulo no painel de histórico e tooltip no botão. */
  historyPanelSubtitle?: string;
  /** Grupos SharePoint (títulos) que podem ver o botão de histórico integrado. Vazio = todos. */
  historyGroupTitles?: string[];
  /**
   * O que fazer ao clicar no histórico integrado. Omitido = actionsOnly (ações + log se ativo + abrir painel).
   * openOnly = abrir só o painel, sem ações, sem gravar item, sem registo de log.
   */
  /** Legado: ignorado pelo runtime; o botão de histórico integrado só abre o painel. */
  historyButtonClickBehavior?: TFormHistoryIntegratedClickBehavior;
  /** Legado: ignorado pelo runtime. */
  historyButtonActions?: TFormButtonAction[];
  /**
   * Onde gravar os ficheiros do controlo «Anexos ao item». Omitido = anexos nativos do item na lista principal.
   */
  attachmentStorageKind?: TFormAttachmentStorageKind;
  /** Só com `attachmentStorageKind === 'documentLibrary'`. Biblioteca e lookup à lista principal. */
  attachmentLibrary?: IFormManagerAttachmentLibraryConfig;
  /** Layout visual do campo de ficheiros anexos (aba Componentes). Omitido = default. */
  attachmentUploadLayout?: TFormAttachmentUploadLayoutKind;
  /** Lista de ficheiros selecionados: nome, miniatura, ícone, etc. Omitido = nameAndSize. */
  attachmentFilePreview?: TFormAttachmentFilePreviewKind;
}
