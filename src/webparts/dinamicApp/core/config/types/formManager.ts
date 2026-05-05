export type TFormManagerFormMode = 'create' | 'edit' | 'view';

/** Nomes internos de autoria e datas de sistema; só exibição e bloqueados no formulário. */
export const FORM_SYSTEM_LIST_METADATA_INTERNAL_NAMES = new Set<string>([
  'Author',
  'Editor',
  'Created',
  'Modified',
]);

/** Tag em `TFormRule.tags`: visibilidade com mesclagem «ocultar prevalece» entre regras marcadas. */
export const FORM_VISIBILITY_PREFER_HIDE_TAG = 'fmVisPreferHide';

export type TFormConditionOp =
  | 'eq'
  | 'ne'
  | 'gt'
  | 'ge'
  | 'lt'
  | 'le'
  | 'contains'
  | 'notContains'
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
  | { kind: 'leaf'; field: string; op: TFormConditionOp; compare?: IFormCompareRef }
  | { kind: 'userGroup'; invert: boolean; groupTitle: string };

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
  /** Se preenchido, a regra **não** aplica a utilizadores que pertençam a algum destes grupos. */
  excludeGroupTitles?: string[];
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
  /** Dias inteiros; mesma sintaxe que valor padrão numérico (`{{Campo}}+10`). Se definido, prevalece sobre minDaysFromToday. */
  minDaysFromTodayExpr?: string;
  maxDaysFromTodayExpr?: string;
  blockWeekends?: boolean;
  /** 0=domingo … 6=sábado (Date.getDay). Datas nesse dia são inválidas. */
  blockedWeekdays?: number[];
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

export type TLookupFilterOperator = 'eq' | 'ne' | 'lt' | 'le' | 'gt' | 'ge' | 'contains' | 'startsWith';

export interface IFormRuleFilterLookup extends IFormRuleBase {
  action: 'filterLookupOptions';
  field: string;
  parentField: string;
  /** Campo na lista ligada a comparar com o valor do campo pai. */
  childField?: string;
  /** Operador OData para a comparação visual. */
  filterOperator?: TLookupFilterOperator;
  /** Legado: modelo OData com `{parent}` substituído pelo Id numérico do campo pai. */
  odataFilterTemplate?: string;
}

export interface IFormRuleSetComputed extends IFormRuleBase {
  action: 'setComputed';
  field: string;
  /**
   * Se true, o valor calculado substitui o controlo em todos os modos (valor gravado ignorado).
   * Omitido ou false: em criação mostra a expressão ao vivo; em edição/visualização mostra o valor gravado,
   * exceto se a expressão da regra for alterada desde a abertura do item (aí volta a calcular ao vivo).
   */
  alwaysLiveComputed?: boolean;
  /**
   * Números: `{{Campo}}`, operadores + - * / ( ).
   * Texto: prefixo `str:` com `{{Campo}}` e tokens dinâmicos entre colchetes, ex. `[me]`, `[myEmail]`, `[today]`, `[query:chave]`.
   * Sem `str:`: vários tokens e/ou `{{Campo}}` com literais (ex. `[me]-[me]`, `[myLogin]-[myEmail]`); o resultado é texto (não interpreta `-` entre números como subtração).
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

/** Prefixo de nomes sintéticos de banner (imagem por URL; não é coluna SharePoint). */
export const FORM_BANNER_INTERNAL_PREFIX = '__formBanner_';

export type TFormAlertVariant = 'info' | 'success' | 'warning' | 'error';

export type TFormFieldConfigKind = 'field' | 'banner' | 'alert';

/** Onde o banner aparece: na sequência da etapa, ou fixo (sticky) no topo/rodapé do formulário. */
export type TFormBannerPlacement = 'inStep' | 'topFixed' | 'bottomFixed';

/** Id fixo da etapa «Ocultos»: campos só no payload / metadados, sem UI no formulário. */
export const FORM_OCULTOS_STEP_ID = 'ocultos';

/** Id fixo da etapa «Fixos»: campos fixos no topo ou rodapé (fora do passador). */
export const FORM_FIXOS_STEP_ID = 'fixos';

/** Onde um campo da etapa Fixos aparece no formulário. */
export type TFixedChromePlacement = 'top' | 'bottom';

/** Como o bloco se posiciona na zona fixa (Fixos ou banner top/bottom fixo). */
export type TChromePositionMode = 'sticky' | 'absolute' | 'flow';

/** Id sintético do botão de histórico integrado (ativar na aba Componentes; lista de log na aba Lista de logs). */
export const FORM_BUILTIN_HISTORY_BUTTON_ID = '__builtin_history';

/** Como apresentar o botão de histórico de versões no formulário. */
export type TFormHistoryButtonKind = 'text' | 'icon' | 'iconAndText';

/** Operadores usados na UI de condicionais de exibição (campo texto). */
export type TTextFieldConditionalDisplayOp =
  | 'eq'
  | 'ne'
  | 'contains'
  | 'notContains'
  | 'isEmpty'
  | 'isFilled';

/** Junção de condições dentro de um grupo de regra condicional (texto). */
export type TTextFieldConditionalGroupOp = 'all' | 'any';

export type TTextFieldConditionalAction = 'show' | 'hide' | 'disable';

export interface ITextFieldConditionalCondition {
  id: string;
  refField: string;
  op: TTextFieldConditionalDisplayOp;
  compareKind: TFormCompareKind;
  compareValue: string;
}

export interface ITextFieldConditionalGroup {
  id: string;
  /** Omitido ou vazio = aplicar em Criar, Editar e Ver. */
  modes: TFormManagerFormMode[];
  /** Omitido ou vazio = todos os utilizadores; caso contrário, pelo menos um grupo. */
  groupTitles?: string[];
  /** Utilizadores nestes grupos não entram nesta regra (opcional). */
  excludeGroupTitles?: string[];
  groupOp: TTextFieldConditionalGroupOp;
  conditions: ITextFieldConditionalCondition[];
  action: TTextFieldConditionalAction;
  /** Opcional: sobrepõe `action` para criar, editar ou só visualização. */
  actionByMode?: Partial<Record<TFormManagerFormMode, TTextFieldConditionalAction>>;
}

export interface ITextFieldConditionalVisibility {
  groups: ITextFieldConditionalGroup[];
}

/** Colunas ocupadas numa grelha de 12 (estilo Bootstrap). */
export type TFormFieldColumnSpan = 3 | 4 | 6 | 8 | 12;

/** Transformação de valor de texto (maiúsculas / minúsculas / capitalizar por palavra). */
export type TFormFieldTextValueTransform = 'uppercase' | 'lowercase' | 'capitalize';

/** Máscara de input para coluna texto (IMask / react-imask). */
export type TFormFieldTextInputMaskKind = 'cpf' | 'telefone' | 'cep' | 'cnpj' | 'custom';

export interface IFormFieldConfig {
  internalName: string;
  /** `banner` = imagem só por URL no formulário; não corresponde a coluna na lista. */
  fieldKind?: TFormFieldConfigKind;
  /** URL da imagem quando `fieldKind === 'banner'`. */
  bannerImageUrl?: string;
  /** Só para `fieldKind === 'alert'`. */
  alertVariant?: TFormAlertVariant;
  /** Só para `fieldKind === 'alert'`. */
  alertTitle?: string;
  /** Só para `fieldKind === 'alert'`. */
  alertMessage?: string;
  /** Só para `fieldKind === 'alert'`. */
  alertIconName?: string;
  /** Só para `fieldKind === 'alert'`. */
  alertFields?: string[];
  /** Só para `fieldKind === 'alert'`. */
  alertWhen?: TFormConditionNode;
  /** Só para `fieldKind === 'alert'`. Omitido = `inStep`. */
  alertPlacement?: TFormBannerPlacement;
  /** Só para `fieldKind === 'alert'`. */
  alertDismissible?: boolean;
  /** Só para `fieldKind === 'alert'`. */
  alertEmphasized?: boolean;
  /** Só para `fieldKind === 'banner'`. Omitido = `inStep`. */
  bannerPlacement?: TFormBannerPlacement;
  /** Largura da imagem em % do contentor do formulário (1–100). Omitido = 100. */
  bannerWidthPercent?: number;
  /** Altura do banner em px. */
  bannerHeightPx?: number;
  /** Só na etapa «Fixos»: fixar no topo ou rodapé do formulário. */
  fixedPlacement?: TFixedChromePlacement;
  /** Fixos ou banner fixo no topo/rodapé: sticky, absoluto ao contentor, ou fluxo normal na zona. */
  chromePositionMode?: TChromePositionMode;
  label?: string;
  helpText?: string;
  placeholder?: string;
  /** Linhas visíveis do textarea (coluna Nota / multiline). */
  textareaRows?: number;
  sectionId?: string;
  visible?: boolean;
  required?: boolean;
  disabled?: boolean;
  readOnly?: boolean;
  width?: 'full' | 'half';
  /** Largura na grelha 12; legado: `width: half` → 6. */
  columnSpan?: TFormFieldColumnSpan;
  /** Sobrepõe `columnSpan` por modo (Criar, Editar, Ver). */
  columnSpanByMode?: Partial<Record<TFormManagerFormMode, TFormFieldColumnSpan>>;
  /** Campos neste grupo abrem em painel/modal */
  modalGroupId?: string;
  /** Seção efetiva quando condição (avaliada no motor com prefixo de regra dedicada) */
  effectiveSectionId?: string;
  /** Maiúsculas / minúsculas / capitalizar por palavra em colunas texto ou nota. */
  textValueTransform?: TFormFieldTextValueTransform;
  /** Máscara ativa para campo texto; omitido = sem máscara. */
  textInputMaskKind?: TFormFieldTextInputMaskKind;
  /** Padrão IMask quando `textInputMaskKind === 'custom'`; pode permanecer guardado ao mudar o kind. */
  textInputMaskCustomPattern?: string;
  /** Regras condicionais de visibilidade (só aplicadas a campos texto na UI de regras). */
  textConditionalVisibility?: ITextFieldConditionalVisibility;
  /**
   * Lookup / lookup multi: campo interno na lista de destino para o texto das opções na lista suspensa.
   * Omitido: usa LookupField definido na coluna SharePoint, ou Title.
   */
  lookupOptionLabelField?: string;
  /**
   * Sub-propriedade a extrair quando o campo de etiqueta é do tipo user/usermulti/lookup/lookupmulti.
   * Omitido: usa Title (ou EMail / LookupValue como fallback automático).
   * Exemplos: 'Title', 'EMail', 'LoginName' (user); 'Title' (lookup).
   */
  lookupOptionLabelSubProp?: string;
  /**
   * Lookup / lookup multi: outros campos internos a solicitar ao REST na lista de destino
   * (Id incluído automaticamente; person/lookup aninhados via metadados).
   */
  lookupOptionExtraSelectFields?: string[];
  /**
   * Lookup / lookup multi: campos da lista ligada a mostrar em só leitura abaixo do dropdown quando há seleção.
   */
  lookupOptionDetailBelowFields?: string[];
}

export function isFormBannerFieldConfig(fc: Pick<IFormFieldConfig, 'internalName' | 'fieldKind'>): boolean {
  return fc.fieldKind === 'banner' || fc.internalName.indexOf(FORM_BANNER_INTERNAL_PREFIX) === 0;
}

export function isFormAlertFieldConfig(fc: Pick<IFormFieldConfig, 'fieldKind'>): boolean {
  return fc.fieldKind === 'alert';
}

export function resolveFieldColumnSpan(
  fc: Pick<IFormFieldConfig, 'columnSpan' | 'width' | 'columnSpanByMode'>,
  mode?: TFormManagerFormMode
): TFormFieldColumnSpan {
  if (mode) {
    const bm = fc.columnSpanByMode?.[mode];
    if (bm === 3 || bm === 4 || bm === 6 || bm === 8 || bm === 12) return bm;
  }
  const c = fc.columnSpan;
  if (c === 3 || c === 4 || c === 6 || c === 8 || c === 12) return c;
  if (fc.width === 'half') return 6;
  return 12;
}

export function resolveBannerPlacement(fc: IFormFieldConfig): TFormBannerPlacement {
  const p = fc.bannerPlacement;
  if (p === 'topFixed' || p === 'bottomFixed' || p === 'inStep') return p;
  return 'inStep';
}

export function resolveAlertPlacement(fc: IFormFieldConfig): TFormBannerPlacement {
  const p = fc.alertPlacement;
  if (p === 'topFixed' || p === 'bottomFixed' || p === 'inStep') return p;
  return 'inStep';
}

export function resolveAlertVariant(fc: IFormFieldConfig): TFormAlertVariant {
  const v = fc.alertVariant;
  if (v === 'info' || v === 'success' || v === 'warning' || v === 'error') return v;
  return 'info';
}

export function resolveBannerWidthPercent(fc: IFormFieldConfig): number {
  const n = fc.bannerWidthPercent;
  if (typeof n !== 'number' || !isFinite(n)) return 100;
  return Math.min(100, Math.max(1, n));
}

export function resolveBannerHeightPx(fc: IFormFieldConfig): number | undefined {
  const n = fc.bannerHeightPx;
  if (typeof n === 'number' && isFinite(n)) return Math.min(2000, Math.max(40, Math.floor(n)));
  const legacy = (fc as { bannerHeightPercent?: number }).bannerHeightPercent;
  if (typeof legacy === 'number' && isFinite(legacy)) return Math.min(2000, Math.max(40, Math.floor(legacy)));
  return undefined;
}

export function resolveFixedPlacement(fc: IFormFieldConfig): TFixedChromePlacement {
  const p = fc.fixedPlacement;
  if (p === 'top' || p === 'bottom') return p;
  return 'top';
}

export function resolveChromePositionMode(fc: IFormFieldConfig): TChromePositionMode {
  const m = fc.chromePositionMode;
  if (m === 'sticky' || m === 'absolute' || m === 'flow') return m;
  return 'sticky';
}

export function resolveTextareaRows(fc: Pick<IFormFieldConfig, 'textareaRows'>, fallback: number): number {
  const n = fc.textareaRows;
  if (typeof n !== 'number' || !isFinite(n)) return fallback;
  const r = Math.floor(n);
  if (r < 1) return fallback;
  return Math.min(50, r);
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
  /**
   * Só avaliado quando a etapa já passa pelo filtro de modos. Omitida = sempre (no contexto desses modos).
   * Mesmo modelo que `when` das regras (E/Ou em folhas sobre campos, grupos SharePoint).
   */
  showStepWhen?: TFormConditionNode;
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
   * Pastas **abaixo** da pasta com nome = ID do item: pode haver várias ao mesmo nível;
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

/** Onde o painel de histórico de auditoria abre (aba Componentes). */
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

/** Alinhado a `MessageBarType` (Fluent): realce do ícone no modal de confirmação. */
export type TFormCustomButtonConfirmKind = 'info' | 'success' | 'warning' | 'error' | 'blocked';

export interface IFormCustomButtonConfirmConfig {
  enabled?: boolean;
  kind?: TFormCustomButtonConfirmKind;
  /** Texto do modal (obrigatório para o modal aparecer no runtime). */
  message?: string;
  /**
   * Campo da lista principal a preencher no modal antes de confirmar.
   * Tipos suportados: texto, nota, número, moeda, sim/não, data, escolha, hiperligação.
   */
  promptFieldInternalName?: string;
}

/** Último efeito do botão após ações e o resto do fluxo (gravar, log, etc.), se não houve erro bloqueante. */
export type TFormCustomButtonFinishAfterRun =
  | { kind: 'redirect'; redirectUrlTemplate: string }
  | { kind: 'clearForm' };

/** Faixa vertical da barra de botões personalizados e histórico integrado. Omitido = inferior. */
export type TFormCustomButtonsBarVertical = 'top' | 'bottom';

/** Alinhamento horizontal dessa barra. Omitido = esquerda. */
export type TFormCustomButtonsBarHorizontal = 'left' | 'right';

/** Slots da paleta Fluent alinhados ao tema do site (SharePoint). `outline` = neutro (DefaultButton). */
export type TFormCustomButtonPaletteSlot =
  | 'outline'
  | 'themePrimary'
  | 'themeSecondary'
  | 'themeTertiary'
  | 'themeDark'
  | 'themeDarkAlt'
  | 'themeDarker'
  | 'themeLight'
  | 'themeLighter'
  | 'themeLighterAlt';

export interface IFormCustomButtonConfig {
  id: string;
  label: string;
  /** Só usado em operation === 'history': texto curto (subtítulo / ajuda). */
  shortDescription?: string;
  appearance?: 'primary' | 'default';
  /**
   * Cor de preenchimento a partir do tema. Omitido = só `appearance` (primary → themePrimary, default → outline).
   */
  themePaletteSlot?: TFormCustomButtonPaletteSlot;
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
  /** Botão oculto para utilizadores nestes grupos (opcional). */
  excludeGroupTitles?: string[];
  /**
   * Se true, o botão só aparece quando todos os campos obrigatórios visíveis estão preenchidos
   * (regras + obrigatório na lista; anexos se obrigatórios). Cumulativo com grupos e condição «when».
   */
  showOnlyWhenAllRequiredFilled?: boolean;
  /**
   * Se true, o botão só aparece em item já gravado quando o utilizador atual é o autor (criador) do item.
   * Cumulativo com grupos SharePoint, exclusões e condição «when». Em modo novo não aparece.
   */
  showOnlyForItemAuthor?: boolean;
  /** Loading ao gravar; omitido usa `defaultSubmitLoadingKind` do gestor. */
  submitLoadingKind?: TFormSubmitLoadingUiKind;
  /**
   * Se ativo com mensagem, mostra um modal antes de qualquer efeito do botão (ações, redirect, gravar, etc.).
   * «Cancelar» não executa nada; «Confirmar» segue o fluxo normal.
   */
  confirmBeforeRun?: IFormCustomButtonConfirmConfig;
  /** Redirecionar ou limpar o formulário após todo o fluxo do botão concluir com sucesso. */
  finishAfterRun?: TFormCustomButtonFinishAfterRun;
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
  /**
   * Cor de realce do registo (slot da paleta Fluent) por id de botão.
   * Omitido por botão → `themePrimary`.
   */
  descriptionPaletteSlotByButtonId?: Record<string, TFormCustomButtonPaletteSlot>;
  /**
   * Se true, nos botões personalizados «Atualizar», acrescenta ao HTML do log as alterações efetivas
   * dos campos (valor inicial ao abrir o item → valor gravado), omitindo quando não há diferença líquida.
   */
  automaticChangesOnUpdate?: boolean;
}

/** Corpo de formulário reutilizável (lista principal ou lista filha vinculada). */
export interface IFormBodyConfig {
  sections: IFormSectionConfig[];
  fields: IFormFieldConfig[];
  rules: TFormRule[];
  steps?: IFormStepConfig[];
}

/** Onde gravar ficheiros por linha da lista vinculada (além dos campos do mini-formulário). */
export type TLinkedChildAttachmentStorageKind =
  | 'none'
  | 'itemAttachments'
  | 'documentLibraryInheritMain'
  | 'documentLibraryCustom';

/** Como as linhas da lista vinculada são apresentadas no formulário principal. */
export type TLinkedChildRowsPresentationKind = 'stack' | 'table' | 'compact' | 'cards';

/** Lista secundária com mini-formulário e Lookup para o item da lista principal. */
export interface IFormLinkedChildFormConfig extends IFormBodyConfig {
  id: string;
  /** Título da lista SharePoint filha. */
  listTitle: string;
  /** Lookup simples na lista filha que aponta para a lista principal. */
  parentLookupFieldInternalName: string;
  /** Rótulo da secção no formulário. */
  title?: string;
  /** Ordem de exibição relativa a outras listas vinculadas (menor primeiro). */
  order?: number;
  minRows?: number;
  maxRows?: number;
  /** Legado: ignorado na vista do formulário (listas vinculadas são sempre expandidas). */
  collapsedDefault?: boolean;
  /** Omitido ou `stack` = blocos em coluna (comportamento original). */
  rowsPresentation?: TLinkedChildRowsPresentationKind;
  /**
   * Id da etapa do formulário principal (`formManager.steps`) onde o bloco aparece no passador.
   * Omitido = primeira etapa do passador (excl. Ocultos/Fixos).
   */
  mainFormStepId?: string;
  /**
   * Anexos por linha: nenhum, anexos nativos do item na lista filha, ou biblioteca (pastas da aba Anexos ou própria).
   * Omitido / `none` = sem bloco de ficheiros extra por linha.
   */
  childAttachmentStorageKind?: TLinkedChildAttachmentStorageKind;
  /**
   * Com `documentLibraryInheritMain`: nome interno do campo Lookup **na biblioteca da aba Anexos**
   * que aponta para a **lista filha** (`listTitle`). Obrigatório nesse modo.
   */
  childAttachmentLibraryLookupToChildListField?: string;
  /**
   * Com `documentLibraryCustom`: biblioteca + árvore; `sourceListLookupFieldInternalName` deve ser o Lookup
   * na biblioteca que aponta para a lista filha.
   */
  childAttachmentLibrary?: IFormManagerAttachmentLibraryConfig;
}

export type TFormPermissionBreakPrincipalKind = 'siteGroup' | 'user' | 'field';

export type TFormPermissionBreakFieldScope = 'main' | 'linked';

/** Uma linha de principal + nível de permissão SharePoint (nome da definição no site). */
export interface IFormPermissionBreakAssignment {
  id: string;
  kind: TFormPermissionBreakPrincipalKind;
  /** Nome exato da definição de permissão (ex.: Leitura, Contribute, Full Control). */
  roleDefinitionName: string;
  /** kind === siteGroup */
  siteGroupId?: number;
  siteGroupTitle?: string;
  /** kind === user — resolvido com web.ensureUser(Key do clientPeoplePicker). */
  userPickerKey?: string;
  userDisplayText?: string;
  /** kind === field */
  fieldScope?: TFormPermissionBreakFieldScope;
  /** Obrigatório quando fieldScope === linked */
  linkedFormId?: string;
  fieldInternalName?: string;
}

export interface IFormManagerPermissionBreakTargets {
  mainListItem?: boolean;
  /**
   * Ids de `linkedChildForms` a incluir. Omitido = todas as listas vinculadas configuradas.
   * Array vazio = nenhuma linha filha recebe ACL por este gestor.
   */
  linkedChildFormIds?: string[];
  /** Itens de ficheiro na biblioteca da aba Anexos com lookup ao item principal. */
  mainAttachmentLibraryFiles?: boolean;
  /** Por id de lista vinculada: ficheiros na biblioteca (herdada ou custom) ligados à linha filha. */
  linkedAttachmentLibraryFilesByFormId?: string[];
}

export interface IFormManagerPermissionBreakConfig {
  enabled?: boolean;
  /** true = copiar atribuições herdadas antes de limpar (raro). */
  copyInheritedAssignments?: boolean;
  /** Manter o autor (Created By) com um nível explícito após limpar. */
  retainAuthor?: boolean;
  /** Definição de permissão para o autor; omitido = Contribute. */
  authorRoleDefinitionName?: string;
  targets?: IFormManagerPermissionBreakTargets;
  assignments?: IFormPermissionBreakAssignment[];
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
  /** Espaço interior (padding) em px em torno do conteúdo do formulário. Omitido ou 0 = sem valor extra. */
  formRootPaddingPx?: number;
  stepNavigation?: IFormStepNavigationConfig;
  /** Colunas da grade gestor (usa mesma origem que listView se vazio) */
  managerColumnFields?: string[];
  /** Ajuda dinâmica por campo quando condição */
  dynamicHelp?: { field: string; when: TFormConditionNode; helpText: string }[];
  /** Botões com ações ao clicar (mostrar/ocultar campos, valores, juntar campos) */
  customButtons?: IFormCustomButtonConfig[];
  /** Faixa vertical dos botões personalizados e do histórico integrado. Omitido = inferior. */
  customButtonsBarVertical?: TFormCustomButtonsBarVertical;
  /** Alinhamento horizontal dessa faixa. Omitido = esquerda. */
  customButtonsBarHorizontal?: TFormCustomButtonsBarHorizontal;
  /** Lista e textos para registo de auditoria por botão. */
  actionLog?: IFormManagerActionLogConfig;
  /** Apresentação das etapas quando há mais de uma */
  stepLayout?: TFormStepLayoutKind;
  /** Cor de destaque do passador e botões de etapa (omitido = primária do tema). */
  stepAccentPaletteSlot?: TFormCustomButtonPaletteSlot;
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
  /** Biblioteca ativa em modo documentLibrary; em anexos ao item pode persistir a última config para reativar sem perder a árvore. */
  attachmentLibrary?: IFormManagerAttachmentLibraryConfig;
  /** Layout visual do campo de ficheiros anexos (aba Componentes). Omitido = default. */
  attachmentUploadLayout?: TFormAttachmentUploadLayoutKind;
  /** Lista de ficheiros selecionados: nome, miniatura, ícone, etc. Omitido = nameAndSize. */
  attachmentFilePreview?: TFormAttachmentFilePreviewKind;
  /** Mini-formulários em listas secundárias ligadas ao item principal via Lookup. */
  linkedChildForms?: IFormLinkedChildFormConfig[];
  /** Quebra de herança e atribuições após gravar (lista principal, vinculadas, opcional biblioteca). */
  permissionBreak?: IFormManagerPermissionBreakConfig;
}
