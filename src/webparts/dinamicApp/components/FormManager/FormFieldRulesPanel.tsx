import * as React from 'react';
import { useEffect, useState, useCallback, useRef, useLayoutEffect, useMemo } from 'react';
import { createPortal } from 'react-dom';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  TextField,
  type ITextField,
  PrimaryButton,
  DefaultButton,
  Checkbox,
  Dropdown,
  IDropdownOption,
  ChoiceGroup,
  type IChoiceGroupOption,
  Link,
  MessageBar,
  MessageBarType,
  Spinner,
} from '@fluentui/react';
import {
  GroupsService,
  FieldsService,
  filterSiteGroupsByNameQuery,
  type IFieldMetadata,
  type IGroupDetails,
} from '../../../../services';
import type {
  IFormFieldConfig,
  TFormFieldTextInputMaskKind,
  TFormManagerFormMode,
  TFormConditionOp,
  TFormRule,
  ITextFieldConditionalCondition,
  ITextFieldConditionalGroup,
  TTextFieldConditionalDisplayOp,
  TTextFieldConditionalGroupOp,
  TTextFieldConditionalAction,
  TFormCompareKind,
} from '../../core/config/types/formManager';
import { FORM_ATTACHMENTS_FIELD_INTERNAL, isFormBannerFieldConfig } from '../../core/config/types/formManager';
import {
  buildFieldUiRules,
  CONDITION_OP_OPTIONS,
  emptyFieldRuleEditorState,
  fieldRuleStateFromRules,
  mergeFieldRuleEditorState,
  newCardId,
  type IFieldRuleEditorState,
  type IWhenUi,
  templateFieldRulesDateNotPast,
  templateFieldRulesEmail,
} from '../../core/formManager/formManagerVisualModel';
import { FormManagerCollapseSection } from './FormManagerComponentsTab';
import { TEXT_INPUT_MASK_CUSTOM_MAX_LEN } from '../../core/formManager/formTextInputMasks';
import { isNoteFieldMeta } from '../../core/listView';

/** Portal de sugestões @ (fora do painel no DOM); ignorar em `Panel.onOuterClick`. */
export const FORM_FIELD_RULES_MENTION_PORTAL_ATTR = 'data-dinamic-rules-mention';

const TEXT_RULES_COLLAPSE_IDS = {
  display: 'textRulesDisplay',
  validation: 'textRulesValidation',
  transform: 'textRulesTransform',
  masks: 'textRulesMasks',
  conditionals: 'textRulesConditionals',
} as const;

const FIELD_RULES_DISABLE_ENABLE_SECTION_ID = 'fieldRulesDisableEnable';

const TEXT_MASK_CHOICE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'none', text: 'Nenhuma' },
  { key: 'cpf', text: 'CPF' },
  { key: 'telefone', text: 'Telefone (BR)' },
  { key: 'cep', text: 'CEP' },
  { key: 'cnpj', text: 'CNPJ' },
  { key: 'custom', text: 'Personalizada (IMask)' },
];

const TEXT_DISPLAY_OP_OPTS: { key: TTextFieldConditionalDisplayOp; text: string }[] = [
  { key: 'eq', text: 'Igual a' },
  { key: 'ne', text: 'Diferente de' },
  { key: 'contains', text: 'Contém' },
  { key: 'notContains', text: 'Não contém' },
  { key: 'isEmpty', text: 'Vazio' },
  { key: 'isFilled', text: 'Não vazio' },
];

const TEXT_COND_GROUP_OP_OPTS: IChoiceGroupOption[] = [
  { key: 'all', text: 'E (todas as condições)' },
  { key: 'any', text: 'OU (pelo menos uma)' },
];

function normSpGroupTitle(s: string): string {
  return s.trim().toLowerCase();
}

const LOOKUP_RULES_LABEL_TYPES: ReadonlyArray<IFieldMetadata['MappedType']> = [
  'text',
  'multiline',
  'choice',
  'multichoice',
  'number',
  'currency',
  'boolean',
  'datetime',
  'url',
  'lookup',
  'lookupmulti',
  'user',
  'usermulti',
];

function lookupRulesEligibleTargetFields(fields: IFieldMetadata[]): IFieldMetadata[] {
  const allow = new Set<string>(LOOKUP_RULES_LABEL_TYPES);
  return fields.filter(
    (f) =>
      !isNoteFieldMeta(f) &&
      f.InternalName !== 'Id' &&
      !f.Hidden &&
      allow.has(f.MappedType)
  );
}

function newTextConditionalCondition(defaultRefField: string): ITextFieldConditionalCondition {
  return {
    id: newCardId(),
    refField: defaultRefField,
    op: 'eq',
    compareKind: 'literal',
    compareValue: '',
  };
}

function newTextConditionalGroup(defaultRefField: string): ITextFieldConditionalGroup {
  return {
    id: newCardId(),
    modes: [],
    groupOp: 'all',
    conditions: [newTextConditionalCondition(defaultRefField)],
    action: 'show',
  };
}

export interface IFormFieldRulesPanelProps {
  isOpen: boolean;
  internalName: string;
  fieldConfig: IFormFieldConfig;
  meta: IFieldMetadata | undefined;
  rules: TFormRule[];
  fieldOptions: IDropdownOption[];
  /** Pastas da árvore em Anexos (biblioteca); para valor calculado = URL da pasta. */
  attachmentLibraryFolderOptions?: IDropdownOption[];
  /** Web da lista principal (lista de dados do formulário) para ler a lista ligada nos lookups. */
  lookupFieldsWebServerRelativeUrl?: string;
  onDismiss: () => void;
  onApply: (nextField: IFormFieldConfig, editor: IFieldRuleEditorState) => void;
}

const MODE_OPTS: { key: TFormManagerFormMode; label: string }[] = [
  { key: 'create', label: 'Criar' },
  { key: 'edit', label: 'Editar' },
  { key: 'view', label: 'Ver' },
];

const ALL_MODES: TFormManagerFormMode[] = ['create', 'edit', 'view'];

const SET_COMPUTED_CONTEXT_TOKENS: { literal: string; hint: string }[] = [
  { literal: '[me]', hint: 'Id numérico do utilizador atual' },
  { literal: '[myId]', hint: 'Igual a [me]' },
  { literal: '[myName]', hint: 'Nome do utilizador' },
  { literal: '[myEmail]', hint: 'E-mail do utilizador' },
  { literal: '[myLogin]', hint: 'Nome de início de sessão' },
  { literal: '[myDepartment]', hint: 'Departamento (se disponível)' },
  { literal: '[myJobTitle]', hint: 'Cargo (se disponível)' },
  { literal: '[siteTitle]', hint: 'Título do site' },
  { literal: '[siteUrl]', hint: 'URL do site' },
  { literal: '[listTitle]', hint: 'Título da lista' },
  { literal: '[today]', hint: 'Data de hoje (ISO)' },
  { literal: '[now]', hint: 'Data e hora atuais (ISO)' },
  { literal: '[tomorrow]', hint: 'Dia seguinte (ISO)' },
  { literal: '[yesterday]', hint: 'Dia anterior (ISO)' },
  { literal: '[startOfMonth]', hint: 'Primeiro dia do mês corrente' },
  { literal: '[endOfMonth]', hint: 'Último dia do mês corrente' },
  { literal: '[startOfYear]', hint: 'Primeiro dia do ano corrente' },
  { literal: '[endOfYear]', hint: 'Último dia do ano corrente' },
  { literal: '[empty]', hint: 'Texto vazio' },
  { literal: '[null]', hint: 'Valor nulo' },
  { literal: '[true]', hint: 'Booleano verdadeiro' },
  { literal: '[false]', hint: 'Booleano falso' },
  { literal: '[query:nome]', hint: 'Valor do parâmetro ?nome= na URL da página' },
];

type TMentionItem = {
  key: string;
  insert: string;
  primary: string;
  secondary: string;
};

function getActiveMentionRange(
  value: string,
  caret: number
): { from: number; to: number; filter: string } | undefined {
  if (caret < 1) return undefined;
  const before = value.slice(0, caret);
  const at = before.lastIndexOf('@');
  if (at === -1) return undefined;
  if (at > 0) {
    const prev = before[at - 1];
    if (
      prev !== ' ' &&
      prev !== '\n' &&
      prev !== '\t' &&
      prev !== '(' &&
      prev !== '[' &&
      prev !== ';' &&
      prev !== ',' &&
      prev !== '+' &&
      prev !== '-' &&
      prev !== '*' &&
      prev !== '/'
    ) {
      return undefined;
    }
  }
  const segment = before.slice(at + 1);
  if (/[\s\n]/.test(segment)) return undefined;
  return { from: at, to: caret, filter: segment };
}

function buildMentionItems(
  filter: string,
  fieldOptions: IDropdownOption[],
  attachmentLibraryFolderOptions: IDropdownOption[]
): TMentionItem[] {
  const f = filter.trim().toLowerCase();
  const match = (s: string): boolean => !f || s.toLowerCase().includes(f);
  const out: TMentionItem[] = [];
  for (let i = 0; i < SET_COMPUTED_CONTEXT_TOKENS.length; i++) {
    const row = SET_COMPUTED_CONTEXT_TOKENS[i];
    if (match(row.literal) || match(row.hint)) {
      out.push({
        key: `t-${row.literal}-${i}`,
        insert: row.literal,
        primary: row.literal,
        secondary: row.hint,
      });
    }
  }
  for (let i = 0; i < fieldOptions.length; i++) {
    const opt = fieldOptions[i];
    const k = String(opt.key);
    const ins = `{{${k}}}`;
    const lab = String(opt.text ?? k);
    if (match(k) || match(lab) || match(ins)) {
      out.push({
        key: `f-${k}-${i}`,
        insert: ins,
        primary: lab,
        secondary: ins,
      });
    }
  }
  for (let i = 0; i < attachmentLibraryFolderOptions.length; i++) {
    const opt = attachmentLibraryFolderOptions[i];
    const k = String(opt.key);
    const ins = `attfolder:${k}`;
    const lab = String(opt.text ?? k);
    if (match(k) || match(lab) || match(ins)) {
      out.push({
        key: `p-${k}-${i}`,
        insert: ins,
        primary: lab,
        secondary: ins,
      });
    }
  }
  return out;
}

function overflowScrollAncestors(start: HTMLElement | null): HTMLElement[] {
  const seen = new Set<HTMLElement>();
  const out: HTMLElement[] = [];
  let n: HTMLElement | null = start?.parentElement ?? null;
  while (n) {
    const st = window.getComputedStyle(n);
    if (
      /(auto|scroll|overlay)/.test(st.overflowY) ||
      /(auto|scroll|overlay)/.test(st.overflowX) ||
      /(auto|scroll|overlay)/.test(st.overflow)
    ) {
      if (!seen.has(n)) {
        seen.add(n);
        out.push(n);
      }
    }
    n = n.parentElement;
  }
  const root = document.documentElement;
  if (!seen.has(root)) out.push(root);
  return out;
}

type TSetComputedRulesBlockProps = {
  ed: IFieldRuleEditorState;
  setEd: React.Dispatch<React.SetStateAction<IFieldRuleEditorState>>;
  fieldOptions: IDropdownOption[];
  attachmentLibraryFolderOptions: IDropdownOption[];
  bordered?: boolean;
};

function SetComputedRulesBlock({
  ed,
  setEd,
  fieldOptions,
  attachmentLibraryFolderOptions,
  bordered,
}: TSetComputedRulesBlockProps): JSX.Element {
  const [formsExprOpen, setFormsExprOpen] = useState(false);
  const [mentionOpen, setMentionOpen] = useState(false);
  const [mentionRange, setMentionRange] = useState<{ from: number; to: number; filter: string } | null>(null);
  const [mentionHighlight, setMentionHighlight] = useState(0);
  const [mentionListPos, setMentionListPos] = useState<{
    top: number;
    left: number;
    width: number;
  } | null>(null);
  const tfRef = useRef<ITextField | null>(null);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const mentionPortalRef = useRef<HTMLDivElement | null>(null);
  const pendingCaretRef = useRef<number | undefined>(undefined);
  const mentionRangeRef = useRef<{ from: number; to: number; filter: string } | undefined>(undefined);

  mentionRangeRef.current = mentionRange ?? undefined;

  const measureMentionListPos = useCallback((): void => {
    const w = wrapRef.current;
    if (!w) return;
    const r = w.getBoundingClientRect();
    setMentionListPos({
      top: r.bottom + 4,
      left: r.left,
      width: r.width,
    });
  }, []);

  const mentionItems = useMemo(() => {
    if (!mentionOpen || !mentionRange) return [];
    return buildMentionItems(mentionRange.filter, fieldOptions, attachmentLibraryFolderOptions);
  }, [mentionOpen, mentionRange, fieldOptions, attachmentLibraryFolderOptions]);

  useLayoutEffect(() => {
    const p = pendingCaretRef.current;
    if (p === undefined || !tfRef.current) return;
    pendingCaretRef.current = undefined;
    const tf = tfRef.current;
    tf.focus();
    requestAnimationFrame(() => {
      try {
        tf.setSelectionRange(p, p);
      } catch {
        //
      }
    });
  }, [ed.computedExpression]);

  useLayoutEffect(() => {
    const show =
      mentionOpen && mentionItems.length > 0 && !ed.computedAttachmentFolderNodeId;
    if (!show) {
      setMentionListPos(null);
      return;
    }
    measureMentionListPos();
  }, [
    mentionOpen,
    mentionItems.length,
    ed.computedAttachmentFolderNodeId,
    ed.computedExpression,
    formsExprOpen,
    measureMentionListPos,
  ]);

  useEffect(() => {
    const show =
      mentionOpen && mentionItems.length > 0 && !ed.computedAttachmentFolderNodeId;
    if (!show) return;
    measureMentionListPos();
    const roots = overflowScrollAncestors(wrapRef.current);
    const upd = (): void => {
      measureMentionListPos();
    };
    roots.forEach((el) => el.addEventListener('scroll', upd, true));
    window.addEventListener('resize', upd);
    return () => {
      roots.forEach((el) => el.removeEventListener('scroll', upd, true));
      window.removeEventListener('resize', upd);
    };
  }, [
    mentionOpen,
    mentionItems.length,
    ed.computedAttachmentFolderNodeId,
    measureMentionListPos,
  ]);

  useEffect(() => {
    const onDocDown = (e: MouseEvent): void => {
      const t = e.target as Node;
      if (wrapRef.current?.contains(t) || mentionPortalRef.current?.contains(t)) return;
      setMentionOpen(false);
    };
    document.addEventListener('mousedown', onDocDown);
    return () => document.removeEventListener('mousedown', onDocDown);
  }, []);

  useEffect(() => {
    if (mentionOpen && mentionItems.length === 0) setMentionOpen(false);
  }, [mentionOpen, mentionItems.length]);

  useEffect(() => {
    setMentionHighlight(0);
  }, [mentionRange?.filter]);

  const applyMentionInsert = useCallback(
    (insertText: string): void => {
      const r = mentionRangeRef.current;
      if (!r) return;
      const cur = ed.computedExpression;
      const next = cur.slice(0, r.from) + insertText + cur.slice(r.to);
      pendingCaretRef.current = r.from + insertText.length;
      setMentionOpen(false);
      setMentionRange(null);
      setEd((p) => ({
        ...p,
        computedExpression: next,
        computedAttachmentFolderNodeId: '',
      }));
    },
    [ed.computedExpression, setEd]
  );

  const handleExprChange = useCallback(
    (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v: string | undefined): void => {
      const raw = v ?? '';
      const el = ev.target as HTMLTextAreaElement;
      const caret =
        typeof el.selectionStart === 'number' ? el.selectionStart : raw.length;
      const range = getActiveMentionRange(raw, caret);
      if (range) {
        const items = buildMentionItems(range.filter, fieldOptions, attachmentLibraryFolderOptions);
        if (items.length > 0) {
          setMentionRange(range);
          setMentionOpen(true);
          setMentionHighlight(0);
        } else {
          setMentionOpen(false);
          setMentionRange(null);
        }
      } else {
        setMentionOpen(false);
        setMentionRange(null);
      }
      setEd((p) => ({
        ...p,
        computedExpression: raw,
        computedAttachmentFolderNodeId: '',
      }));
    },
    [fieldOptions, attachmentLibraryFolderOptions, setEd]
  );

  const handleExprKeyDown = useCallback(
    (ev: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
      if (!mentionOpen || mentionItems.length === 0) return;
      if (ev.key === 'ArrowDown') {
        ev.preventDefault();
        setMentionHighlight((h) => Math.min(mentionItems.length - 1, h + 1));
      } else if (ev.key === 'ArrowUp') {
        ev.preventDefault();
        setMentionHighlight((h) => Math.max(0, h - 1));
      } else if (ev.key === 'Enter' && !ev.shiftKey) {
        ev.preventDefault();
        const it = mentionItems[mentionHighlight];
        if (it) applyMentionInsert(it.insert);
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        setMentionOpen(false);
        setMentionRange(null);
      }
    },
    [mentionOpen, mentionItems, mentionHighlight, applyMentionInsert]
  );

  const rootStyles =
    bordered !== false
      ? { root: { borderTop: '1px solid #edebe9', paddingTop: 12 } }
      : undefined;

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={rootStyles}>
      <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
        Valor calculado (setComputed)
      </Text>
      <Stack
        tokens={{ childrenGap: 10 }}
        styles={{
          root: {
            border: '1px solid #edebe9',
            borderRadius: 8,
            padding: 12,
            background: '#faf9f8',
          },
        }}
      >
        <FormManagerCollapseSection
          title="Formas de expressão"
          isOpen={formsExprOpen}
          onToggle={() => setFormsExprOpen((v) => !v)}
        >
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            <strong>Número:</strong> só operadores <code style={{ fontSize: 12 }}>+ − * / ( )</code> e referências{' '}
            <code style={{ fontSize: 12 }}>{'{{NomeInternoDoCampo}}'}</code> (substituídas por valores numéricos dos
            campos).
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            <strong>Texto:</strong> prefixo <code style={{ fontSize: 12 }}>str:</code>, depois texto com{' '}
            <code style={{ fontSize: 12 }}>{'{{campo}}'}</code> e tokens entre parêntesis retos abaixo.
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            <strong>Diferença em dias entre duas datas:</strong>{' '}
            <code style={{ fontSize: 12 }}>{'{{DAYS:CampoDataA:CampoDataB}}'}</code> dentro da parte numérica.
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            <strong>Só um token:</strong> pode usar apenas <code style={{ fontSize: 12 }}>[myEmail]</code> (sem{' '}
            <code style={{ fontSize: 12 }}>str:</code>) se a expressão for só o token.
          </Text>
          <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130', marginTop: 4 } }}>
            Tokens de contexto (copiar como está; maiúsculas/minúsculas aceites onde aplicável)
          </Text>
          <Stack tokens={{ childrenGap: 6 }}>
            {SET_COMPUTED_CONTEXT_TOKENS.map((row) => (
              <Stack key={row.literal} horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
                <code style={{ fontSize: 12, flexShrink: 0 }}>{row.literal}</code>
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  {row.hint}
                </Text>
              </Stack>
            ))}
          </Stack>
        </FormManagerCollapseSection>
        {attachmentLibraryFolderOptions.length > 0 ? (
          <>
            <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130', marginTop: 8 } }}>
              Pastas na biblioteca de anexos
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Gera o URL da pasta ligado ao item quando o formulário corre com anexos configurados. Expressão gravada:{' '}
              <code style={{ fontSize: 12 }}>attfolder:idDoNó</code>.
            </Text>
            <Stack tokens={{ childrenGap: 8 }}>
              {attachmentLibraryFolderOptions.map((opt) => {
                const key = String(opt.key);
                const formula = `attfolder:${key}`;
                const active = ed.computedAttachmentFolderNodeId === key;
                return (
                  <Stack
                    key={key}
                    horizontal
                    verticalAlign="center"
                    horizontalAlign="space-between"
                    tokens={{ childrenGap: 8 }}
                    wrap
                    styles={{
                      root: {
                        border: active ? '1px solid #0078d4' : '1px solid #edebe9',
                        borderRadius: 4,
                        padding: '8px 10px',
                        background: active ? '#f3f9ff' : '#ffffff',
                      },
                    }}
                  >
                    <Stack tokens={{ childrenGap: 4 }} styles={{ root: { minWidth: 0, flex: 1 } }}>
                      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                        {opt.text}
                      </Text>
                      <code style={{ fontSize: 12, wordBreak: 'break-all' }}>{formula}</code>
                    </Stack>
                    <DefaultButton
                      text={active ? 'Pasta selecionada' : 'Usar esta pasta'}
                      disabled={active}
                      onClick={() =>
                        setEd((p) => ({
                          ...p,
                          computedAttachmentFolderNodeId: key,
                          computedExpression: '',
                        }))
                      }
                    />
                  </Stack>
                );
              })}
            </Stack>
          </>
        ) : null}
      </Stack>
      {ed.computedAttachmentFolderNodeId ? (
        <DefaultButton
          text="Editar expressão manual (sem pasta)"
          onClick={() =>
            setEd((p) => ({
              ...p,
              computedAttachmentFolderNodeId: '',
            }))
          }
        />
      ) : null}
      <div ref={wrapRef} style={{ position: 'relative', width: '100%' }}>
        <TextField
          label={
            ed.computedAttachmentFolderNodeId
              ? 'Expressão (limpe a pasta selecionada acima para editar)'
              : 'Expressão'
          }
          description={
            ed.computedAttachmentFolderNodeId
              ? undefined
              : 'Digite @ para sugestões (tokens, campos numéricos, pastas de anexos).'
          }
          multiline
          rows={3}
          value={ed.computedAttachmentFolderNodeId ? '' : ed.computedExpression}
          disabled={!!ed.computedAttachmentFolderNodeId}
          componentRef={tfRef}
          onChange={handleExprChange}
          onKeyDown={handleExprKeyDown}
        />
        {mentionOpen &&
        mentionItems.length > 0 &&
        !ed.computedAttachmentFolderNodeId &&
        mentionListPos
          ? createPortal(
              <div
                ref={mentionPortalRef}
                {...{ [FORM_FIELD_RULES_MENTION_PORTAL_ATTR]: '' }}
                role="listbox"
                aria-label="Sugestões de expressão"
                style={{
                  position: 'fixed',
                  left: mentionListPos.left,
                  top: mentionListPos.top,
                  width: mentionListPos.width,
                  maxWidth:
                    typeof window !== 'undefined'
                      ? Math.max(0, window.innerWidth - mentionListPos.left - 8)
                      : mentionListPos.width,
                  zIndex: 10000000,
                  minWidth: 280,
                  maxHeight: 280,
                  overflowY: 'auto',
                  border: '1px solid #edebe9',
                  borderRadius: 4,
                  boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
                  background: '#ffffff',
                  boxSizing: 'border-box',
                }}
                onMouseDown={(e) => e.preventDefault()}
              >
                {mentionItems.map((it, idx) => (
                  <div
                    key={it.key}
                    role="option"
                    aria-selected={idx === mentionHighlight}
                    style={{
                      padding: '8px 10px',
                      cursor: 'pointer',
                      background: idx === mentionHighlight ? '#edebe9' : 'transparent',
                      borderBottom:
                        idx < mentionItems.length - 1 ? '1px solid #f3f2f1' : undefined,
                    }}
                    onMouseEnter={() => setMentionHighlight(idx)}
                    onMouseDown={(e) => {
                      e.preventDefault();
                      applyMentionInsert(it.insert);
                    }}
                  >
                    <Text variant="small" styles={{ root: { fontWeight: 600, display: 'block' } }}>
                      {it.primary}
                    </Text>
                    <Text variant="small" styles={{ root: { color: '#605e5c', fontSize: 11 } }}>
                      {it.secondary}
                    </Text>
                  </div>
                ))}
              </div>,
              document.body
            )
          : null}
      </div>
      <Checkbox
        label="Sempre expressão ao vivo (ignora valor gravado em edição e visualização)"
        checked={ed.computedLiveInEditView}
        onChange={(_, c) => setEd((p) => ({ ...p, computedLiveInEditView: !!c }))}
      />
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Por omissão mantém-se o valor gravado ao abrir o item; se alterar a expressão aqui, o resultado calculado
        volta a aparecer até gravar outra vez.
      </Text>
    </Stack>
  );
}

function FieldRulesDisableEnableCollapseContent(props: {
  ed: IFieldRuleEditorState;
  setEd: React.Dispatch<React.SetStateAction<IFieldRuleEditorState>>;
  fieldOptions: IDropdownOption[];
}): JSX.Element {
  const { ed, setEd, fieldOptions } = props;
  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Condição no mesmo estilo das regras condicionais. Se ambas forem verdadeiras, «Tornar editável quando»
        prevalece sobre «Desativar quando».
      </Text>
      <Checkbox
        label="Desativar este campo quando a condição for verdadeira"
        checked={ed.disableWhenActive}
        onChange={(_, c) => setEd((p) => ({ ...p, disableWhenActive: !!c }))}
      />
      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
        <Dropdown
          label="Campo"
          options={fieldOptions}
          selectedKey={ed.disableWhenUi.field}
          disabled={!ed.disableWhenActive}
          onChange={(_, o) =>
            o &&
            setEd((p) => ({
              ...p,
              disableWhenUi: { ...p.disableWhenUi, field: String(o.key) },
            }))
          }
          styles={{ dropdown: { width: 160 } }}
        />
        <Dropdown
          label="Operador"
          options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
          selectedKey={ed.disableWhenUi.op}
          disabled={!ed.disableWhenActive}
          onChange={(_, o) =>
            o &&
            setEd((p) => ({
              ...p,
              disableWhenUi: { ...p.disableWhenUi, op: o.key as TFormConditionOp },
            }))
          }
          styles={{ dropdown: { width: 150 } }}
        />
        <Dropdown
          label="Comparar"
          options={[
            { key: 'literal', text: 'Texto fixo' },
            { key: 'field', text: 'Campo' },
            { key: 'token', text: 'Token' },
          ]}
          selectedKey={ed.disableWhenUi.compareKind}
          disabled={!ed.disableWhenActive}
          onChange={(_, o) =>
            o &&
            setEd((p) => ({
              ...p,
              disableWhenUi: { ...p.disableWhenUi, compareKind: o.key as IWhenUi['compareKind'] },
            }))
          }
          styles={{ dropdown: { width: 112 } }}
        />
        <TextField
          label="Valor"
          value={ed.disableWhenUi.compareValue}
          disabled={
            !ed.disableWhenActive ||
            ed.disableWhenUi.op === 'isEmpty' ||
            ed.disableWhenUi.op === 'isFilled' ||
            ed.disableWhenUi.op === 'isTrue' ||
            ed.disableWhenUi.op === 'isFalse'
          }
          onChange={(_, v) =>
            setEd((p) => ({
              ...p,
              disableWhenUi: { ...p.disableWhenUi, compareValue: v ?? '' },
            }))
          }
          styles={{ fieldGroup: { minWidth: 120 } }}
        />
      </Stack>
      <Checkbox
        label="Tornar editável quando a condição for verdadeira (sobrepor desativação acima)"
        checked={ed.enableWhenActive}
        onChange={(_, c) => setEd((p) => ({ ...p, enableWhenActive: !!c }))}
      />
      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
        <Dropdown
          label="Campo"
          options={fieldOptions}
          selectedKey={ed.enableWhenUi.field}
          disabled={!ed.enableWhenActive}
          onChange={(_, o) =>
            o &&
            setEd((p) => ({
              ...p,
              enableWhenUi: { ...p.enableWhenUi, field: String(o.key) },
            }))
          }
          styles={{ dropdown: { width: 160 } }}
        />
        <Dropdown
          label="Operador"
          options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
          selectedKey={ed.enableWhenUi.op}
          disabled={!ed.enableWhenActive}
          onChange={(_, o) =>
            o &&
            setEd((p) => ({
              ...p,
              enableWhenUi: { ...p.enableWhenUi, op: o.key as TFormConditionOp },
            }))
          }
          styles={{ dropdown: { width: 150 } }}
        />
        <Dropdown
          label="Comparar"
          options={[
            { key: 'literal', text: 'Texto fixo' },
            { key: 'field', text: 'Campo' },
            { key: 'token', text: 'Token' },
          ]}
          selectedKey={ed.enableWhenUi.compareKind}
          disabled={!ed.enableWhenActive}
          onChange={(_, o) =>
            o &&
            setEd((p) => ({
              ...p,
              enableWhenUi: { ...p.enableWhenUi, compareKind: o.key as IWhenUi['compareKind'] },
            }))
          }
          styles={{ dropdown: { width: 112 } }}
        />
        <TextField
          label="Valor"
          value={ed.enableWhenUi.compareValue}
          disabled={
            !ed.enableWhenActive ||
            ed.enableWhenUi.op === 'isEmpty' ||
            ed.enableWhenUi.op === 'isFilled' ||
            ed.enableWhenUi.op === 'isTrue' ||
            ed.enableWhenUi.op === 'isFalse'
          }
          onChange={(_, v) =>
            setEd((p) => ({
              ...p,
              enableWhenUi: { ...p.enableWhenUi, compareValue: v ?? '' },
            }))
          }
          styles={{ fieldGroup: { minWidth: 120 } }}
        />
      </Stack>
    </Stack>
  );
}

export const FormFieldRulesPanel: React.FC<IFormFieldRulesPanelProps> = ({
  isOpen,
  internalName,
  fieldConfig,
  meta,
  rules,
  fieldOptions,
  attachmentLibraryFolderOptions = [],
  lookupFieldsWebServerRelativeUrl,
  onDismiss,
  onApply,
}) => {
  const [fc, setFc] = useState<IFormFieldConfig>(fieldConfig);
  const [ed, setEd] = useState<IFieldRuleEditorState>(() => emptyFieldRuleEditorState());
  const [textRulesOpen, setTextRulesOpen] = useState<Record<string, boolean>>({});
  const [lookupRulesOpen, setLookupRulesOpen] = useState<Record<string, boolean>>({});
  const groupsService = useMemo(() => new GroupsService(), []);
  const [siteGroups, setSiteGroups] = useState<IGroupDetails[]>([]);
  const [siteGroupsLoading, setSiteGroupsLoading] = useState(false);
  const [siteGroupsErr, setSiteGroupsErr] = useState<string>();
  const [spGroupRuleNameFilter, setSpGroupRuleNameFilter] = useState('');

  const loadSiteGroups = useCallback((): void => {
    setSiteGroupsErr(undefined);
    setSiteGroupsLoading(true);
    groupsService
      .getSiteGroups()
      .then((g) => {
        setSiteGroups(g);
        setSiteGroupsLoading(false);
      })
      .catch((e) => {
        setSiteGroups([]);
        setSiteGroupsLoading(false);
        setSiteGroupsErr(e instanceof Error ? e.message : String(e));
      });
  }, [groupsService]);

  useEffect(() => {
    if (!isOpen) return;
    loadSiteGroups();
  }, [isOpen, loadSiteGroups]);

  const siteGroupsSorted = useMemo(() => {
    const g = siteGroups.slice();
    g.sort((a, b) => (a.Title < b.Title ? -1 : a.Title > b.Title ? 1 : 0));
    return g;
  }, [siteGroups]);

  const siteGroupsSortedForRules = useMemo(
    () => filterSiteGroupsByNameQuery(siteGroupsSorted, spGroupRuleNameFilter),
    [siteGroupsSorted, spGroupRuleNameFilter]
  );

  useEffect(() => {
    if (!isOpen) return;
    setFc({ ...fieldConfig });
    const st = fieldRuleStateFromRules(internalName, rules);
    const df = String(fieldOptions[0]?.key ?? 'Title');
    if (!st.disableWhenActive && !st.enableWhenActive) {
      st.disableWhenUi = { ...st.disableWhenUi, field: df };
      st.enableWhenUi = { ...st.enableWhenUi, field: df };
    }
    setEd(st);
  }, [isOpen, internalName, fieldConfig, rules, fieldOptions]);

  const mt = meta?.MappedType ?? 'unknown';
  const isTextRulesLikeText = mt === 'text' || mt === 'multiline';
  const fieldsServiceLookup = useMemo(() => new FieldsService(), []);
  const [lookupDestFields, setLookupDestFields] = useState<IFieldMetadata[]>([]);
  const [lookupDestErr, setLookupDestErr] = useState<string>();
  const [lookupDestLoading, setLookupDestLoading] = useState(false);

  useEffect(() => {
    if (!isOpen || (mt !== 'lookup' && mt !== 'lookupmulti') || !meta?.LookupList) {
      setLookupDestFields([]);
      setLookupDestErr(undefined);
      setLookupDestLoading(false);
      return;
    }
    let cancel = false;
    setLookupDestLoading(true);
    setLookupDestErr(undefined);
    const lw = lookupFieldsWebServerRelativeUrl?.trim() || undefined;
    fieldsServiceLookup
      .getFields(meta.LookupList, lw)
      .then((fields) => {
        if (!cancel) setLookupDestFields(fields);
      })
      .catch((e) => {
        if (!cancel) {
          setLookupDestFields([]);
          setLookupDestErr(e instanceof Error ? e.message : String(e));
        }
      })
      .finally(() => {
        if (!cancel) setLookupDestLoading(false);
      });
    return (): void => {
      cancel = true;
    };
  }, [isOpen, mt, meta?.LookupList, lookupFieldsWebServerRelativeUrl, fieldsServiceLookup]);

  const title = meta?.Title ?? internalName;

  const lookupRulesEligibleFlat = useMemo(
    (): IFieldMetadata[] => lookupRulesEligibleTargetFields(lookupDestFields),
    [lookupDestFields]
  );

  const lookupLabelFieldOptions = useMemo((): IDropdownOption[] => {
    const head: IDropdownOption[] = [{ key: '__default', text: '(Padrão da coluna no SharePoint)' }];
    const list = lookupRulesEligibleFlat.slice();
    list.sort((a, b) =>
      `${a.Title} (${a.InternalName})`.localeCompare(`${b.Title} (${b.InternalName})`, undefined, {
        sensitivity: 'base',
      })
    );
    return head.concat(list.map((f) => ({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` })));
  }, [lookupRulesEligibleFlat]);

  const toggleModeRow = useCallback((m: TFormManagerFormMode, checked: boolean) => {
    setEd((prev) => {
      let next = prev.modes.length === 0 ? ALL_MODES.slice() : prev.modes.slice();
      if (checked) {
        if (next.indexOf(m) === -1) next.push(m);
      } else {
        next = next.filter((x) => x !== m);
      }
      if (next.length === ALL_MODES.length) return { ...prev, modes: [] };
      return { ...prev, modes: next };
    });
  }, []);

  const modeRowChecked = useCallback((m: TFormManagerFormMode): boolean => {
    return ed.modes.length === 0 || ed.modes.indexOf(m) !== -1;
  }, [ed.modes]);

  const refFieldOptions = useMemo(
    () => fieldOptions.filter((o) => String(o.key) !== internalName),
    [fieldOptions, internalName]
  );
  const defaultRefField = refFieldOptions[0] ? String(refFieldOptions[0].key) : '';

  const patchGroupModes = useCallback((groupId: string, m: TFormManagerFormMode, checked: boolean): void => {
    setFc((p) => {
      const groups = p.textConditionalVisibility?.groups ?? [];
      const nextGroups = groups.map((g) => {
        if (g.id !== groupId) return g;
        let next = g.modes.length === 0 ? ALL_MODES.slice() : g.modes.slice();
        if (checked) {
          if (next.indexOf(m) === -1) next.push(m);
        } else {
          next = next.filter((x) => x !== m);
        }
        if (next.length === ALL_MODES.length) next = [];
        return { ...g, modes: next };
      });
      return { ...p, textConditionalVisibility: { groups: nextGroups } };
    });
  }, []);

  const groupModeChecked = useCallback((modes: TFormManagerFormMode[], m: TFormManagerFormMode): boolean => {
    return modes.length === 0 || modes.indexOf(m) !== -1;
  }, []);

  const toggleTextRulesSection = useCallback((id: string): void => {
    setTextRulesOpen((prev) => ({ ...prev, [id]: !prev[id] }));
  }, []);
  const isTextRulesOpen = useCallback((id: string): boolean => textRulesOpen[id] === true, [textRulesOpen]);

  const toggleLookupRulesSection = useCallback((id: string): void => {
    setLookupRulesOpen((prev) => ({ ...prev, [id]: !prev[id] }));
  }, []);
  const isLookupRulesOpen = useCallback((id: string): boolean => lookupRulesOpen[id] === true, [lookupRulesOpen]);

  const handleApply = (): void => {
    onApply(fc, ed);
    onDismiss();
  };

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.medium}
      headerText={`Configurar regras — ${title}`}
      onDismiss={onDismiss}
      onOuterClick={(ev) => {
        const t = ev?.target;
        if (t instanceof Element && t.closest(`[${FORM_FIELD_RULES_MENTION_PORTAL_ATTR}]`)) return;
        onDismiss();
      }}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        {!isTextRulesLikeText && (
          <>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {internalName} · {mt}
              {fc.sectionId ? ` · etapa ${fc.sectionId}` : ''}
            </Text>
            <Text variant="small">Aplicar regras geradas apenas nos modos:</Text>
            <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
              {MODE_OPTS.map((m) => (
                <Checkbox
                  key={m.key}
                  label={m.label}
                  checked={modeRowChecked(m.key)}
                  onChange={(_, c) => toggleModeRow(m.key, !!c)}
                />
              ))}
            </Stack>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Vazio = todos os modos. Desmarque um para restringir.
            </Text>
            {mt !== 'lookup' && mt !== 'lookupmulti' ? (
              <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
                <DefaultButton
                  text="Modelo: data não no passado"
                  onClick={() => setEd((prev) => mergeFieldRuleEditorState(prev, templateFieldRulesDateNotPast()))}
                />
                <DefaultButton
                  text="Modelo: validar e-mail"
                  onClick={() => setEd((prev) => mergeFieldRuleEditorState(prev, templateFieldRulesEmail()))}
                />
              </Stack>
            ) : null}
            {mt === 'url' && (
              <TextField
                label="Placeholder"
                value={fc.placeholder ?? ''}
                onChange={(_, v) => setFc((p) => ({ ...p, placeholder: v || undefined }))}
              />
            )}
            <TextField
              label="Texto de ajuda (campo)"
              multiline
              rows={2}
              value={fc.helpText ?? ''}
              onChange={(_, v) => setFc((p) => ({ ...p, helpText: v || undefined }))}
            />
            {(mt !== 'lookup' && mt !== 'lookupmulti') ? (
              <>
                <TextField
                  label="Valor padrão (token ou texto; aplica se vazio)"
                  value={ed.defaultValue}
                  onChange={(_, v) => setEd((p) => ({ ...p, defaultValue: v ?? '' }))}
                />
                {internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL && !isFormBannerFieldConfig(fieldConfig) ? (
                  <SetComputedRulesBlock
                    ed={ed}
                    setEd={setEd}
                    fieldOptions={fieldOptions}
                    attachmentLibraryFolderOptions={attachmentLibraryFolderOptions}
                  />
                ) : null}
              </>
            ) : null}
            <FormManagerCollapseSection
              title="Desativar / ativar o campo"
              isOpen={isTextRulesOpen(FIELD_RULES_DISABLE_ENABLE_SECTION_ID)}
              onToggle={() => toggleTextRulesSection(FIELD_RULES_DISABLE_ENABLE_SECTION_ID)}
            >
              <FieldRulesDisableEnableCollapseContent ed={ed} setEd={setEd} fieldOptions={fieldOptions} />
            </FormManagerCollapseSection>
          </>
        )}
        {isTextRulesLikeText && (
          <Stack tokens={{ childrenGap: 10 }}>
            <FormManagerCollapseSection
              title="Exibição"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.display)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.display)}
            >
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {internalName} · {mt}
                {fc.sectionId ? ` · etapa ${fc.sectionId}` : ''}
              </Text>
              <Text variant="small">Aplicar regras geradas apenas nos modos:</Text>
              <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                {MODE_OPTS.map((m) => (
                  <Checkbox
                    key={m.key}
                    label={m.label}
                    checked={modeRowChecked(m.key)}
                    onChange={(_, c) => toggleModeRow(m.key, !!c)}
                  />
                ))}
              </Stack>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Vazio = todos os modos. Desmarque um para restringir.
              </Text>
              <TextField
                label="Placeholder"
                value={fc.placeholder ?? ''}
                onChange={(_, v) => setFc((p) => ({ ...p, placeholder: v || undefined }))}
              />
              {mt === 'multiline' ? (
                <TextField
                  label="Linhas do textarea (altura inicial)"
                  type="number"
                  min={1}
                  max={50}
                  value={fc.textareaRows !== undefined ? String(fc.textareaRows) : ''}
                  onChange={(_, v) =>
                    setFc((p) => {
                      const next: IFormFieldConfig = { ...p };
                      const t = (v ?? '').trim();
                      if (!t) {
                        delete next.textareaRows;
                        return next;
                      }
                      const n = Number(t);
                      if (!isFinite(n)) return next;
                      next.textareaRows = Math.min(50, Math.max(1, Math.floor(n)));
                      return next;
                    })
                  }
                />
              ) : null}
              <TextField
                label="Texto de ajuda (campo)"
                multiline
                rows={2}
                value={fc.helpText ?? ''}
                onChange={(_, v) => setFc((p) => ({ ...p, helpText: v || undefined }))}
              />
              <TextField
                label="Valor padrão (token ou texto; aplica se vazio)"
                value={ed.defaultValue}
                onChange={(_, v) => setEd((p) => ({ ...p, defaultValue: v ?? '' }))}
              />
              {internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL && !isFormBannerFieldConfig(fieldConfig) && (
                <SetComputedRulesBlock
                  ed={ed}
                  setEd={setEd}
                  fieldOptions={fieldOptions}
                  attachmentLibraryFolderOptions={attachmentLibraryFolderOptions}
                  bordered={false}
                />
              )}
              <Checkbox
                label="Somente leitura"
                checked={fc.readOnly === true}
                onChange={(_, c) =>
                  setFc((p) => ({
                    ...p,
                    ...(c ? { readOnly: true } : { readOnly: undefined }),
                  }))
                }
              />
              <Checkbox
                label="Ocultar no formulário"
                checked={fc.visible === false}
                onChange={(_, c) =>
                  setFc((p) => {
                    const next: IFormFieldConfig = { ...p };
                    if (c) next.visible = false;
                    else delete next.visible;
                    return next;
                  })
                }
              />
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Validação"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.validation)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.validation)}
            >
              <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
                <DefaultButton
                  text="Modelo: data não no passado"
                  onClick={() => setEd((prev) => mergeFieldRuleEditorState(prev, templateFieldRulesDateNotPast()))}
                />
                <DefaultButton
                  text="Modelo: validar e-mail"
                  onClick={() => setEd((prev) => mergeFieldRuleEditorState(prev, templateFieldRulesEmail()))}
                />
              </Stack>
              <Checkbox
                label="Obrigatório"
                checked={fc.required === true}
                onChange={(_, c) =>
                  setFc((p) => ({
                    ...p,
                    ...(c ? { required: true } : { required: undefined }),
                  }))
                }
              />
              <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
                <TextField
                  label="Mín. caracteres"
                  value={ed.validateValue.minLength}
                  onChange={(_, v) =>
                    setEd((p) => ({ ...p, validateValue: { ...p.validateValue, minLength: v ?? '' } }))
                  }
                />
                <TextField
                  label="Máx. caracteres"
                  value={ed.validateValue.maxLength}
                  onChange={(_, v) =>
                    setEd((p) => ({ ...p, validateValue: { ...p.validateValue, maxLength: v ?? '' } }))
                  }
                />
              </Stack>
              <TextField
                label="Regex (padrão)"
                value={ed.validateValue.pattern}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, pattern: v ?? '' } }))
                }
              />
              <TextField
                label="Mensagem se falhar o padrão"
                value={ed.validateValue.patternMessage}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, patternMessage: v ?? '' } }))
                }
              />
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Pré-visualização: {buildFieldUiRules(internalName, ed, fc).length} regra(s) gerada(s) para este campo.
              </Text>
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Transformação"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.transform)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.transform)}
            >
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Aplica a colunas de texto e múltiplas linhas, para valor digitado, predefinido por regras e
                resultados de expressão (p. ex. regra «Valor calculado»).
              </Text>
              <Stack tokens={{ childrenGap: 4 }}>
                <Checkbox
                  label="Maiúsculas"
                  checked={fc.textValueTransform === 'uppercase'}
                  onChange={(_, c) =>
                    setFc((p) => {
                      const next: IFormFieldConfig = { ...p };
                      if (c) next.textValueTransform = 'uppercase';
                      else if (p.textValueTransform === 'uppercase') delete next.textValueTransform;
                      return next;
                    })
                  }
                />
                <Checkbox
                  label="Minúsculas"
                  checked={fc.textValueTransform === 'lowercase'}
                  onChange={(_, c) =>
                    setFc((p) => {
                      const next: IFormFieldConfig = { ...p };
                      if (c) next.textValueTransform = 'lowercase';
                      else if (p.textValueTransform === 'lowercase') delete next.textValueTransform;
                      return next;
                    })
                  }
                />
                <Checkbox
                  label="Capitalizar"
                  checked={fc.textValueTransform === 'capitalize'}
                  onChange={(_, c) =>
                    setFc((p) => {
                      const next: IFormFieldConfig = { ...p };
                      if (c) next.textValueTransform = 'capitalize';
                      else if (p.textValueTransform === 'capitalize') delete next.textValueTransform;
                      return next;
                    })
                  }
                />
              </Stack>
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Máscaras"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.masks)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.masks)}
            >
              <Text variant="small" styles={{ root: { color: '#605e5c', marginBottom: 8 } }}>
                Guia:{' '}
                <Link href="https://imask.js.org/guide" target="_blank" rel="noopener noreferrer">
                  imask.js.org/guide
                </Link>
                .
              </Text>
              <ChoiceGroup
                selectedKey={fc.textInputMaskKind ?? 'none'}
                options={TEXT_MASK_CHOICE_OPTIONS}
                onChange={(_, o) => {
                  if (!o) return;
                  const k = String(o.key);
                  setFc((p) => {
                    const next: IFormFieldConfig = { ...p };
                    if (k === 'none') delete next.textInputMaskKind;
                    else next.textInputMaskKind = k as TFormFieldTextInputMaskKind;
                    return next;
                  });
                }}
              />
              {fc.textInputMaskKind === 'custom' ? (
                <TextField
                  label="Padrão IMask"
                  multiline
                  rows={3}
                  placeholder="Ex.: 00/00/0000 ou AA-0000"
                  value={fc.textInputMaskCustomPattern ?? ''}
                  onChange={(_, v) =>
                    setFc((p) => {
                      const raw = v ?? '';
                      const cut = raw.slice(0, TEXT_INPUT_MASK_CUSTOM_MAX_LEN);
                      const next: IFormFieldConfig = { ...p };
                      if (cut) next.textInputMaskCustomPattern = cut;
                      else delete next.textInputMaskCustomPattern;
                      return next;
                    })
                  }
                />
              ) : null}
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Desativar / ativar o campo"
              isOpen={isTextRulesOpen(FIELD_RULES_DISABLE_ENABLE_SECTION_ID)}
              onToggle={() => toggleTextRulesSection(FIELD_RULES_DISABLE_ENABLE_SECTION_ID)}
            >
              <FieldRulesDisableEnableCollapseContent ed={ed} setEd={setEd} fieldOptions={fieldOptions} />
            </FormManagerCollapseSection>
            <FormManagerCollapseSection
              title="Condicionais"
              isOpen={isTextRulesOpen(TEXT_RULES_COLLAPSE_IDS.conditionals)}
              onToggle={() => toggleTextRulesSection(TEXT_RULES_COLLAPSE_IDS.conditionals)}
            >
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {internalName} · {mt}
                {fc.sectionId ? ` · etapa ${fc.sectionId}` : ''}
              </Text>
              <Text variant="small">
                Visibilidade dinâmica deste campo: pode haver vários grupos; cada grupo é avaliado de forma
                independente. Se mais do que um se aplicar e as ações divergirem, prevalece ocultar.
              </Text>
              <PrimaryButton
                text="Adicionar grupo de regra"
                disabled={!refFieldOptions.length}
                onClick={() =>
                  setFc((p) => ({
                    ...p,
                    textConditionalVisibility: {
                      groups: [
                        ...(p.textConditionalVisibility?.groups ?? []),
                        newTextConditionalGroup(defaultRefField),
                      ],
                    },
                  }))
                }
              />
              {!refFieldOptions.length ? (
                <Text variant="small" styles={{ root: { color: '#a4262c' } }}>
                  Não há outros campos no formulário para referenciar nas condições.
                </Text>
              ) : null}
              {(fc.textConditionalVisibility?.groups ?? []).map((g, gi) => (
                <Stack
                  key={g.id}
                  tokens={{ childrenGap: 10 }}
                  styles={{
                    root: {
                      border: '1px solid #edebe9',
                      borderRadius: 4,
                      padding: 12,
                      marginTop: 10,
                      background: '#faf9f8',
                    },
                  }}
                >
                  <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                    <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                      Grupo {gi + 1}
                    </Text>
                    <DefaultButton
                      text="Remover grupo"
                      onClick={() =>
                        setFc((p) => {
                          const nextList = (p.textConditionalVisibility?.groups ?? []).filter((x) => x.id !== g.id);
                          const next: IFormFieldConfig = { ...p };
                          if (!nextList.length) delete next.textConditionalVisibility;
                          else next.textConditionalVisibility = { groups: nextList };
                          return next;
                        })
                      }
                    />
                  </Stack>
                  <Text variant="small">Aplicar esta regra apenas nos modos:</Text>
                  <Stack horizontal tokens={{ childrenGap: 16 }} wrap>
                    {MODE_OPTS.map((m) => (
                      <Checkbox
                        key={`${g.id}-${m.key}`}
                        label={m.label}
                        checked={groupModeChecked(g.modes, m.key)}
                        onChange={(_, c) => patchGroupModes(g.id, m.key, !!c)}
                      />
                    ))}
                  </Stack>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Vazio = todos os modos. Desmarque um para restringir.
                  </Text>
                  <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                    Grupos do SharePoint
                  </Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Vazio = a regra aplica-se a todos. Se marcar grupos, só quem pertencer a pelo menos um deles (e
                    cumprir as condições acima) fica abrangido.
                  </Text>
                  <TextField
                    placeholder="Filtrar grupos por nome"
                    value={spGroupRuleNameFilter}
                    onChange={(_: unknown, v?: string) => setSpGroupRuleNameFilter(v ?? '')}
                    styles={{ root: { maxWidth: 420 } }}
                  />
                  {siteGroupsLoading && <Spinner label="A carregar grupos do site…" />}
                  {siteGroupsErr ? (
                    <>
                      <MessageBar messageBarType={MessageBarType.warning}>{siteGroupsErr}</MessageBar>
                      <DefaultButton text="Tentar carregar grupos novamente" onClick={() => loadSiteGroups()} />
                    </>
                  ) : null}
                  {!siteGroupsLoading ? (
                    <Stack
                      tokens={{ childrenGap: 6 }}
                      styles={{
                        root: {
                          maxHeight: 240,
                          overflowY: 'auto',
                          border: '1px solid #edebe9',
                          borderRadius: 4,
                          padding: 8,
                        },
                      }}
                    >
                      {(g.groupTitles ?? [])
                        .filter(
                          (t) =>
                            !siteGroups.some(
                              (sg) => normSpGroupTitle(sg.Title) === normSpGroupTitle(t)
                            )
                        )
                        .filter((t) => {
                          const q = spGroupRuleNameFilter.trim().toLowerCase();
                          return !q || t.toLowerCase().includes(q);
                        })
                        .map((t, oi) => (
                          <Checkbox
                            key={`tx-orphan-${g.id}-${oi}-${t}`}
                            label={`${t} (guardado; não na lista do site)`}
                            checked
                            onChange={(_, c) => {
                              if (c) return;
                              setFc((p) => ({
                                ...p,
                                textConditionalVisibility: {
                                  groups: (p.textConditionalVisibility?.groups ?? []).map((gr) => {
                                    if (gr.id !== g.id) return gr;
                                    const cur = gr.groupTitles ?? [];
                                    const n = normSpGroupTitle(t);
                                    const next = cur.filter((x) => normSpGroupTitle(x) !== n);
                                    const out: ITextFieldConditionalGroup = { ...gr };
                                    if (next.length) out.groupTitles = next;
                                    else delete out.groupTitles;
                                    return out;
                                  }),
                                },
                              }));
                            }}
                          />
                        ))}
                      {siteGroupsSortedForRules.map((sg) => {
                        const cur = g.groupTitles ?? [];
                        const n = normSpGroupTitle(sg.Title);
                        const checked = cur.some((x) => normSpGroupTitle(x) === n);
                        return (
                          <Checkbox
                            key={`tx-sg-${g.id}-${sg.Id}`}
                            label={sg.Title}
                            title={sg.Description || undefined}
                            checked={checked}
                            onChange={(_, c) => {
                              setFc((p) => ({
                                ...p,
                                textConditionalVisibility: {
                                  groups: (p.textConditionalVisibility?.groups ?? []).map((gr) => {
                                    if (gr.id !== g.id) return gr;
                                    const prevTitles = gr.groupTitles ?? [];
                                    let next: string[];
                                    if (c) {
                                      next = checked ? prevTitles : prevTitles.concat([sg.Title]);
                                    } else {
                                      next = prevTitles.filter((x) => normSpGroupTitle(x) !== n);
                                    }
                                    const out: ITextFieldConditionalGroup = { ...gr };
                                    if (next.length) out.groupTitles = next;
                                    else delete out.groupTitles;
                                    return out;
                                  }),
                                },
                              }));
                            }}
                          />
                        );
                      })}
                      {siteGroupsSorted.length > 0 &&
                      !siteGroupsSortedForRules.length &&
                      spGroupRuleNameFilter.trim() ? (
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                          Nenhum grupo corresponde ao filtro.
                        </Text>
                      ) : null}
                      {!siteGroupsSorted.length && !(g.groupTitles ?? []).length ? (
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                          Nenhum grupo no site.
                        </Text>
                      ) : null}
                    </Stack>
                  ) : null}
                  <ChoiceGroup
                    label="Operador lógico entre condições"
                    selectedKey={g.groupOp}
                    options={TEXT_COND_GROUP_OP_OPTS}
                    onChange={(_, opt) =>
                      opt &&
                      setFc((p) => ({
                        ...p,
                        textConditionalVisibility: {
                          groups: (p.textConditionalVisibility?.groups ?? []).map((gr) =>
                            gr.id === g.id ? { ...gr, groupOp: opt.key as TTextFieldConditionalGroupOp } : gr
                          ),
                        },
                      }))
                    }
                  />
                  <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                    Condições
                  </Text>
                  {g.conditions.map((c) => (
                    <Stack
                      key={c.id}
                      horizontal
                      wrap
                      tokens={{ childrenGap: 8 }}
                      verticalAlign="end"
                      styles={{ root: { alignItems: 'flex-end' } }}
                    >
                      <Dropdown
                        label="Campo"
                        options={refFieldOptions}
                        selectedKey={c.refField || undefined}
                        onChange={(_, o) =>
                          o &&
                          setFc((p) => ({
                            ...p,
                            textConditionalVisibility: {
                              groups: (p.textConditionalVisibility?.groups ?? []).map((gr) =>
                                gr.id !== g.id
                                  ? gr
                                  : {
                                      ...gr,
                                      conditions: gr.conditions.map((row) =>
                                        row.id === c.id ? { ...row, refField: String(o.key) } : row
                                      ),
                                    }
                              ),
                            },
                          }))
                        }
                        styles={{ dropdown: { width: 180 } }}
                      />
                      <Dropdown
                        label="Operador"
                        options={TEXT_DISPLAY_OP_OPTS.map((x) => ({ key: x.key, text: x.text }))}
                        selectedKey={c.op}
                        onChange={(_, o) =>
                          o &&
                          setFc((p) => ({
                            ...p,
                            textConditionalVisibility: {
                              groups: (p.textConditionalVisibility?.groups ?? []).map((gr) =>
                                gr.id !== g.id
                                  ? gr
                                  : {
                                      ...gr,
                                      conditions: gr.conditions.map((row) =>
                                        row.id === c.id
                                          ? { ...row, op: o.key as TTextFieldConditionalDisplayOp }
                                          : row
                                      ),
                                    }
                              ),
                            },
                          }))
                        }
                        styles={{ dropdown: { width: 160 } }}
                      />
                      <Dropdown
                        label="Comparar"
                        options={[
                          { key: 'literal', text: 'Texto fixo' },
                          { key: 'field', text: 'Campo' },
                          { key: 'token', text: 'Token' },
                        ]}
                        selectedKey={c.compareKind}
                        disabled={c.op === 'isEmpty' || c.op === 'isFilled'}
                        onChange={(_, o) =>
                          o &&
                          setFc((p) => ({
                            ...p,
                            textConditionalVisibility: {
                              groups: (p.textConditionalVisibility?.groups ?? []).map((gr) =>
                                gr.id !== g.id
                                  ? gr
                                  : {
                                      ...gr,
                                      conditions: gr.conditions.map((row) =>
                                        row.id === c.id ? { ...row, compareKind: o.key as TFormCompareKind } : row
                                      ),
                                    }
                              ),
                            },
                          }))
                        }
                        styles={{ dropdown: { width: 112 } }}
                      />
                      <TextField
                        label="Valor"
                        value={c.compareValue}
                        disabled={c.op === 'isEmpty' || c.op === 'isFilled'}
                        onChange={(_, v) =>
                          setFc((p) => ({
                            ...p,
                            textConditionalVisibility: {
                              groups: (p.textConditionalVisibility?.groups ?? []).map((gr) =>
                                gr.id !== g.id
                                  ? gr
                                  : {
                                      ...gr,
                                      conditions: gr.conditions.map((row) =>
                                        row.id === c.id ? { ...row, compareValue: v ?? '' } : row
                                      ),
                                    }
                              ),
                            },
                          }))
                        }
                        styles={{ fieldGroup: { minWidth: 140 } }}
                      />
                      <DefaultButton
                        text="Remover"
                        disabled={g.conditions.length < 2}
                        onClick={() =>
                          setFc((p) => ({
                            ...p,
                            textConditionalVisibility: {
                              groups: (p.textConditionalVisibility?.groups ?? []).map((gr) => {
                                if (gr.id !== g.id) return gr;
                                const filt = gr.conditions.filter((row) => row.id !== c.id);
                                return {
                                  ...gr,
                                  conditions: filt.length
                                    ? filt
                                    : [newTextConditionalCondition(defaultRefField)],
                                };
                              }),
                            },
                          }))
                        }
                      />
                    </Stack>
                  ))}
                  <DefaultButton
                    text="Adicionar condição"
                    onClick={() =>
                      setFc((p) => ({
                        ...p,
                        textConditionalVisibility: {
                          groups: (p.textConditionalVisibility?.groups ?? []).map((gr) =>
                            gr.id === g.id
                              ? {
                                  ...gr,
                                  conditions: [...gr.conditions, newTextConditionalCondition(defaultRefField)],
                                }
                              : gr
                          ),
                        },
                      }))
                    }
                  />
                  <Dropdown
                    label="Ação deste grupo"
                    options={[
                      { key: 'show', text: 'Mostrar' },
                      { key: 'hide', text: 'Ocultar' },
                      { key: 'disable', text: 'Desabilitar' },
                    ]}
                    selectedKey={g.action}
                    onChange={(_, o) =>
                      o &&
                      setFc((p) => ({
                        ...p,
                        textConditionalVisibility: {
                          groups: (p.textConditionalVisibility?.groups ?? []).map((gr) =>
                            gr.id === g.id ? { ...gr, action: o.key as TTextFieldConditionalAction } : gr
                          ),
                        },
                      }))
                    }
                    styles={{ dropdown: { maxWidth: 280 } }}
                  />
                </Stack>
              ))}
            </FormManagerCollapseSection>
          </Stack>
        )}
        {(mt === 'url' || mt === 'unknown') && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Validação de texto</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              <TextField
                label="Mín. caracteres"
                value={ed.validateValue.minLength}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, minLength: v ?? '' } }))
                }
              />
              <TextField
                label="Máx. caracteres"
                value={ed.validateValue.maxLength}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, maxLength: v ?? '' } }))
                }
              />
            </Stack>
            <TextField
              label="Regex (padrão)"
              value={ed.validateValue.pattern}
              onChange={(_, v) =>
                setEd((p) => ({ ...p, validateValue: { ...p.validateValue, pattern: v ?? '' } }))
              }
            />
            <TextField
              label="Mensagem se falhar o padrão"
              value={ed.validateValue.patternMessage}
              onChange={(_, v) =>
                setEd((p) => ({ ...p, validateValue: { ...p.validateValue, patternMessage: v ?? '' } }))
              }
            />
          </Stack>
        )}
        {(mt === 'number' || mt === 'currency') && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Validação numérica</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              <TextField
                label="Mínimo"
                type="number"
                value={ed.validateValue.minNumber}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, minNumber: v ?? '' } }))
                }
              />
              <TextField
                label="Máximo"
                type="number"
                value={ed.validateValue.maxNumber}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateValue: { ...p.validateValue, maxNumber: v ?? '' } }))
                }
              />
            </Stack>
          </Stack>
        )}
        {mt === 'datetime' && (
          <Stack tokens={{ childrenGap: 8 }}>
            <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>Validação de data</Text>
            <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
              <TextField
                label="Mín. dias a partir de hoje"
                value={ed.validateDate.minDaysFromToday}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateDate: { ...p.validateDate, minDaysFromToday: v ?? '' } }))
                }
              />
              <TextField
                label="Máx. dias a partir de hoje"
                value={ed.validateDate.maxDaysFromToday}
                onChange={(_, v) =>
                  setEd((p) => ({ ...p, validateDate: { ...p.validateDate, maxDaysFromToday: v ?? '' } }))
                }
              />
            </Stack>
            <Checkbox
              label="Bloquear fins de semana"
              checked={ed.validateDate.blockWeekends}
              onChange={(_, c) =>
                setEd((p) => ({ ...p, validateDate: { ...p.validateDate, blockWeekends: !!c } }))
              }
            />
            <Dropdown
              label="Data &gt;= campo"
              options={[{ key: '', text: '—' }, ...fieldOptions]}
              selectedKey={ed.validateDate.gteField || ''}
              onChange={(_, o) =>
                setEd((p) => ({
                  ...p,
                  validateDate: { ...p.validateDate, gteField: o ? String(o.key) : '' },
                }))
              }
            />
            <Dropdown
              label="Data &lt;= campo"
              options={[{ key: '', text: '—' }, ...fieldOptions]}
              selectedKey={ed.validateDate.lteField || ''}
              onChange={(_, o) =>
                setEd((p) => ({
                  ...p,
                  validateDate: { ...p.validateDate, lteField: o ? String(o.key) : '' },
                }))
              }
            />
            <TextField
              label="Mensagem de erro"
              value={ed.validateDate.message}
              onChange={(_, v) =>
                setEd((p) => ({ ...p, validateDate: { ...p.validateDate, message: v ?? '' } }))
              }
            />
          </Stack>
        )}
        {(mt === 'choice' || mt === 'multichoice') && (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Condições «se outro campo = X» entre colunas: JSON do gestor. Neste painel: obrigatoriedade e validação de
            texto, quando aplicável.
          </Text>
        )}
        {(mt === 'lookup' || mt === 'lookupmulti') && meta?.LookupList && (
          <FormManagerCollapseSection
            title="Lista ligada (texto das opções)"
            isOpen={isLookupRulesOpen('lookupLabel')}
            onToggle={() => toggleLookupRulesSection('lookupLabel')}
          >
            <Stack tokens={{ childrenGap: 10 }} styles={{ root: { maxWidth: 480 } }}>
              {lookupDestLoading && <Spinner />}
              {lookupDestErr && (
                <MessageBar messageBarType={MessageBarType.error}>{lookupDestErr}</MessageBar>
              )}
              <Dropdown
                label="Campo para o texto das opções"
                options={lookupLabelFieldOptions}
                selectedKey={fc.lookupOptionLabelField?.trim() ? fc.lookupOptionLabelField.trim() : '__default'}
                disabled={lookupDestLoading}
                onChange={(_, opt): void =>
                  setFc((p): IFormFieldConfig => {
                    const k = String(opt?.key ?? '');
                    const { lookupOptionLabelField: _l, lookupOptionLabelSubProp: _s, ...rest } = p;
                    if (!k || k === '__default') return rest;
                    return { ...rest, lookupOptionLabelField: k };
                  })
                }
              />
              {(() => {
                const selName = fc.lookupOptionLabelField?.trim() || meta.LookupField?.trim() || 'Title';
                const labelMeta = lookupRulesEligibleFlat.find((f) => f.InternalName === selName);
                const isUser = labelMeta?.MappedType === 'user' || labelMeta?.MappedType === 'usermulti';
                const isLookupSub = labelMeta?.MappedType === 'lookup' || labelMeta?.MappedType === 'lookupmulti';
                const isMulti = labelMeta?.MappedType === 'usermulti' || labelMeta?.MappedType === 'lookupmulti';
                if (!isUser && !isLookupSub) return null;
                const subPropOptions: IDropdownOption[] = isUser
                  ? [
                      { key: '', text: '(Padrão — Nome)' },
                      { key: 'Title', text: 'Nome (Title)' },
                      { key: 'EMail', text: 'E-mail (EMail)' },
                    ]
                  : [
                      { key: '', text: '(Padrão — Title)' },
                      { key: 'Title', text: 'Valor do lookup (Title)' },
                    ];
                return (
                  <Stack tokens={{ childrenGap: 6 }}>
                    <Dropdown
                      label="Propriedade a exibir"
                      options={subPropOptions}
                      selectedKey={fc.lookupOptionLabelSubProp ?? ''}
                      onChange={(_, o): void =>
                        setFc((p): IFormFieldConfig => {
                          const k = String(o?.key ?? '');
                          const { lookupOptionLabelSubProp: _omit, ...rest } = p;
                          if (!k) return rest;
                          return { ...rest, lookupOptionLabelSubProp: k };
                        })
                      }
                    />
                    {isMulti && (
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        Campo multi-valor — valores concatenados com "; " no texto da opção.
                      </Text>
                    )}
                  </Stack>
                );
              })()}
            </Stack>
          </FormManagerCollapseSection>
        )}
        {(mt === 'lookup' || mt === 'lookupmulti') && meta?.LookupList && (
          <FormManagerCollapseSection
            title="Detalhe abaixo da seleção"
            isOpen={isLookupRulesOpen('lookupDetailBelow')}
            onToggle={() => toggleLookupRulesSection('lookupDetailBelow')}
          >
            <Stack tokens={{ childrenGap: 8 }} styles={{ root: { maxWidth: 480 } }}>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Após escolher uma opção, mostra campos da lista ligada em só leitura por baixo do lookup.
              </Text>
              {lookupDestLoading && <Spinner />}
              {lookupDestErr && (
                <MessageBar messageBarType={MessageBarType.error}>{lookupDestErr}</MessageBar>
              )}
              <Stack tokens={{ childrenGap: 4 }}>
                {lookupRulesEligibleFlat.map((f) => (
                  <Checkbox
                    key={`det-${f.InternalName}`}
                    label={`${f.Title} (${f.InternalName})`}
                    checked={(fc.lookupOptionDetailBelowFields ?? []).indexOf(f.InternalName) !== -1}
                    disabled={lookupDestLoading}
                    onChange={(_, checked): void =>
                      setFc((p): IFormFieldConfig => {
                        const prev = p.lookupOptionDetailBelowFields ?? [];
                        let next = prev.slice();
                        const ix = next.indexOf(f.InternalName);
                        if (checked && ix === -1) next.push(f.InternalName);
                        if (!checked && ix !== -1) next.splice(ix, 1);
                        next.sort();
                        if (next.length === 0) {
                          const { lookupOptionDetailBelowFields: _omit, ...rest } = p;
                          return rest;
                        }
                        return { ...p, lookupOptionDetailBelowFields: next };
                      })
                    }
                  />
                ))}
              </Stack>
            </Stack>
          </FormManagerCollapseSection>
        )}
        {(mt === 'lookup' || mt === 'lookupmulti') && (
          <FormManagerCollapseSection
            title="Filtrar opções"
            isOpen={isLookupRulesOpen('lookupFilter')}
            onToggle={() => toggleLookupRulesSection('lookupFilter')}
          >
            <Stack tokens={{ childrenGap: 8 }}>
              <Dropdown
                label="Campo pai (filtro)"
                options={[{ key: '', text: '—' }, ...fieldOptions]}
                selectedKey={ed.filterLookup.parentField || ''}
                onChange={(_, o) =>
                  setEd((p) => ({
                    ...p,
                    filterLookup: { ...p.filterLookup, parentField: o ? String(o.key) : '', childField: '', filterOperator: '' },
                  }))
                }
              />
              <Dropdown
                label="Comparador"
                options={[
                  { key: '', text: '—' },
                  { key: 'eq', text: 'Igual a (eq)' },
                  { key: 'ne', text: 'Diferente de (ne)' },
                  { key: 'lt', text: 'Menor que (lt)' },
                  { key: 'le', text: 'Menor ou igual (le)' },
                  { key: 'gt', text: 'Maior que (gt)' },
                  { key: 'ge', text: 'Maior ou igual (ge)' },
                  { key: 'contains', text: 'Contém (substringof)' },
                  { key: 'startsWith', text: 'Começa com (startswith)' },
                ]}
                selectedKey={ed.filterLookup.filterOperator || ''}
                disabled={!ed.filterLookup.parentField}
                onChange={(_, o) =>
                  setEd((p) => ({
                    ...p,
                    filterLookup: { ...p.filterLookup, filterOperator: (o ? String(o.key) : '') as typeof p.filterLookup.filterOperator },
                  }))
                }
              />
              {meta?.LookupList && lookupRulesEligibleFlat.length > 0 ? (
                <Dropdown
                  label="Campo na lista filho"
                  options={[
                    { key: '', text: '—' },
                    ...lookupRulesEligibleFlat.map((f) => ({
                      key: f.InternalName,
                      text: `${f.Title} (${f.InternalName})`,
                    })),
                  ]}
                  selectedKey={ed.filterLookup.childField || ''}
                  disabled={!ed.filterLookup.parentField || !ed.filterLookup.filterOperator}
                  onChange={(_, o) =>
                    setEd((p) => ({
                      ...p,
                      filterLookup: { ...p.filterLookup, childField: o ? String(o.key) : '' },
                    }))
                  }
                />
              ) : (
                <TextField
                  label="Campo na lista filho"
                  value={ed.filterLookup.childField}
                  disabled={!ed.filterLookup.parentField || !ed.filterLookup.filterOperator}
                  onChange={(_, v) =>
                    setEd((p) => ({
                      ...p,
                      filterLookup: { ...p.filterLookup, childField: v ?? '' },
                    }))
                  }
                />
              )}
            </Stack>
          </FormManagerCollapseSection>
        )}
        {mt === 'boolean' && (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Use valor padrão acima (true/false). Visibilidade condicional: opções neste painel ou JSON do gestor.
          </Text>
        )}
        {mt !== 'text' && mt !== 'multiline' && (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Pré-visualização: {buildFieldUiRules(internalName, ed, fc).length} regra(s) gerada(s) para este campo.
          </Text>
        )}
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Aplicar" onClick={handleApply} />
          <DefaultButton text="Cancelar" onClick={onDismiss} />
        </Stack>
      </Stack>
    </Panel>
  );
};
