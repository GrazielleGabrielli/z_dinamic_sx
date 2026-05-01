import * as React from 'react';
import { useMemo, useCallback, useRef } from 'react';
import { Stack, Text, PrimaryButton, DefaultButton, Dropdown, TextField, IDropdownOption } from '@fluentui/react';
import type { TFormConditionOp, TFormRule } from '../../core/config/types/formManager';
import {
  compileConditionalCard,
  CONDITIONAL_EFFECT_OPTIONS,
  CONDITION_OP_OPTIONS,
  customRulesOnly,
  describeConditionalCardPT,
  describeRule,
  mergeCardRulesIntoAll,
  newCardId,
  parseConditionalCardsFromRules,
  templateConditionalShowWhenCompare,
  templateConditionalShowWhenEquals,
  templateFieldRulesChoiceRequiresOther,
  type IConditionalEffectUi,
  type IConditionalRuleCard,
  type IWhenUi,
  type TConditionalEffectKind,
} from '../../core/formManager/formManagerVisualModel';

function parseCsvFieldNames(s: string): string[] {
  return s
    .split(/[,;]/)
    .map((x) => x.trim())
    .filter(Boolean);
}

function fieldNamesToCsv(names: string[]): string {
  return names.join(', ');
}

function emptyEffect(): IConditionalEffectUi {
  return { kind: 'showField', targetField: '' };
}

export interface IFormManagerLinkedChildConditionalRulesBlockProps {
  rules: TFormRule[];
  fieldOptions: IDropdownOption[];
  onRulesChange: (next: TFormRule[]) => void;
}

export const FormManagerLinkedChildConditionalRulesBlock: React.FC<
  IFormManagerLinkedChildConditionalRulesBlockProps
> = ({ rules, fieldOptions, onRulesChange }) => {
  const rulesRef = useRef(rules);
  rulesRef.current = rules;

  const conditionalCards = useMemo(() => parseConditionalCardsFromRules(rules).cards, [rules]);
  const customs = useMemo(() => customRulesOnly(rules), [rules]);

  const setCards = useCallback(
    (cards: IConditionalRuleCard[]) => {
      onRulesChange(mergeCardRulesIntoAll(rulesRef.current, cards));
    },
    [onRulesChange]
  );

  const readCards = useCallback(
    (): IConditionalRuleCard[] => parseConditionalCardsFromRules(rulesRef.current).cards,
    []
  );

  const addConditionalCard = useCallback(() => {
    const df = String(fieldOptions[0]?.key ?? 'Title');
    const card: IConditionalRuleCard = {
      id: newCardId(),
      when: { field: df, op: 'eq', compareKind: 'literal', compareValue: '' },
      effects: [emptyEffect()],
    };
    setCards(readCards().concat([card]));
  }, [fieldOptions, readCards, setCards]);

  const patchCard = useCallback(
    (index: number, patch: Partial<IConditionalRuleCard>) => {
      const cur = readCards();
      const c = cur[index];
      if (!c) return;
      const next = cur.map((x, i) => (i === index ? { ...x, ...patch } : x));
      setCards(next);
    },
    [readCards, setCards]
  );

  const patchWhen = useCallback(
    (index: number, w: Partial<IWhenUi>) => {
      const cur = readCards();
      const c = cur[index];
      if (!c) return;
      const next = cur.map((x, i) => (i === index ? { ...x, when: { ...x.when, ...w } } : x));
      setCards(next);
    },
    [readCards, setCards]
  );

  const patchEffect = useCallback(
    (cardIndex: number, effIndex: number, patch: Partial<IConditionalEffectUi>) => {
      const cur = readCards();
      const c = cur[cardIndex];
      if (!c) return;
      const effects = c.effects.map((e, i) => (i === effIndex ? { ...e, ...patch } : e));
      setCards(cur.map((x, i) => (i === cardIndex ? { ...x, effects } : x)));
    },
    [readCards, setCards]
  );

  const addEffect = useCallback(
    (cardIndex: number) => {
      const cur = readCards();
      const c = cur[cardIndex];
      if (!c) return;
      patchCard(cardIndex, { effects: c.effects.concat([emptyEffect()]) });
    },
    [readCards, patchCard]
  );

  const removeEffect = useCallback(
    (cardIndex: number, effIndex: number) => {
      const cur = readCards();
      const c = cur[cardIndex];
      if (!c) return;
      patchCard(cardIndex, { effects: c.effects.filter((_, i) => i !== effIndex) });
    },
    [readCards, patchCard]
  );

  const duplicateCard = useCallback(
    (index: number) => {
      const cur = readCards();
      const c = cur[index];
      if (!c) return;
      const copy: IConditionalRuleCard = {
        ...c,
        id: newCardId(),
        effects: c.effects.map((e) => ({ ...e })),
      };
      const next = cur.slice();
      next.splice(index + 1, 0, copy);
      setCards(next);
    },
    [readCards, setCards]
  );

  const removeCard = useCallback(
    (index: number) => {
      setCards(readCards().filter((_, i) => i !== index));
    },
    [readCards, setCards]
  );

  const applyPresetConditional = useCallback(
    (preset: 'showWhenEq' | 'choiceRequire' | 'showWhenContains' | 'showWhenNe' | 'showWhenGt') => {
      const a = String(fieldOptions[0]?.key ?? 'Title');
      const b = String(fieldOptions[1]?.key ?? a);
      const cur = readCards();
      if (preset === 'showWhenEq') {
        const card = templateConditionalShowWhenEquals(a, '', b);
        setCards(cur.concat([card]));
      } else if (preset === 'choiceRequire') {
        setCards(cur.concat([templateFieldRulesChoiceRequiresOther(a, '', b)]));
      } else if (preset === 'showWhenContains') {
        setCards(cur.concat([templateConditionalShowWhenCompare(a, 'contains', '', b)]));
      } else if (preset === 'showWhenNe') {
        setCards(cur.concat([templateConditionalShowWhenCompare(a, 'ne', '', b)]));
      } else {
        setCards(cur.concat([templateConditionalShowWhenCompare(a, 'gt', '0', b)]));
      }
    },
    [fieldOptions, readCards, setCards]
  );

  const opts = fieldOptions.length ? fieldOptions : [{ key: 'Title', text: 'Title' }];

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 12 } }}>
      <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
        Regras condicionais (só campos desta lista filha)
      </Text>
    
      <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
        <PrimaryButton text="Nova regra" onClick={addConditionalCard} disabled={!fieldOptions.length} />
        <DefaultButton
          text="Modelo: mostrar B quando A = valor"
          onClick={() => applyPresetConditional('showWhenEq')}
          disabled={!fieldOptions.length}
        />
        <DefaultButton
          text="Modelo: mostrar B quando A contém texto"
          onClick={() => applyPresetConditional('showWhenContains')}
          disabled={!fieldOptions.length}
        />
        <DefaultButton
          text="Modelo: mostrar B quando A ≠ valor"
          onClick={() => applyPresetConditional('showWhenNe')}
          disabled={!fieldOptions.length}
        />
        <DefaultButton
          text="Modelo: mostrar B quando A > número"
          onClick={() => applyPresetConditional('showWhenGt')}
          disabled={!fieldOptions.length}
        />
        <DefaultButton
          text="Modelo: obrigar B quando A = valor"
          onClick={() => applyPresetConditional('choiceRequire')}
          disabled={!fieldOptions.length}
        />
      </Stack>
  
      {conditionalCards.map((card, ci) => (
        <Stack
          key={card.id}
          styles={{ root: { border: '1px solid #edebe9', padding: 12, borderRadius: 4 } }}
          tokens={{ childrenGap: 8 }}
        >
          <Stack horizontal horizontalAlign="space-between">
            <Text styles={{ root: { fontWeight: 600 } }}>{describeConditionalCardPT(card)}</Text>
            <Stack horizontal tokens={{ childrenGap: 4 }}>
              <DefaultButton text="Duplicar" onClick={() => duplicateCard(ci)} />
              <DefaultButton text="Excluir" onClick={() => removeCard(ci)} />
            </Stack>
          </Stack>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Quando
          </Text>
          <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
            <Dropdown
              label="Campo"
              options={opts}
              selectedKey={card.when.field}
              onChange={(_, o) => o && patchWhen(ci, { field: String(o.key) })}
            />
            <Dropdown
              label="Operador"
              options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
              selectedKey={card.when.op}
              onChange={(_, o) => o && patchWhen(ci, { op: o.key as TFormConditionOp })}
              calloutProps={{ calloutMaxHeight: 400 }}
              styles={{ dropdown: { minWidth: 200 } }}
            />
            <Dropdown
              label="Comparar com"
              options={[
                { key: 'literal', text: 'Texto fixo' },
                { key: 'field', text: 'Outro campo' },
                { key: 'token', text: 'Token' },
              ]}
              selectedKey={card.when.compareKind}
              onChange={(_, o) => o && patchWhen(ci, { compareKind: o.key as IWhenUi['compareKind'] })}
            />
            <TextField
              label="Valor"
              value={card.when.compareValue}
              onChange={(_, v) => patchWhen(ci, { compareValue: v ?? '' })}
              disabled={
                card.when.op === 'isEmpty' ||
                card.when.op === 'isFilled' ||
                card.when.op === 'isTrue' ||
                card.when.op === 'isFalse'
              }
            />
          </Stack>
          <TextField
            label="Incluir: grupos SharePoint (títulos, vírgula)"
            description="Vazio = qualquer utilizador. Com valores, só aplica se o utilizador pertencer a pelo menos um grupo."
            value={fieldNamesToCsv(card.groupTitles ?? [])}
            onChange={(_, v) => {
              const parsed = parseCsvFieldNames(v ?? '');
              patchCard(ci, { groupTitles: parsed.length ? parsed : undefined });
            }}
          />
          <TextField
            label="Excluir: grupos SharePoint (títulos, vírgula)"
            description="Vazio = não excluir. Com valores, a regra não aplica a quem pertencer a algum destes grupos."
            value={fieldNamesToCsv(card.excludeGroupTitles ?? [])}
            onChange={(_, v) => {
              const parsed = parseCsvFieldNames(v ?? '');
              patchCard(ci, { excludeGroupTitles: parsed.length ? parsed : undefined });
            }}
          />
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Então
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Uma linha por combinação (efeito + campo). «Adicionar efeito» acrescenta outra linha com a mesma
            condição «Quando».
          </Text>
          {card.effects.map((eff, ei) => (
            <Stack
              key={`${card.id}-eff-${ei}`}
              tokens={{ childrenGap: 8 }}
              styles={{
                root: {
                  paddingTop: ei > 0 ? 10 : 0,
                  marginTop: ei > 0 ? 10 : 0,
                  borderTop: ei > 0 ? '1px solid #edebe9' : undefined,
                },
              }}
            >
              <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
                Efeito {ei + 1}
              </Text>
              <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                <Dropdown
                  label="Efeito"
                  options={CONDITIONAL_EFFECT_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                  selectedKey={eff.kind}
                  onChange={(_, o) => o && patchEffect(ci, ei, { kind: o.key as TConditionalEffectKind })}
                />
                {eff.kind !== 'message' && (
                  <Dropdown
                    label="Campo alvo"
                    options={[{ key: '', text: '—' }, ...opts]}
                    selectedKey={eff.targetField ?? ''}
                    onChange={(_, o) => patchEffect(ci, ei, { targetField: o ? String(o.key) : undefined })}
                  />
                )}
                {eff.kind === 'message' && (
                  <>
                    <Dropdown
                      label="Tipo"
                      options={[
                        { key: 'info', text: 'Info' },
                        { key: 'warning', text: 'Aviso' },
                        { key: 'error', text: 'Erro' },
                      ]}
                      selectedKey={eff.messageVariant ?? 'info'}
                      onChange={(_, o) =>
                        o && patchEffect(ci, ei, { messageVariant: o.key as 'info' | 'warning' | 'error' })
                      }
                    />
                    <TextField
                      label="Texto"
                      value={eff.messageText ?? ''}
                      onChange={(_, v) => patchEffect(ci, ei, { messageText: v ?? '' })}
                    />
                  </>
                )}
                <DefaultButton text="Remover efeito" onClick={() => removeEffect(ci, ei)} />
              </Stack>
            </Stack>
          ))}
          <DefaultButton text="Adicionar efeito" onClick={() => addEffect(ci)} />
          <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
            Prévia: {compileConditionalCard(card).length} regra(s) gerada(s)
          </Text>
        </Stack>
      ))}
      {!conditionalCards.length && (
        <Text variant="small">Nenhuma regra condicional nesta lista vinculada.</Text>
      )}
      {!!customs.length && (
        <Stack tokens={{ childrenGap: 6 }}>
          <Text variant="smallPlus" styles={{ root: { fontWeight: 600 } }}>
            Regras só no motor (não editadas por esta UI)
          </Text>
          {customs.map((r) => (
            <Text key={r.id} variant="small" styles={{ root: { color: '#605e5c' } }}>
              {r.id}: {describeRule(r)}
            </Text>
          ))}
        </Stack>
      )}
    </Stack>
  );
};
