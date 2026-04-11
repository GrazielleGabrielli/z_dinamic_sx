import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import {
  Stack,
  Text,
  Dropdown,
  DefaultButton,
  IconButton,
  TextField,
  Checkbox,
  MessageBar,
  MessageBarType,
  IDropdownOption,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type {
  IFormStepConfig,
  TFormButtonAction,
  TFormConditionNode,
  TFormConditionOp,
} from '../../core/config/types/formManager';
import { FORM_FIXOS_STEP_ID, FORM_OCULTOS_STEP_ID } from '../../core/config/types/formManager';
import {
  CONDITION_OP_OPTIONS,
  whenUiToNode,
  whenNodeToUi,
  type IWhenUi,
} from '../../core/formManager/formManagerVisualModel';
import { FormManagerCollapseSection } from './FormManagerComponentsTab';

const BUTTON_ACTION_KIND_OPTIONS: IDropdownOption[] = [
  { key: 'showFields', text: 'Mostrar campos' },
  { key: 'hideFields', text: 'Ocultar campos' },
  { key: 'setFieldValue', text: 'Definir valor de um campo' },
  { key: 'joinFields', text: 'Juntar vários campos num campo' },
];

function reorderByIndex<T>(arr: T[], from: number, to: number): T[] {
  if (from === to || from < 0 || to < 0 || from >= arr.length || to >= arr.length) return arr.slice();
  const next = arr.slice();
  const moved = next.splice(from, 1);
  const item = moved[0] as T;
  next.splice(to, 0, item);
  return next;
}

function actionKindLabel(kind: TFormButtonAction['kind']): string {
  const hit = BUTTON_ACTION_KIND_OPTIONS.find((o) => o.key === kind);
  return hit ? String(hit.text) : kind;
}

function defaultActionForKind(kind: TFormButtonAction['kind']): TFormButtonAction {
  switch (kind) {
    case 'hideFields':
      return { kind: 'hideFields', fields: [] };
    case 'setFieldValue':
      return { kind: 'setFieldValue', field: '', valueTemplate: '' };
    case 'joinFields':
      return { kind: 'joinFields', targetField: '', valueTemplate: '', sourceFields: [], separator: ' ' };
    default:
      return { kind: 'showFields', fields: [] };
  }
}

type TWhenEditorParsed =
  | { kind: 'nested' }
  | { kind: 'single'; row: IWhenUi }
  | { kind: 'composite'; combiner: 'all' | 'any'; rows: IWhenUi[] };

function parseActionWhen(when: TFormConditionNode | undefined, defaultRow: IWhenUi): TWhenEditorParsed {
  if (!when) return { kind: 'single', row: defaultRow };
  if (when.kind === 'leaf') {
    return { kind: 'single', row: whenNodeToUi(when) ?? defaultRow };
  }
  if (when.kind === 'all' || when.kind === 'any') {
    const rows: IWhenUi[] = [];
    for (let ci = 0; ci < when.children.length; ci++) {
      const ch = when.children[ci];
      if (ch.kind !== 'leaf') return { kind: 'nested' };
      rows.push(whenNodeToUi(ch) ?? defaultRow);
    }
    if (!rows.length) return { kind: 'composite', combiner: when.kind, rows: [defaultRow] };
    return { kind: 'composite', combiner: when.kind, rows };
  }
  return { kind: 'nested' };
}

function buildWhenFromEditor(parsed: Exclude<TWhenEditorParsed, { kind: 'nested' }>): TFormConditionNode {
  if (parsed.kind === 'single') {
    return whenUiToNode(parsed.row);
  }
  return {
    kind: parsed.combiner,
    children: parsed.rows.map((r) => whenUiToNode(r)),
  };
}

function buttonSetFieldValueChoiceDropdown(
  fieldInternalName: string,
  valueTemplate: string | undefined,
  fieldMeta: IFieldMetadata[]
): { options: IDropdownOption[]; selectedKey: string } | null {
  const tpl = valueTemplate ?? '';
  const low = tpl.trim().toLowerCase();
  if (low.length >= 4 && low.slice(0, 4) === 'str:') {
    return null;
  }
  let fm: IFieldMetadata | undefined;
  for (let i = 0; i < fieldMeta.length; i++) {
    if (fieldMeta[i].InternalName === fieldInternalName) {
      fm = fieldMeta[i];
      break;
    }
  }
  const choices =
    fm && fm.MappedType === 'choice' && fm.Choices && fm.Choices.length > 0 ? fm.Choices : null;
  if (!choices) {
    return null;
  }
  const opts: IDropdownOption[] = [{ key: '', text: '—' }];
  for (let i = 0; i < choices.length; i++) {
    const c = choices[i];
    opts.push({ key: c, text: c });
  }
  if (tpl && choices.indexOf(tpl) === -1) {
    opts.push({ key: tpl, text: `${tpl} (valor atual)` });
  }
  return { options: opts, selectedKey: tpl };
}

export interface IFormManagerChainedActionsBlockProps {
  actions: TFormButtonAction[];
  patchAction: (actionIndex: number, next: TFormButtonAction) => void;
  removeAction: (actionIndex: number) => void;
  addAction: () => void;
  patchActionCondition: (actionIndex: number, when: TFormConditionNode | undefined) => void;
  reactKeysPrefix: string;
  meta: IFieldMetadata[];
  metaSortedForPool: IFieldMetadata[];
  steps: IFormStepConfig[];
  fieldOptions: IDropdownOption[];
  loading: boolean;
  getDefaultWhenUi: () => IWhenUi;
}

export function FormManagerChainedActionsBlock(props: IFormManagerChainedActionsBlockProps): JSX.Element {
  const {
    actions,
    patchAction,
    removeAction,
    addAction,
    patchActionCondition,
    reactKeysPrefix,
    meta,
    metaSortedForPool,
    steps,
    fieldOptions,
    loading,
    getDefaultWhenUi,
  } = props;

  const [actionOpen, setActionOpen] = useState<Record<number, boolean>>({});
  const prevActionLenRef = useRef(actions.length);
  useEffect(() => {
    if (actions.length < prevActionLenRef.current) {
      setActionOpen({});
    }
    prevActionLenRef.current = actions.length;
  }, [actions.length]);

  return (
    <>
      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
        Ações (por ordem)
      </Text>
      {actions.map((act, ai) => (
        <FormManagerCollapseSection
          key={`${reactKeysPrefix}-act-${ai}`}
          title={`Ação ${ai + 1} · ${actionKindLabel(act.kind)}`}
          isOpen={actionOpen[ai] === true}
          onToggle={() =>
            setActionOpen((prev) => ({
              ...prev,
              [ai]: !(prev[ai] === true),
            }))
          }
        >
          <Stack tokens={{ childrenGap: 8 }}>
          <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
            <Dropdown
              label="Tipo"
              options={BUTTON_ACTION_KIND_OPTIONS}
              selectedKey={act.kind}
              onChange={(_, o) => {
                if (!o) return;
                patchAction(ai, defaultActionForKind(String(o.key) as TFormButtonAction['kind']));
              }}
            />
            <DefaultButton text="Remover ação" onClick={() => removeAction(ai)} />
          </Stack>
          {(act.kind === 'showFields' || act.kind === 'hideFields') && (
            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                Campos
              </Text>
              <Stack
                tokens={{ childrenGap: 6 }}
                styles={{
                  root: {
                    maxHeight: 280,
                    overflowY: 'auto',
                    border: '1px solid #edebe9',
                    borderRadius: 4,
                    padding: 8,
                  },
                }}
              >
                {act.fields
                  .filter((fn) => !metaSortedForPool.some((m) => m.InternalName === fn))
                  .map((fn) => (
                    <Checkbox
                      key={`${reactKeysPrefix}-orphan-${ai}-${fn}`}
                      label={`${fn} (referência guardada)`}
                      checked
                      onChange={(_, c) => {
                        if (c) return;
                        patchAction(ai, {
                          ...act,
                          fields: act.fields.filter((x) => x !== fn),
                        });
                      }}
                    />
                  ))}
                {metaSortedForPool.map((m) => {
                  const fn = m.InternalName;
                  const checked = act.fields.indexOf(fn) !== -1;
                  return (
                    <Checkbox
                      key={fn}
                      label={`${m.Title} (${fn})`}
                      checked={checked}
                      onChange={(_, c) => {
                        let next: string[];
                        if (c) {
                          next = checked ? act.fields : act.fields.concat([fn]);
                        } else {
                          next = act.fields.filter((x) => x !== fn);
                        }
                        patchAction(ai, { ...act, fields: next });
                      }}
                    />
                  );
                })}
                {!metaSortedForPool.length && (
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    {loading ? 'A carregar campos da lista…' : 'Nenhum campo disponível para selecionar.'}
                  </Text>
                )}
              </Stack>
            </Stack>
          )}
          {act.kind === 'showFields' &&
            steps.filter((s) => s.id !== FORM_OCULTOS_STEP_ID && s.id !== FORM_FIXOS_STEP_ID).length > 1 && (
            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Campos só na aba Ocultos: escolha em que etapa devem surgir ao executar esta ação.
              </Text>
              <Dropdown
                label="Etapa onde mostrar"
                options={[
                  { key: '', text: '— escolher —' },
                  ...steps
                    .filter((s) => s.id !== FORM_OCULTOS_STEP_ID && s.id !== FORM_FIXOS_STEP_ID)
                    .map((s) => ({ key: s.id, text: s.title })),
                ]}
                selectedKey={act.displayOnStepId ?? ''}
                onChange={(_, o) => {
                  if (!o) return;
                  const key = String(o.key);
                  patchAction(ai, {
                    kind: 'showFields',
                    fields: act.fields,
                    ...(key ? { displayOnStepId: key } : {}),
                    ...(act.when ? { when: act.when } : {}),
                  });
                }}
              />
            </Stack>
          )}
          {act.kind === 'setFieldValue' &&
            (() => {
              const choiceVal = buttonSetFieldValueChoiceDropdown(act.field, act.valueTemplate, meta);
              return (
                <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                  <Dropdown
                    label="Campo"
                    options={[{ key: '', text: '—' }, ...fieldOptions]}
                    selectedKey={act.field || ''}
                    onChange={(_, o) =>
                      patchAction(ai, {
                        ...act,
                        field: o ? String(o.key) : '',
                      })
                    }
                  />
                  {choiceVal ? (
                    <Dropdown
                      label="Valor"
                      styles={{ root: { minWidth: 280 } }}
                      options={choiceVal.options}
                      selectedKey={choiceVal.selectedKey}
                      onChange={(_, o) =>
                        patchAction(ai, {
                          ...act,
                          valueTemplate: o ? String(o.key) : '',
                        })
                      }
                    />
                  ) : (
                    <TextField
                      label="Valor fixo ou str:{{Campo}}"
                      styles={{ root: { minWidth: 280 } }}
                      value={act.valueTemplate}
                      onChange={(_, v) => patchAction(ai, { ...act, valueTemplate: v ?? '' })}
                    />
                  )}
                </Stack>
              );
            })()}
          {act.kind === 'joinFields' && (
            <Stack tokens={{ childrenGap: 10 }}>
              <Dropdown
                label="Campo destino"
                options={[{ key: '', text: '—' }, ...fieldOptions]}
                selectedKey={act.targetField || ''}
                onChange={(_, o) =>
                  patchAction(ai, {
                    ...act,
                    targetField: o ? String(o.key) : '',
                  })
                }
              />
              <TextField
                label="Modelo de texto"
                multiline
                rows={5}
                value={act.valueTemplate ?? ''}
                onChange={(_, v) => patchAction(ai, { ...act, valueTemplate: v ?? '' })}
                description="Placeholders: {{NomeInterno}}. Ex.: Número: {{Numero}} — Obra: {{Title}}. Vazio = junção simples com separador e ordem da lista abaixo."
              />
              <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                Campos na ordem (modo simples ou botão + para acrescentar placeholders ao modelo)
              </Text>
              <Dropdown
                key={`join-add-${reactKeysPrefix}-${ai}-${act.sourceFields.join('|')}`}
                label="Adicionar campo à ordem"
                options={[
                  { key: '', text: '—' },
                  ...metaSortedForPool
                    .filter((m) => act.sourceFields.indexOf(m.InternalName) === -1)
                    .map((m) => ({
                      key: m.InternalName,
                      text: `${m.Title} (${m.InternalName})`,
                    })),
                ]}
                selectedKey=""
                onChange={(_, o) => {
                  if (!o || o.key === '') return;
                  const k = String(o.key);
                  if (act.sourceFields.indexOf(k) !== -1) return;
                  patchAction(ai, {
                    ...act,
                    sourceFields: act.sourceFields.concat([k]),
                  });
                }}
              />
              <Stack tokens={{ childrenGap: 6 }}>
                {act.sourceFields.map((fn, idx) => {
                  const m = metaSortedForPool.find((x) => x.InternalName === fn);
                  const label = m ? `${m.Title} (${fn})` : `${fn} (referência guardada)`;
                  return (
                    <Stack
                      horizontal
                      verticalAlign="center"
                      key={`join-row-${reactKeysPrefix}-${ai}-${idx}-${fn}`}
                      tokens={{ childrenGap: 6 }}
                      wrap
                    >
                      <Text styles={{ root: { flex: '1 1 200px', minWidth: 0 } }}>{label}</Text>
                      <IconButton
                        iconProps={{ iconName: 'ChevronUp' }}
                        disabled={idx === 0}
                        title="Subir"
                        onClick={() =>
                          patchAction(ai, {
                            ...act,
                            sourceFields: reorderByIndex(act.sourceFields, idx, idx - 1),
                          })
                        }
                      />
                      <IconButton
                        iconProps={{ iconName: 'ChevronDown' }}
                        disabled={idx === act.sourceFields.length - 1}
                        title="Descer"
                        onClick={() =>
                          patchAction(ai, {
                            ...act,
                            sourceFields: reorderByIndex(act.sourceFields, idx, idx + 1),
                          })
                        }
                      />
                      <IconButton
                        iconProps={{ iconName: 'Add' }}
                        title={`Acrescentar {{${fn}}} ao modelo`}
                        onClick={() => {
                          const cur = act.valueTemplate ?? '';
                          patchAction(ai, {
                            ...act,
                            valueTemplate: cur + `{{${fn}}}`,
                          });
                        }}
                      />
                      <IconButton
                        iconProps={{ iconName: 'Delete' }}
                        title="Remover da ordem"
                        onClick={() =>
                          patchAction(ai, {
                            ...act,
                            sourceFields: act.sourceFields.filter((_, i) => i !== idx),
                          })
                        }
                      />
                    </Stack>
                  );
                })}
              </Stack>
              <TextField
                label="Separador (só com modelo vazio)"
                value={act.separator}
                onChange={(_, v) => patchAction(ai, { ...act, separator: v ?? ' ' })}
              />
            </Stack>
          )}
          <Checkbox
            label="Só executar esta ação se (avalia valores já alterados pelas ações acima)"
            checked={!!act.when}
            onChange={(_, c) => {
              if (c) {
                patchAction(ai, {
                  ...act,
                  when: whenUiToNode(getDefaultWhenUi()),
                });
              } else {
                const { when: _rm, ...rest } = act as TFormButtonAction & { when?: TFormConditionNode };
                patchAction(ai, rest as TFormButtonAction);
              }
            }}
          />
          {act.when &&
            (() => {
              const defaultRow = getDefaultWhenUi();
              const parsed = parseActionWhen(act.when, defaultRow);
              if (parsed.kind === 'nested') {
                return (
                  <MessageBar messageBarType={MessageBarType.warning}>
                    Condição com agrupamentos aninhados (E dentro de OU, etc.): edite no JSON avançado ou use só um
                    nível — várias linhas combinadas apenas com «Todas (E)» ou «Pelo menos uma (OU)».
                  </MessageBar>
                );
              }
              const modeKey: 'single' | 'all' | 'any' =
                parsed.kind === 'single' ? 'single' : parsed.combiner;
              const rows: IWhenUi[] = parsed.kind === 'single' ? [parsed.row] : parsed.rows;
              const write = (next: Exclude<TWhenEditorParsed, { kind: 'nested' }>): void => {
                patchActionCondition(ai, buildWhenFromEditor(next));
              };
              const patchRow = (ri: number, partial: Partial<IWhenUi>): void => {
                const nextRows = rows.map((r, i) => (i === ri ? { ...r, ...partial } : r));
                if (modeKey === 'single') {
                  write({ kind: 'single', row: nextRows[0] ?? defaultRow });
                } else {
                  write({ kind: 'composite', combiner: modeKey, rows: nextRows });
                }
              };
              return (
                <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 4 } }}>
                  <Dropdown
                    label="Como combinar as condições"
                    options={[
                      { key: 'single', text: 'Uma condição' },
                      { key: 'all', text: 'Todas têm de ser verdade (E)' },
                      { key: 'any', text: 'Pelo menos uma verdadeira (OU)' },
                    ]}
                    selectedKey={modeKey}
                    onChange={(_, o) => {
                      if (!o) return;
                      const k = String(o.key) as 'single' | 'all' | 'any';
                      const cur = rows.slice();
                      if (k === 'single') {
                        write({ kind: 'single', row: cur[0] ?? defaultRow });
                        return;
                      }
                      const nextRows = cur.length >= 2 ? cur : [cur[0] ?? defaultRow, { ...defaultRow }];
                      write({ kind: 'composite', combiner: k, rows: nextRows });
                    }}
                  />
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Cada linha é uma condição (campo, operador, valor). Em modo E, todas têm de ser verdadeiras; em
                    modo OU, basta uma.
                  </Text>
                  {rows.map((row, ri) => (
                    <Stack
                      key={`${reactKeysPrefix}-when-${ai}-${ri}`}
                      tokens={{ childrenGap: 8 }}
                      styles={{
                        root: {
                          border: '1px solid #edebe9',
                          borderRadius: 4,
                          padding: 10,
                          background: '#faf9f8',
                        },
                      }}
                    >
                      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                        <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                          Condição {ri + 1}
                        </Text>
                        {modeKey !== 'single' && rows.length > 1 && (
                          <DefaultButton
                            text="Remover linha"
                            onClick={() => {
                              const nextRows = rows.filter((_, i) => i !== ri);
                              if (nextRows.length === 1) {
                                write({ kind: 'single', row: nextRows[0] ?? defaultRow });
                              } else {
                                write({ kind: 'composite', combiner: modeKey, rows: nextRows });
                              }
                            }}
                          />
                        )}
                      </Stack>
                      <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                        <Dropdown
                          label="Campo"
                          options={fieldOptions}
                          selectedKey={row.field}
                          onChange={(_, o) => o && patchRow(ri, { field: String(o.key) })}
                        />
                        <Dropdown
                          label="Operador"
                          options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                          selectedKey={row.op}
                          onChange={(_, o) => o && patchRow(ri, { op: o.key as TFormConditionOp })}
                        />
                        <Dropdown
                          label="Comparar com"
                          options={[
                            { key: 'literal', text: 'Texto fixo' },
                            { key: 'field', text: 'Outro campo' },
                            { key: 'token', text: 'Token' },
                          ]}
                          selectedKey={row.compareKind}
                          onChange={(_, o) =>
                            o && patchRow(ri, { compareKind: o.key as IWhenUi['compareKind'] })
                          }
                        />
                        <TextField
                          label="Valor"
                          value={row.compareValue}
                          onChange={(_, v) => patchRow(ri, { compareValue: v ?? '' })}
                          disabled={
                            row.op === 'isEmpty' ||
                            row.op === 'isFilled' ||
                            row.op === 'isTrue' ||
                            row.op === 'isFalse'
                          }
                        />
                      </Stack>
                    </Stack>
                  ))}
                  {modeKey !== 'single' && (
                    <DefaultButton
                      text="Adicionar condição"
                      onClick={() => {
                        write({
                          kind: 'composite',
                          combiner: modeKey,
                          rows: rows.concat([{ ...defaultRow }]),
                        });
                      }}
                    />
                  )}
                </Stack>
              );
            })()}
          </Stack>
        </FormManagerCollapseSection>
      ))}
      <DefaultButton text="Adicionar ação" onClick={addAction} />
    </>
  );
}
