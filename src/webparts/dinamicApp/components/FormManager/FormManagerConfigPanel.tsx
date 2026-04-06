import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import {
  Panel,
  PanelType,
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  TextField,
  Dropdown,
  IDropdownOption,
  Checkbox,
  Spinner,
  MessageBar,
  MessageBarType,
  Pivot,
  PivotItem,
  Link,
  Icon,
} from '@fluentui/react';
import { FieldsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import type {
  IFormManagerConfig,
  IFormFieldConfig,
  IFormSectionConfig,
  IFormStepConfig,
  TFormConditionOp,
  TFormRule,
} from '../../core/config/types/formManager';
import { getDefaultFormManagerConfig } from '../../core/config/utils';
import { sanitizeFormManagerConfig } from '../../core/formManager/sanitizeFormManagerConfig';
import {
  buildFieldUiRules,
  compileConditionalCard,
  customRulesOnly,
  describeConditionalCardPT,
  describeRule,
  mergeAttachmentUiRule,
  mergeCardRulesIntoAll,
  mergeFieldRules,
  newCardId,
  parseAttachmentUiRule,
  parseConditionalCardsFromRules,
  countFieldUiRules,
  CONDITIONAL_EFFECT_OPTIONS,
  CONDITION_OP_OPTIONS,
  type IConditionalEffectUi,
  type IConditionalRuleCard,
  type IWhenUi,
  type TConditionalEffectKind,
  fieldRuleStateFromRules,
  templateConditionalShowWhenEquals,
  templateFieldRulesChoiceRequiresOther,
} from '../../core/formManager/formManagerVisualModel';
import { FormFieldRulesPanel } from './FormFieldRulesPanel';

function newId(prefix: string): string {
  return `${prefix}_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
}

function reorderByIndex<T>(arr: T[], from: number, to: number): T[] {
  if (from === to || from < 0 || to < 0 || from >= arr.length || to >= arr.length) return arr.slice();
  const next = arr.slice();
  const moved = next.splice(from, 1);
  const item = moved[0] as T;
  next.splice(to, 0, item);
  return next;
}

const DND_FIELD = 'fm/field:';
const DND_STEP = 'fm/step:';
const DND_MCOL = 'fm/mcol:';
const DND_POOL = 'fm/pool:';
const DND_FS = 'fm/fs:';

function dragPayload(kind: string, index: number): string {
  return kind + String(index);
}

function parseDragIndex(data: string, prefix: string): number | undefined {
  if (data.indexOf(prefix) !== 0) return undefined;
  const n = parseInt(data.slice(prefix.length), 10);
  return isNaN(n) ? undefined : n;
}

function dragPayloadPool(internalName: string): string {
  return DND_POOL + encodeURIComponent(internalName);
}

function parsePoolDrag(data: string): string | undefined {
  if (data.indexOf(DND_POOL) !== 0) return undefined;
  try {
    return decodeURIComponent(data.slice(DND_POOL.length));
  } catch {
    return undefined;
  }
}

function dragPayloadFieldInStep(stepIdx: number, idxInStep: number, internalName: string): string {
  return DND_FS + String(stepIdx) + ':' + String(idxInStep) + ':' + encodeURIComponent(internalName);
}

function parseFieldInStepDrag(data: string): { fromStep: number; fromIdx: number; name: string } | undefined {
  if (data.indexOf(DND_FS) !== 0) return undefined;
  const rest = data.slice(DND_FS.length);
  const p1 = rest.indexOf(':');
  const p2 = rest.indexOf(':', p1 + 1);
  if (p1 === -1 || p2 === -1) return undefined;
  const fromStep = parseInt(rest.slice(0, p1), 10);
  const fromIdx = parseInt(rest.slice(p1 + 1, p2), 10);
  let name = '';
  try {
    name = decodeURIComponent(rest.slice(p2 + 1));
  } catch {
    return undefined;
  }
  if (isNaN(fromStep) || isNaN(fromIdx) || !name) return undefined;
  return { fromStep, fromIdx, name };
}

function insertFieldNameIntoStep(
  st: IFormStepConfig[],
  fieldName: string,
  toStepIdx: number,
  insertBefore: number
): IFormStepConfig[] {
  const next = st.map((s) => ({
    ...s,
    fieldNames: s.fieldNames.filter((n) => n !== fieldName),
  }));
  const tgt = next[toStepIdx];
  if (!tgt) return next;
  const fn = tgt.fieldNames.slice();
  const pos = Math.max(0, Math.min(insertBefore, fn.length));
  fn.splice(pos, 0, fieldName);
  next[toStepIdx] = { ...tgt, fieldNames: fn };
  return next;
}

function fieldsAlignedToSteps(flds: IFormFieldConfig[], st: IFormStepConfig[]): IFormFieldConfig[] {
  const byName: Record<string, IFormFieldConfig> = {};
  for (let i = 0; i < flds.length; i++) {
    byName[flds[i].internalName] = flds[i];
  }
  const out: IFormFieldConfig[] = [];
  const seen: Record<string, boolean> = {};
  for (let s = 0; s < st.length; s++) {
    const sid = st[s].id;
    for (let j = 0; j < st[s].fieldNames.length; j++) {
      const n = st[s].fieldNames[j];
      const fc = byName[n];
      if (fc) {
        out.push({ ...fc, sectionId: sid });
        seen[n] = true;
      }
    }
  }
  for (let i = 0; i < flds.length; i++) {
    const n = flds[i].internalName;
    if (!seen[n]) {
      out.push({ ...flds[i], sectionId: st[0]?.id ?? flds[i].sectionId });
    }
  }
  return out;
}

function resyncStepsOrderFromFields(flds: IFormFieldConfig[], st: IFormStepConfig[]): IFormStepConfig[] {
  const orderMap: Record<string, number> = {};
  for (let i = 0; i < flds.length; i++) {
    orderMap[flds[i].internalName] = i;
  }
  return st.map((s) => ({
    ...s,
    fieldNames: s.fieldNames.slice().sort((a, b) => (orderMap[a] ?? 99999) - (orderMap[b] ?? 99999)),
  }));
}

function numOpt(s: string): number | undefined {
  const t = s.trim();
  if (!t) return undefined;
  const n = Number(t);
  return isNaN(n) ? undefined : n;
}

function defaultWhenUi(meta: IFieldMetadata[]): IWhenUi {
  const f = meta[0]?.InternalName ?? 'Title';
  return { field: f, op: 'eq', compareKind: 'literal', compareValue: '' };
}

function emptyEffect(): IConditionalEffectUi {
  return { kind: 'showField', targetField: '' };
}

function sectionsFromSteps(steps: IFormStepConfig[]): IFormSectionConfig[] {
  const out: IFormSectionConfig[] = [];
  for (let i = 0; i < steps.length; i++) {
    out.push({ id: steps[i].id, title: steps[i].title, visible: true });
  }
  return out;
}

function inferStepsFromLegacy(sections: IFormSectionConfig[], flds: IFormFieldConfig[]): IFormStepConfig[] {
  const out: IFormStepConfig[] = [];
  const defaultSid = sections[0]?.id ?? 'main';
  for (let i = 0; i < sections.length; i++) {
    const sec = sections[i];
    const fieldNames: string[] = [];
    for (let j = 0; j < flds.length; j++) {
      const sid = flds[j].sectionId ?? defaultSid;
      if (sid === sec.id) fieldNames.push(flds[j].internalName);
    }
    out.push({ id: sec.id, title: sec.title, fieldNames: fieldNames.slice() });
  }
  if (out.length === 0) {
    const fn: string[] = [];
    for (let k = 0; k < flds.length; k++) {
      fn.push(flds[k].internalName);
    }
    return [{ id: 'main', title: 'Geral', fieldNames: fn }];
  }
  return out;
}

function ensureAtLeastOneStep(st: IFormStepConfig[]): IFormStepConfig[] {
  if (st.length > 0) return st;
  return [{ id: 'main', title: 'Geral', fieldNames: [] }];
}

function buildInitialFieldsAndSteps(v: IFormManagerConfig): {
  fields: IFormFieldConfig[];
  steps: IFormStepConfig[];
} {
  const stepsSrc =
    v.steps && v.steps.length > 0
      ? v.steps.map((st) => ({ ...st, fieldNames: st.fieldNames.slice() }))
      : inferStepsFromLegacy(v.sections, v.fields);
  return normalizeFieldsIntoSteps(
    v.fields.map((f) => ({ ...f })),
    ensureAtLeastOneStep(stepsSrc)
  );
}

function normalizeFieldsIntoSteps(
  flds: IFormFieldConfig[],
  stepsIn: IFormStepConfig[]
): { fields: IFormFieldConfig[]; steps: IFormStepConfig[] } {
  const base = ensureAtLeastOneStep(
    stepsIn.map((s) => ({ ...s, fieldNames: s.fieldNames.slice() }))
  );
  const nextSteps = base.map((s) => ({ ...s, fieldNames: [] as string[] }));
  const nextFields = flds.map((f) => ({ ...f }));
  for (let i = 0; i < nextFields.length; i++) {
    const name = nextFields[i].internalName;
    let stepIdx = 0;
    let assigned = false;
    for (let j = 0; j < base.length; j++) {
      if (base[j].fieldNames.indexOf(name) !== -1) {
        stepIdx = j;
        assigned = true;
        break;
      }
    }
    if (!assigned) {
      const sid = nextFields[i].sectionId ?? base[0].id;
      stepIdx = 0;
      for (let k = 0; k < base.length; k++) {
        if (base[k].id === sid) {
          stepIdx = k;
          break;
        }
      }
    }
    nextSteps[stepIdx].fieldNames.push(name);
    nextFields[i].sectionId = nextSteps[stepIdx].id;
  }
  return { fields: nextFields, steps: nextSteps };
}

export interface IFormManagerConfigPanelProps {
  isOpen: boolean;
  listTitle: string;
  value: IFormManagerConfig;
  onSave: (next: IFormManagerConfig) => void;
  onDismiss: () => void;
}

export const FormManagerConfigPanel: React.FC<IFormManagerConfigPanelProps> = ({
  isOpen,
  listTitle,
  value,
  onSave,
  onDismiss,
}) => {
  const [fields, setFields] = useState<IFormFieldConfig[]>(() => buildInitialFieldsAndSteps(value).fields);
  const [rules, setRules] = useState<TFormRule[]>(() => value.rules ?? []);
  const [steps, setSteps] = useState<IFormStepConfig[]>(() => buildInitialFieldsAndSteps(value).steps);
  const [helpJson, setHelpJson] = useState(() => JSON.stringify(value.dynamicHelp ?? [], null, 2));
  const [managerColumnFields, setManagerColumnFields] = useState<string[]>(() => value.managerColumnFields ?? []);
  const [attachMin, setAttachMin] = useState('');
  const [attachMax, setAttachMax] = useState('');
  const [attachMsg, setAttachMsg] = useState('');
  const [meta, setMeta] = useState<IFieldMetadata[]>([]);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState<string | undefined>(undefined);
  const [jsonOpen, setJsonOpen] = useState(false);
  const [fieldPanelName, setFieldPanelName] = useState<string | null>(null);

  const fieldsService = useMemo(() => new FieldsService(), []);

  useEffect(() => {
    if (!isOpen) return;
    const norm = buildInitialFieldsAndSteps(value);
    setFields(norm.fields);
    setSteps(norm.steps);
    setRules(value.rules ?? []);
    setHelpJson(JSON.stringify(value.dynamicHelp ?? [], null, 2));
    setManagerColumnFields(value.managerColumnFields ?? []);
    const att = parseAttachmentUiRule(value.rules ?? []);
    setAttachMin(att.minCount);
    setAttachMax(att.maxCount);
    setAttachMsg(att.message);
    setErr(undefined);
    setFieldPanelName(null);
  }, [isOpen, value]);

  useEffect(() => {
    if (!isOpen || !listTitle.trim()) return;
    setLoading(true);
    fieldsService
      .getVisibleFields(listTitle.trim())
      .then((f) => {
        setMeta(f);
        setLoading(false);
      })
      .catch(() => setLoading(false));
  }, [isOpen, listTitle, fieldsService]);

  const fieldOptions: IDropdownOption[] = useMemo(
    () => meta.map((f) => ({ key: f.InternalName, text: `${f.Title} (${f.InternalName})` })),
    [meta]
  );

  const conditionalCards = useMemo(() => parseConditionalCardsFromRules(rules).cards, [rules]);

  const customs = useMemo(() => customRulesOnly(rules), [rules]);

  const setCardsAndRules = useCallback((cards: IConditionalRuleCard[]) => {
    setRules((r) => mergeCardRulesIntoAll(r, cards));
  }, []);

  const addField = (internalName: string): void => {
    if (!internalName) return;
    setSteps((prevSteps) => {
      const st = ensureAtLeastOneStep(prevSteps);
      let already = false;
      for (let s = 0; s < st.length; s++) {
        if (st[s].fieldNames.indexOf(internalName) !== -1) {
          already = true;
          break;
        }
      }
      if (already) return prevSteps;
      const sid = st[0].id;
      const nextSteps = st.map((s, i) =>
        i === 0 ? { ...s, fieldNames: s.fieldNames.concat([internalName]) } : s
      );
      setFields((prev) => {
        const withF = prev.some((f) => f.internalName === internalName)
          ? prev
          : prev.concat([{ internalName, sectionId: sid }]);
        return fieldsAlignedToSteps(withF, nextSteps);
      });
      return nextSteps;
    });
  };

  const removeField = (internalName: string): void => {
    setFields((prev) => prev.filter((f) => f.internalName !== internalName));
    setSteps((prev) =>
      prev.map((s) => ({
        ...s,
        fieldNames: s.fieldNames.filter((n) => n !== internalName),
      }))
    );
  };

  const reorderField = (from: number, to: number): void => {
    setFields((prev) => {
      const next = reorderByIndex(prev, from, to);
      setSteps((st) => resyncStepsOrderFromFields(next, st));
      return next;
    });
  };

  const updateFieldAt = (internalName: string, patch: Partial<IFormFieldConfig>): void => {
    setFields((prev) => prev.map((f) => (f.internalName === internalName ? { ...f, ...patch } : f)));
  };

  const handleStructureFieldDrop = useCallback((toStepIdx: number, insertBefore: number) => {
    return (e: React.DragEvent<HTMLElement>): void => {
      e.preventDefault();
      e.stopPropagation();
      const d = e.dataTransfer.getData('text/plain');
      const poolName = parsePoolDrag(d);
      if (poolName) {
        setSteps((prevSteps) => {
          const nextSteps = insertFieldNameIntoStep(prevSteps, poolName, toStepIdx, insertBefore);
          setFields((prevFields) => {
            let f = prevFields;
            const sid = nextSteps[toStepIdx] ? nextSteps[toStepIdx].id : '';
            let has = false;
            for (let i = 0; i < f.length; i++) {
              if (f[i].internalName === poolName) {
                has = true;
                break;
              }
            }
            if (!has) {
              f = f.concat([{ internalName: poolName, sectionId: sid }]);
            }
            return fieldsAlignedToSteps(f, nextSteps);
          });
          return nextSteps;
        });
        return;
      }
      const fs = parseFieldInStepDrag(d);
      if (fs) {
        setSteps((prevSteps) => {
          const nextSteps = insertFieldNameIntoStep(prevSteps, fs.name, toStepIdx, insertBefore);
          setFields((prevFields) => fieldsAlignedToSteps(prevFields, nextSteps));
          return nextSteps;
        });
      }
    };
  }, []);

  const handleSave = (): void => {
    setErr(undefined);
    let dynamicHelp: IFormManagerConfig['dynamicHelp'];
    try {
      const h = JSON.parse(helpJson || '[]');
      dynamicHelp = Array.isArray(h) && h.length > 0 ? h : undefined;
    } catch {
      setErr('JSON de ajuda dinâmica inválido.');
      return;
    }
    const withRules = mergeAttachmentUiRule(rules, {
      minCount: numOpt(attachMin),
      maxCount: numOpt(attachMax),
      message: attachMsg,
    });
    const sectionsOut = sectionsFromSteps(steps);
    const raw: IFormManagerConfig = {
      sections: sectionsOut,
      fields,
      rules: withRules,
      steps,
      ...(dynamicHelp ? { dynamicHelp } : {}),
      ...(managerColumnFields.length ? { managerColumnFields } : {}),
    };
    const sanitized = sanitizeFormManagerConfig(raw);
    if (!sanitized) {
      setErr('Configuração inválida.');
      return;
    }
    onSave(sanitized);
    onDismiss();
  };

  const addStep = (): void => {
    setSteps((prev) => [...prev, { id: newId('step'), title: 'Nova etapa', fieldNames: [] }]);
  };

  const updateStep = (i: number, patch: Partial<IFormStepConfig>): void => {
    setSteps((prev) => prev.map((s, j) => (j === i ? { ...s, ...patch } : s)));
  };

  const reorderStep = (from: number, to: number): void => {
    setSteps((prev) => {
      const n = reorderByIndex(prev, from, to);
      setFields((flds) => fieldsAlignedToSteps(flds, n));
      return n;
    });
  };

  const removeStep = (i: number): void => {
    setSteps((prev) => {
      if (prev.length <= 1) return prev;
      const removed = prev[i];
      if (!removed) return prev;
      const next = prev.filter((_, j) => j !== i);
      const t0 = next[0];
      if (!t0) return prev;
      const merged = t0.fieldNames.slice();
      for (let k = 0; k < removed.fieldNames.length; k++) {
        const n = removed.fieldNames[k];
        if (merged.indexOf(n) === -1) merged.push(n);
      }
      next[0] = { ...t0, fieldNames: merged };
      setFields((pf) =>
        fieldsAlignedToSteps(
          pf.map((f) =>
            removed.fieldNames.indexOf(f.internalName) !== -1
              ? { ...f, sectionId: next[0].id }
              : f
          ),
          next
        )
      );
      return next;
    });
  };

  const toggleManagerCol = (internalName: string, checked: boolean): void => {
    setManagerColumnFields((prev) => {
      if (checked) {
        if (prev.indexOf(internalName) !== -1) return prev;
        return prev.concat([internalName]);
      }
      return prev.filter((x) => x !== internalName);
    });
  };

  const reorderManagerCol = (from: number, to: number): void => {
    setManagerColumnFields((prev) => reorderByIndex(prev, from, to));
  };

  let fieldPanelConfig: IFormFieldConfig | undefined;
  let fieldPanelMeta: IFieldMetadata | undefined;
  if (fieldPanelName) {
    for (let i = 0; i < fields.length; i++) {
      if (fields[i].internalName === fieldPanelName) {
        fieldPanelConfig = fields[i];
        break;
      }
    }
    for (let j = 0; j < meta.length; j++) {
      if (meta[j].InternalName === fieldPanelName) {
        fieldPanelMeta = meta[j];
        break;
      }
    }
  }

  const previewConfigJson = useMemo(() => {
    const withRules = mergeAttachmentUiRule(rules, {
      minCount: numOpt(attachMin),
      maxCount: numOpt(attachMax),
      message: attachMsg,
    });
    let dynamicHelp: IFormManagerConfig['dynamicHelp'];
    try {
      const h = JSON.parse(helpJson || '[]');
      dynamicHelp = Array.isArray(h) && h.length > 0 ? h : undefined;
    } catch {
      dynamicHelp = undefined;
    }
    const sectionsOut = sectionsFromSteps(steps);
    const raw: IFormManagerConfig = {
      sections: sectionsOut,
      fields,
      rules: withRules,
      steps,
      ...(dynamicHelp ? { dynamicHelp } : {}),
      ...(managerColumnFields.length ? { managerColumnFields } : {}),
    };
    return JSON.stringify(raw, null, 2);
  }, [fields, rules, steps, helpJson, managerColumnFields, attachMin, attachMax, attachMsg]);

  const addConditionalCard = (): void => {
    const card: IConditionalRuleCard = {
      id: newCardId(),
      when: defaultWhenUi(meta),
      effects: [emptyEffect()],
    };
    setCardsAndRules(conditionalCards.concat([card]));
  };

  const patchCard = (index: number, patch: Partial<IConditionalRuleCard>): void => {
    const next = conditionalCards.map((c, i) => (i === index ? { ...c, ...patch } : c));
    setCardsAndRules(next);
  };

  const patchWhen = (index: number, w: Partial<IWhenUi>): void => {
    const c = conditionalCards[index];
    if (!c) return;
    patchCard(index, { when: { ...c.when, ...w } });
  };

  const patchEffect = (cardIndex: number, effIndex: number, patch: Partial<IConditionalEffectUi>): void => {
    const c = conditionalCards[cardIndex];
    if (!c) return;
    const effects = c.effects.map((e, i) => (i === effIndex ? { ...e, ...patch } : e));
    patchCard(cardIndex, { effects });
  };

  const addEffect = (cardIndex: number): void => {
    const c = conditionalCards[cardIndex];
    if (!c) return;
    patchCard(cardIndex, { effects: c.effects.concat([emptyEffect()]) });
  };

  const removeEffect = (cardIndex: number, effIndex: number): void => {
    const c = conditionalCards[cardIndex];
    if (!c) return;
    patchCard(cardIndex, { effects: c.effects.filter((_, i) => i !== effIndex) });
  };

  const duplicateCard = (index: number): void => {
    const c = conditionalCards[index];
    if (!c) return;
    const copy: IConditionalRuleCard = {
      ...c,
      id: newCardId(),
      effects: c.effects.map((e) => ({ ...e })),
    };
    const next = conditionalCards.slice();
    next.splice(index + 1, 0, copy);
    setCardsAndRules(next);
  };

  const removeCard = (index: number): void => {
    setCardsAndRules(conditionalCards.filter((_, i) => i !== index));
  };

  const applyPresetConditional = (preset: 'showWhenEq' | 'choiceRequire'): void => {
    const a = meta[0]?.InternalName ?? 'A';
    const b = meta[1]?.InternalName ?? 'B';
    if (preset === 'showWhenEq') {
      const card = templateConditionalShowWhenEquals(a, '', b);
      card.when.compareValue = '';
      setCardsAndRules(conditionalCards.concat([card]));
    } else {
      setCardsAndRules(conditionalCards.concat([templateFieldRulesChoiceRequiresOther(a, '', b)]));
    }
  };

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.large}
      headerText="Configurar formulário e regras"
      onDismiss={onDismiss}
    >
      {loading && <Spinner label="Campos da lista..." />}
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Link onClick={() => setJsonOpen(true)}>Ver JSON gerado</Link>
      </Stack>
      <Pivot>
        <PivotItem headerText="Estrutura">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Arraste campos para dentro de cada etapa e reordene-os pela alça. O id da etapa é gravado como seção no JSON. Reordene etapas pela alça no cabeçalho.
            </Text>
            <Stack horizontal tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Nova etapa" onClick={addStep} />
            </Stack>
            {steps.map((st, si) => (
              <Stack
                key={st.id}
                styles={{ root: { border: '1px solid #edebe9', padding: 12, borderRadius: 4 } }}
                tokens={{ childrenGap: 8 }}
              >
                <Stack
                  horizontal
                  verticalAlign="end"
                  tokens={{ childrenGap: 8 }}
                  wrap
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    e.dataTransfer.dropEffect = 'move';
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    e.stopPropagation();
                    const from = parseDragIndex(e.dataTransfer.getData('text/plain'), DND_STEP);
                    if (from === undefined || from === si) return;
                    reorderStep(from, si);
                  }}
                >
                  <span
                    draggable
                    title="Arrastar etapa"
                    onDragStart={(e) => {
                      e.dataTransfer.setData('text/plain', dragPayload(DND_STEP, si));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                  >
                    <Icon iconName="GripperBarVertical" />
                  </span>
                  <TextField
                    label={`Título da etapa (${st.id})`}
                    value={st.title}
                    onChange={(_, v) => updateStep(si, { title: v ?? '' })}
                  />
                  <DefaultButton text="Remover etapa" onClick={() => removeStep(si)} />
                </Stack>
                <Stack tokens={{ childrenGap: 6 }} styles={{ root: { marginTop: 4 } }}>
                  {st.fieldNames.map((fname, fIdx) => {
                    let mm: IFieldMetadata | undefined;
                    for (let mi = 0; mi < meta.length; mi++) {
                      if (meta[mi].InternalName === fname) {
                        mm = meta[mi];
                        break;
                      }
                    }
                    return (
                      <Stack
                        key={fname}
                        horizontal
                        verticalAlign="center"
                        tokens={{ childrenGap: 8 }}
                        wrap
                        styles={{
                          root: {
                            padding: '8px 10px',
                            background: '#faf9f8',
                            borderRadius: 4,
                            border: '1px solid #edebe9',
                          },
                        }}
                        onDragOver={(e) => {
                          e.preventDefault();
                          e.stopPropagation();
                          e.dataTransfer.dropEffect = 'move';
                        }}
                        onDrop={handleStructureFieldDrop(si, fIdx)}
                      >
                        <span
                          draggable
                          title="Arrastar campo"
                          onDragStart={(e) => {
                            e.dataTransfer.setData('text/plain', dragPayloadFieldInStep(si, fIdx, fname));
                            e.dataTransfer.effectAllowed = 'move';
                          }}
                          style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                        >
                          <Icon iconName="GripperBarVertical" />
                        </span>
                        <Text styles={{ root: { fontWeight: 600, minWidth: 120 } }}>
                          {mm ? mm.Title : fname}
                        </Text>
                        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                          {fname} · {mm ? mm.MappedType : '—'}
                        </Text>
                        <DefaultButton text="Regras…" onClick={() => setFieldPanelName(fname)} />
                        <DefaultButton text="Remover" onClick={() => removeField(fname)} />
                      </Stack>
                    );
                  })}
                  <Stack
                    styles={{
                      root: {
                        minHeight: 40,
                        padding: 8,
                        borderRadius: 4,
                        border: '1px dashed #c8c6c4',
                        background: '#ffffff',
                      },
                    }}
                    onDragOver={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      e.dataTransfer.dropEffect = 'move';
                    }}
                    onDrop={handleStructureFieldDrop(si, st.fieldNames.length)}
                  >
                    <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                      Soltar aqui para colocar no fim desta etapa
                    </Text>
                  </Stack>
                </Stack>
              </Stack>
            ))}
            <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>Campos fora do formulário</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Arraste um campo para uma etapa acima ou marque para incluir na primeira etapa.
            </Text>
            {meta.map((m) => {
              let inForm = false;
              for (let i = 0; i < fields.length; i++) {
                if (fields[i].internalName === m.InternalName) {
                  inForm = true;
                  break;
                }
              }
              if (inForm) return null;
              return (
                <Stack
                  key={m.InternalName}
                  horizontal
                  verticalAlign="center"
                  tokens={{ childrenGap: 8 }}
                  wrap
                >
                  <span
                    draggable
                    title="Arrastar para uma etapa"
                    onDragStart={(e) => {
                      e.dataTransfer.setData('text/plain', dragPayloadPool(m.InternalName));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                  >
                    <Icon iconName="GripperBarVertical" />
                  </span>
                  <Checkbox
                    label={`${m.Title} (${m.InternalName})`}
                    checked={false}
                    onChange={(_, c) => (c ? addField(m.InternalName) : undefined)}
                  />
                  <Text variant="small" styles={{ root: { minWidth: 80 } }}>{m.MappedType}</Text>
                </Stack>
              );
            })}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Regras rápidas">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Ajustes base por campo. Reordene linhas arrastando a alça. Regras geradas pela UI aparecem no motor com prefixo ui_f_.
            </Text>
            {fields.map((fc, fIdx) => {
                let m: IFieldMetadata | undefined;
                for (let mi = 0; mi < meta.length; mi++) {
                  if (meta[mi].InternalName === fc.internalName) {
                    m = meta[mi];
                    break;
                  }
                }
              const n = countFieldUiRules(fc.internalName, rules);
              const def = fieldRuleStateFromRules(fc.internalName, rules).defaultValue;
              return (
                <Stack
                  key={fc.internalName}
                  horizontal
                  tokens={{ childrenGap: 8 }}
                  verticalAlign="end"
                  wrap
                  onDragOver={(e) => {
                    e.preventDefault();
                    e.dataTransfer.dropEffect = 'move';
                  }}
                  onDrop={(e) => {
                    e.preventDefault();
                    const from = parseDragIndex(e.dataTransfer.getData('text/plain'), DND_FIELD);
                    if (from === undefined || from === fIdx) return;
                    reorderField(from, fIdx);
                  }}
                >
                  <span
                    draggable
                    title="Arrastar para reordenar"
                    onDragStart={(e) => {
                      e.dataTransfer.setData('text/plain', dragPayload(DND_FIELD, fIdx));
                      e.dataTransfer.effectAllowed = 'move';
                    }}
                    style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                  >
                    <Icon iconName="GripperBarVertical" />
                  </span>
                  <Text styles={{ root: { minWidth: 140, fontWeight: 600 } }}>{fc.internalName}</Text>
                  <Checkbox
                    label="Visível"
                    checked={fc.visible !== false}
                    onChange={(_, c) => updateFieldAt(fc.internalName, { visible: !!c })}
                  />
                  <Checkbox
                    label="Obrigatório"
                    checked={fc.required === true}
                    onChange={(_, c) => updateFieldAt(fc.internalName, { required: !!c })}
                  />
                  <Checkbox
                    label="Só leitura"
                    checked={fc.readOnly === true}
                    onChange={(_, c) => updateFieldAt(fc.internalName, { readOnly: !!c })}
                  />
                  <Checkbox
                    label="Desativado"
                    checked={fc.disabled === true}
                    onChange={(_, c) => updateFieldAt(fc.internalName, { disabled: !!c })}
                  />
                  <TextField
                    label="Ajuda"
                    value={fc.helpText ?? ''}
                    onChange={(_, v) => updateFieldAt(fc.internalName, { helpText: v || undefined })}
                  />
                  <TextField
                    label="Padrão (texto/token)"
                    value={def}
                    onChange={(_, v) => {
                      const st = fieldRuleStateFromRules(fc.internalName, rules);
                      st.defaultValue = v ?? '';
                      setRules((r) => mergeFieldRules(r, fc.internalName, buildFieldUiRules(fc.internalName, st)));
                    }}
                  />
                  <DefaultButton
                    text={n ? `${n} regra(s)` : 'Regras…'}
                    onClick={() => setFieldPanelName(fc.internalName)}
                  />
                  {m && <Text variant="small">({m.MappedType})</Text>}
                </Stack>
              );
            })}
            {!fields.length && <Text>Nenhum campo no formulário. Use a aba Estrutura.</Text>}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Regras condicionais">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }}>
              <PrimaryButton text="Nova regra" onClick={addConditionalCard} />
              <DefaultButton
                text="Modelo: mostrar B quando A = valor"
                onClick={() => applyPresetConditional('showWhenEq')}
              />
              <DefaultButton
                text="Modelo: obrigar B quando A = valor"
                onClick={() => applyPresetConditional('choiceRequire')}
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
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>Quando</Text>
                <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                  <Dropdown
                    label="Campo"
                    options={fieldOptions}
                    selectedKey={card.when.field}
                    onChange={(_, o) => o && patchWhen(ci, { field: String(o.key) })}
                  />
                  <Dropdown
                    label="Operador"
                    options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                    selectedKey={card.when.op}
                    onChange={(_, o) => o && patchWhen(ci, { op: o.key as TFormConditionOp })}
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
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>Então</Text>
                {card.effects.map((eff, ei) => (
                  <Stack key={ei} horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
                    <Dropdown
                      label="Efeito"
                      options={CONDITIONAL_EFFECT_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
                      selectedKey={eff.kind}
                      onChange={(_, o) =>
                        o && patchEffect(ci, ei, { kind: o.key as TConditionalEffectKind })
                      }
                    />
                    {eff.kind !== 'message' && (
                      <Dropdown
                        label="Campo alvo"
                        options={[{ key: '', text: '—' }, ...fieldOptions]}
                        selectedKey={eff.targetField ?? ''}
                        onChange={(_, o) =>
                          patchEffect(ci, ei, { targetField: o ? String(o.key) : undefined })
                        }
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
                            o &&
                            patchEffect(ci, ei, { messageVariant: o.key as 'info' | 'warning' | 'error' })
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
                ))}
                <DefaultButton text="Adicionar efeito" onClick={() => addEffect(ci)} />
                <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                  Prévia: {compileConditionalCard(card).length} regra(s) gerada(s)
                </Text>
              </Stack>
            ))}
            {!conditionalCards.length && <Text>Nenhuma regra condicional. Use &quot;Nova regra&quot;.</Text>}
          </Stack>
        </PivotItem>
        <PivotItem headerText="Ajuda dinâmica">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Etapas do formulário são configuradas na aba Estrutura. Aqui: ajuda condicional (JSON avançado).
            </Text>
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Ajuda dinâmica (JSON)</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Array de {'{'} field, when, helpText {'}'} — formato avançado.
            </Text>
            <TextField multiline rows={10} value={helpJson} onChange={(_, v) => setHelpJson(v ?? '')} />
          </Stack>
        </PivotItem>
        <PivotItem headerText="Gestor">
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Colunas da grade gestor. Ordem abaixo.
            </Text>
            {meta.map((m) => (
              <Checkbox
                key={m.InternalName}
                label={`${m.Title} (${m.InternalName})`}
                checked={managerColumnFields.indexOf(m.InternalName) !== -1}
                onChange={(_, c) => toggleManagerCol(m.InternalName, !!c)}
              />
            ))}
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Ordem das colunas selecionadas (arraste pela alça)</Text>
            {managerColumnFields.map((name, mi) => (
              <Stack
                key={name}
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 8 }}
                onDragOver={(e) => {
                  e.preventDefault();
                  e.dataTransfer.dropEffect = 'move';
                }}
                onDrop={(e) => {
                  e.preventDefault();
                  const from = parseDragIndex(e.dataTransfer.getData('text/plain'), DND_MCOL);
                  if (from === undefined || from === mi) return;
                  reorderManagerCol(from, mi);
                }}
              >
                <span
                  draggable
                  title="Arrastar para reordenar"
                  onDragStart={(e) => {
                    e.dataTransfer.setData('text/plain', dragPayload(DND_MCOL, mi));
                    e.dataTransfer.effectAllowed = 'move';
                  }}
                  style={{ cursor: 'grab', display: 'flex', alignItems: 'center', color: '#605e5c' }}
                >
                  <Icon iconName="GripperBarVertical" />
                </span>
                <Text styles={{ root: { minWidth: 160 } }}>{name}</Text>
              </Stack>
            ))}
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Anexos (formulário)</Text>
            <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
              <TextField label="Mín. arquivos" value={attachMin} onChange={(_, v) => setAttachMin(v ?? '')} />
              <TextField label="Máx. arquivos" value={attachMax} onChange={(_, v) => setAttachMax(v ?? '')} />
              <TextField
                label="Mensagem"
                value={attachMsg}
                onChange={(_, v) => setAttachMsg(v ?? '')}
                styles={{ root: { minWidth: 280 } }}
              />
            </Stack>
          </Stack>
        </PivotItem>
      </Pivot>
      {!!customs.length && (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 16 } }}>
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
      <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 24 } }}>
        <PrimaryButton text="Salvar" onClick={handleSave} />
        <DefaultButton
          text="Restaurar padrão (estrutura)"
          onClick={() => {
            const d = getDefaultFormManagerConfig();
            const st = d.steps && d.steps.length ? d.steps : [{ id: 'main', title: 'Geral', fieldNames: [] }];
            setSteps(st.map((x) => ({ ...x, fieldNames: x.fieldNames.slice() })));
            setFields(d.fields.slice());
          }}
        />
        <DefaultButton text="Cancelar" onClick={onDismiss} />
      </Stack>
      <Panel
        isOpen={jsonOpen}
        type={PanelType.medium}
        headerText="JSON gerado (somente leitura)"
        onDismiss={() => setJsonOpen(false)}
      >
        <TextField multiline readOnly rows={22} value={previewConfigJson} />
        <DefaultButton styles={{ root: { marginTop: 12 } }} text="Fechar" onClick={() => setJsonOpen(false)} />
      </Panel>
      {fieldPanelName && fieldPanelConfig && (
        <FormFieldRulesPanel
          isOpen={true}
          internalName={fieldPanelName}
          fieldConfig={fieldPanelConfig}
          meta={fieldPanelMeta}
          rules={rules}
          fieldOptions={fieldOptions}
          onDismiss={() => setFieldPanelName(null)}
          onApply={(nextFc, editor) => {
            setFields((prev) => prev.map((f) => (f.internalName === fieldPanelName ? { ...f, ...nextFc } : f)));
            setRules((r) => mergeFieldRules(r, fieldPanelName, buildFieldUiRules(fieldPanelName, editor)));
          }}
        />
      )}
    </Panel>
  );
};
