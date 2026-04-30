import type { IFieldMetadata } from '../../../../services';

export interface IConfirmPromptEditorState {
  text: string;
  bool: boolean;
  dateIso: string | null;
  choiceKey: string;
}

const ELIGIBLE = new Set<IFieldMetadata['MappedType']>([
  'text',
  'multiline',
  'url',
  'number',
  'currency',
  'boolean',
  'datetime',
  'choice',
]);

export function isConfirmPromptEligibleField(m: IFieldMetadata): boolean {
  if (m.ReadOnlyField === true || m.Hidden === true) return false;
  return ELIGIBLE.has(m.MappedType);
}

export function initConfirmPromptEditor(
  meta: IFieldMetadata,
  current: unknown
): IConfirmPromptEditorState {
  const base: IConfirmPromptEditorState = {
    text: '',
    bool: false,
    dateIso: null,
    choiceKey: '',
  };
  switch (meta.MappedType) {
    case 'boolean':
      return { ...base, bool: current === true || current === 1 };
    case 'number':
    case 'currency':
      return {
        ...base,
        text: current !== null && current !== undefined ? String(current) : '',
      };
    case 'datetime': {
      if (current === null || current === undefined || current === '') return base;
      const d = new Date(String(current));
      return { ...base, dateIso: isNaN(d.getTime()) ? null : d.toISOString() };
    }
    case 'choice': {
      const s = current !== null && current !== undefined ? String(current) : '';
      return { ...base, choiceKey: s };
    }
    default:
      return { ...base, text: current !== null && current !== undefined ? String(current) : '' };
  }
}

export function confirmPromptEditorToValue(
  meta: IFieldMetadata,
  ed: IConfirmPromptEditorState
): unknown {
  switch (meta.MappedType) {
    case 'boolean':
      return ed.bool;
    case 'number':
    case 'currency': {
      const t = ed.text.trim();
      if (!t) return null;
      const n = Number(ed.text);
      return isNaN(n) ? null : n;
    }
    case 'datetime':
      return ed.dateIso;
    case 'choice': {
      const k = ed.choiceKey.trim();
      return k.length > 0 ? k : null;
    }
    default:
      return ed.text;
  }
}

export function confirmPromptEditorIsFilled(
  meta: IFieldMetadata,
  ed: IConfirmPromptEditorState
): boolean {
  switch (meta.MappedType) {
    case 'boolean':
      return true;
    case 'number':
    case 'currency':
      return ed.text.trim().length > 0;
    case 'datetime':
      return ed.dateIso !== null && ed.dateIso !== '';
    case 'choice':
      return ed.choiceKey.trim().length > 0;
    default:
      return ed.text.trim().length > 0;
  }
}
