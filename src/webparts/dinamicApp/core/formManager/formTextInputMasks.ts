import type { FactoryOpts } from 'imask';
import type { TFormFieldTextInputMaskKind } from '../config/types/formManager';

/** Valor persistido = texto formatado (máscara visível). Futuro: opcional `unmaskOnSubmit` no motor. */
export const TEXT_INPUT_MASK_CUSTOM_MAX_LEN = 500;

const MASK_KIND_SET = new Set<TFormFieldTextInputMaskKind>([
  'cpf',
  'telefone',
  'cep',
  'cnpj',
  'custom',
]);

export function isTextInputMaskKind(v: unknown): v is TFormFieldTextInputMaskKind {
  return typeof v === 'string' && MASK_KIND_SET.has(v as TFormFieldTextInputMaskKind);
}

/**
 * Opções IMask para o campo; `null` = sem máscara aplicável (nenhuma ou custom vazio).
 */
export function resolveTextInputMaskOptions(
  kind: TFormFieldTextInputMaskKind | undefined,
  customPattern: string | undefined
): FactoryOpts | null {
  if (!kind) return null;
  if (kind === 'custom') {
    const p = (customPattern ?? '').trim();
    if (!p) return null;
    return { mask: p };
  }
  switch (kind) {
    case 'cpf':
      return { mask: '000.000.000-00' };
    case 'cnpj':
      return { mask: '00.000.000/0000-00' };
    case 'cep':
      return { mask: '00000-000' };
    case 'telefone':
      return {
        mask: [{ mask: '(00) 0000-0000' }, { mask: '(00) 00000-0000' }],
      };
    default:
      return null;
  }
}
