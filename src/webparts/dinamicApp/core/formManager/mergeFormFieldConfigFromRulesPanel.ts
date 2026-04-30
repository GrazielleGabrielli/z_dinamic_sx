import type { IFormFieldConfig } from '../config/types/formManager';

const KEYS_RESET_WHEN_ABSENT_FROM_PANEL: (keyof IFormFieldConfig)[] = [
  'textInputMaskKind',
  'textInputMaskCustomPattern',
  'textValueTransform',
  'visible',
  'textConditionalVisibility',
];

export function mergeFormFieldConfigFromRulesPanel(
  previous: IFormFieldConfig,
  next: IFormFieldConfig
): IFormFieldConfig {
  const merged = { ...previous, ...next } as IFormFieldConfig;
  const nextRec = next as unknown as Record<string, unknown>;
  const out = merged as unknown as Record<string, unknown>;
  for (const k of KEYS_RESET_WHEN_ABSENT_FROM_PANEL) {
    if (!(k in nextRec)) {
      delete out[k as string];
    }
  }
  return out as unknown as IFormFieldConfig;
}
