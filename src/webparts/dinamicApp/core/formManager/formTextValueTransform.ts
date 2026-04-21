import type { IFieldMetadata } from '../../../../services';
import type { IFormFieldConfig, TFormFieldTextValueTransform } from '../config/types/formManager';

export function applyFormFieldTextTransform(s: string, t: TFormFieldTextValueTransform): string {
  if (t === 'uppercase') return s.toLocaleUpperCase('pt-BR');
  if (t === 'lowercase') return s.toLocaleLowerCase('pt-BR');
  return s.replace(
    /\S+/g,
    (w) => w.charAt(0).toLocaleUpperCase('pt-BR') + w.slice(1).toLocaleLowerCase('pt-BR')
  );
}

export function applyTextTransformsToRecordValues(
  values: Record<string, unknown>,
  fieldConfigs: readonly IFormFieldConfig[],
  metaByName: ReadonlyMap<string, IFieldMetadata>
): Record<string, unknown> {
  let changed = false;
  const out = { ...values };
  for (let i = 0; i < fieldConfigs.length; i++) {
    const fc = fieldConfigs[i];
    const t = fc.textValueTransform;
    if (!t) continue;
    const name = fc.internalName;
    const m = metaByName.get(name);
    if (m?.MappedType !== 'text' && m?.MappedType !== 'multiline') continue;
    const v = out[name];
    if (typeof v !== 'string') continue;
    const nv = applyFormFieldTextTransform(v, t);
    if (nv !== v) {
      out[name] = nv;
      changed = true;
    }
  }
  return changed ? out : values;
}
