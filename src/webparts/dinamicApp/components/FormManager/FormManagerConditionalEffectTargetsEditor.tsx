import * as React from 'react';
import { useState } from 'react';
import { Stack, TextField, Dropdown, IDropdownOption } from '@fluentui/react';
import {
  resolveEffectTargetFields,
  splitEffectTargetsFromCsv,
  type IConditionalEffectUi,
} from '../../core/formManager/formManagerVisualModel';

const ADD_SENTINEL = '__add_none__';

export interface IFormManagerConditionalEffectTargetsEditorProps {
  effect: IConditionalEffectUi;
  fieldOptions: IDropdownOption[];
  onPatch: (patch: Partial<IConditionalEffectUi>) => void;
}

export const FormManagerConditionalEffectTargetsEditor: React.FC<
  IFormManagerConditionalEffectTargetsEditorProps
> = ({ effect, fieldOptions, onPatch }) => {
  const [adderKey, setAdderKey] = useState(ADD_SENTINEL);
  const csv = resolveEffectTargetFields(effect).join(', ');
  const addOpts: IDropdownOption[] = [
    { key: ADD_SENTINEL, text: '— Acrescentar campo —' },
    ...fieldOptions,
  ];
  return (
    <Stack tokens={{ childrenGap: 6 }} styles={{ root: { minWidth: 260, flex: 1 } }}>
      <TextField
        label="Campos alvo (internos, vírgula)"
        description="Um efeito pode aplicar-se a vários campos à vez."
        multiline
        rows={2}
        value={csv}
        onChange={(_, v) => onPatch(splitEffectTargetsFromCsv(v ?? ''))}
      />
      <Dropdown
        label="Acrescentar à lista"
        options={addOpts}
        selectedKey={adderKey}
        onChange={(_, o) => {
          if (!o || o.key === ADD_SENTINEL) return;
          const k = String(o.key);
          const cur = resolveEffectTargetFields(effect);
          if (cur.indexOf(k) !== -1) {
            setAdderKey(ADD_SENTINEL);
            return;
          }
          onPatch(splitEffectTargetsFromCsv([...cur, k].join(', ')));
          setAdderKey(ADD_SENTINEL);
        }}
      />
    </Stack>
  );
};
