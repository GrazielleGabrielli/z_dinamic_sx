import * as React from 'react';
import { Stack, Text } from '@fluentui/react';
import { IWizardFormState } from '../types';
import { FormStepLayoutPicker } from '../../FormManager/FormStepLayoutUi';

interface IStep3FormStepLayoutProps {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
}

export const Step3FormStepLayout: React.FC<IStep3FormStepLayoutProps> = ({ form, onChange }) => (
  <Stack tokens={{ childrenGap: 20 }}>
    <Stack.Item>
      <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
        Layout das etapas
      </Text>
      <Text variant="medium" styles={{ root: { color: '#605e5c', marginTop: 4, display: 'block' } }}>
        Escolha como as etapas do formulário aparecem quando houver mais de uma. Pode alterar depois nas configurações da webpart.
      </Text>
    </Stack.Item>
    <FormStepLayoutPicker
      value={form.formStepLayout}
      onChange={(id) => onChange({ formStepLayout: id })}
    />
  </Stack>
);
