import * as React from 'react';
import { Stack, Text, Toggle, Dropdown, IDropdownOption } from '@fluentui/react';
import { IWizardFormState, PAGE_SIZE_OPTIONS } from '../types';

interface IStep4Props {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
}

const pageSizeDropdownOptions: IDropdownOption[] = PAGE_SIZE_OPTIONS.map((n) => ({
  key: n,
  text: `${n} itens por página`,
}));

export const Step4Pagination: React.FC<IStep4Props> = ({ form, onChange }) => {
  const handleToggle = (_: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    onChange({ paginationEnabled: !!checked });
  };

  const handlePageSize = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    if (!opt) return;
    onChange({ pageSize: opt.key as number });
  };

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          Paginação
        </Text>
        <Text variant="medium" styles={{ root: { color: '#605e5c', marginTop: 4, display: 'block' } }}>
          A paginação server-side garante performance mesmo em listas grandes.
        </Text>
      </Stack.Item>

      <Toggle
        label="Habilitar paginação"
        checked={form.paginationEnabled}
        onChange={handleToggle}
        onText="Sim"
        offText="Não"
      />

      {form.paginationEnabled && (
        <Dropdown
          label="Itens por página (padrão)"
          options={pageSizeDropdownOptions}
          selectedKey={form.pageSize}
          onChange={handlePageSize}
          styles={{ root: { maxWidth: 260 } }}
        />
      )}

      {form.paginationEnabled && (
        <Stack
          tokens={{ childrenGap: 6 }}
          styles={{
            root: {
              background: '#f3f9ff',
              border: '1px solid #c7e0f4',
              borderRadius: 6,
              padding: '12px 16px',
            },
          }}
        >
          <Text variant="small" styles={{ root: { color: '#004578', fontWeight: 600 } }}>
            Opções disponíveis para o usuário
          </Text>
          <Text variant="small" styles={{ root: { color: '#004578' } }}>
            {form.pageSizeOptions.join(', ')} itens por página
          </Text>
        </Stack>
      )}
    </Stack>
  );
};
