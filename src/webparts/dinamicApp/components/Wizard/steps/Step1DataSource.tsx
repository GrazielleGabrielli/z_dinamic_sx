import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Stack,
  Text,
  ChoiceGroup,
  IChoiceGroupOption,
  Dropdown,
  IDropdownOption,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from '@fluentui/react';
import { ListsService, IListSummary } from '../../../../../services';
import { IWizardFormState } from '../types';

interface IStep1Props {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
}

const sourceKindOptions: IChoiceGroupOption[] = [
  { key: 'list', text: 'Lista' },
  { key: 'library', text: 'Biblioteca de documentos' },
];

export const Step1DataSource: React.FC<IStep1Props> = ({ form, onChange }) => {
  const [allSources, setAllSources] = useState<IListSummary[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);

  useEffect(() => {
    setLoading(true);
    const service = new ListsService();
    service
      .getLists(false)
      .then((data) => {
        setAllSources(data);
        setLoading(false);
      })
      .catch((err: Error) => {
        setError(`Não foi possível carregar as origens: ${err.message}`);
        setLoading(false);
      });
  }, []);

  const filtered = allSources.filter((l) =>
    form.kind === 'library' ? l.IsLibrary : !l.IsLibrary
  );

  const dropdownOptions: IDropdownOption[] = filtered.map((l) => ({
    key: l.Title,
    text: l.Title,
  }));

  const handleKindChange = (
    _: React.FormEvent<HTMLElement | HTMLInputElement> | undefined,
    opt?: IChoiceGroupOption
  ): void => {
    if (!opt) return;
    onChange({ kind: opt.key as 'list' | 'library', title: '' });
  };

  const handleTitleChange = (
    _: React.FormEvent<HTMLDivElement>,
    opt?: IDropdownOption
  ): void => {
    if (!opt) return;
    onChange({ title: opt.key as string });
  };

  return (
    <Stack tokens={{ childrenGap: 20 }}>
      <Stack.Item>
        <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
          Fonte de dados
        </Text>
        <Text variant="medium" styles={{ root: { color: '#605e5c', marginTop: 4, display: 'block' } }}>
          Selecione o tipo de origem e a lista ou biblioteca que será usada.
        </Text>
      </Stack.Item>

      <ChoiceGroup
        label="Tipo de origem"
        options={sourceKindOptions}
        selectedKey={form.kind}
        onChange={handleKindChange}
      />

      {loading && <Spinner size={SpinnerSize.medium} label="Carregando origens disponíveis..." />}

      {error !== undefined && (
        <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
          {error}
        </MessageBar>
      )}

      {!loading && error === undefined && (
        <Dropdown
          label={form.kind === 'library' ? 'Biblioteca' : 'Lista'}
          placeholder={`Selecione uma ${form.kind === 'library' ? 'biblioteca' : 'lista'}`}
          options={dropdownOptions}
          selectedKey={form.title || undefined}
          onChange={handleTitleChange}
          disabled={dropdownOptions.length === 0}
        />
      )}

      {!loading && error === undefined && dropdownOptions.length === 0 && (
        <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
          Nenhuma {form.kind === 'library' ? 'biblioteca' : 'lista'} encontrada neste site.
        </Text>
      )}
    </Stack>
  );
};
