import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
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
import { ListsService, IListSummary, WebsService, IWebSummary } from '../../../../../services';
import { IWizardFormState } from '../types';

interface IStep1Props {
  form: IWizardFormState;
  onChange: (partial: Partial<IWizardFormState>) => void;
  currentWebServerRelativeUrl: string;
}

const sourceKindOptions: IChoiceGroupOption[] = [
  { key: 'list', text: 'Lista' },
  { key: 'library', text: 'Biblioteca de documentos' },
];

function normPath(s: string): string {
  const t = (s || '').trim().replace(/\/+$/, '') || '/';
  return t.startsWith('/') ? t : `/${t}`;
}

export const Step1DataSource: React.FC<IStep1Props> = ({ form, onChange, currentWebServerRelativeUrl }) => {
  const [allSources, setAllSources] = useState<IListSummary[]>([]);
  const [sites, setSites] = useState<IWebSummary[]>([]);
  const [currentWebTitle, setCurrentWebTitle] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);

  const selectedSitePath = useMemo(() => {
    const fromForm = (form.dataSourceWebServerRelativeUrl ?? '').trim();
    if (fromForm) return normPath(fromForm);
    return normPath(currentWebServerRelativeUrl);
  }, [form.dataSourceWebServerRelativeUrl, currentWebServerRelativeUrl]);

  const loadSites = useCallback((): void => {
    const ws = new WebsService();
    ws
      .getCurrentWeb()
      .then((cur) => {
        setCurrentWebTitle(cur.Title || '');
        return ws.getDirectSubsites();
      })
      .then((subs) => setSites(subs))
      .catch(() => setSites([]));
  }, []);

  useEffect(() => {
    loadSites();
  }, [loadSites]);

  useEffect(() => {
    setLoading(true);
    setError(undefined);
    const service = new ListsService();
    const webArg =
      normPath(selectedSitePath) === normPath(currentWebServerRelativeUrl)
        ? undefined
        : selectedSitePath;
    service
      .getLists(false, webArg)
      .then((data) => {
        setAllSources(data);
        setLoading(false);
      })
      .catch((err: Error) => {
        setError(`Não foi possível carregar as origens: ${err.message}`);
        setLoading(false);
      });
  }, [form.kind, selectedSitePath, currentWebServerRelativeUrl]);

  const filtered = allSources.filter((l) =>
    form.kind === 'library' ? l.IsLibrary : !l.IsLibrary
  );

  const siteOptions: IDropdownOption[] = useMemo(() => {
    const curPath = normPath(currentWebServerRelativeUrl);
    const opts: IDropdownOption[] = [
      {
        key: curPath,
        text: currentWebTitle ? `Este site (${currentWebTitle})` : 'Este site',
      },
    ];
    const seen = new Set<string>([curPath]);
    for (let i = 0; i < sites.length; i++) {
      const p = normPath(sites[i].ServerRelativeUrl);
      if (seen.has(p)) continue;
      seen.add(p);
      opts.push({
        key: p,
        text: `${sites[i].Title || p} — ${p}`,
      });
    }
    return opts;
  }, [sites, currentWebServerRelativeUrl, currentWebTitle]);

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

  const handleSiteChange = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    if (!opt) return;
    const key = String(opt.key);
    const curNorm = normPath(currentWebServerRelativeUrl);
    const nextNorm = normPath(key);
    if (nextNorm === curNorm) {
      onChange({ dataSourceWebServerRelativeUrl: undefined, title: '' });
    } else {
      onChange({ dataSourceWebServerRelativeUrl: nextNorm, title: '' });
    }
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
          Selecione o site (ou subsite direto), o tipo de origem e a lista ou biblioteca.
        </Text>
      </Stack.Item>

      <Dropdown
        label="Site"
        options={siteOptions}
        selectedKey={selectedSitePath}
        onChange={handleSiteChange}
        disabled={siteOptions.length === 0}
      />

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
