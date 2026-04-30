import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import {
  Stack,
  Text,
  Toggle,
  Dropdown,
  IDropdownOption,
  Checkbox,
  TextField,
  DefaultButton,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
} from '@fluentui/react';
import {
  GroupsService,
  UsersService,
  WebsService,
  IWebSummary,
  filterSiteGroupsByNameQuery,
} from '../../../../services';
import type { IGroupDetails } from '../../../../services/groups/types';
import type { IListViewModeAccessConfig } from '../../core/config/types';
import { normWebPath } from '../../core/listView/viewModeAccess';

const PAGE_KEY = '__page__';
const LIST_KEY = '__list__';

function accessSummary(a: IListViewModeAccessConfig | undefined): string {
  if (!a) return '';
  const g = a.allowedGroupIds?.length ?? 0;
  const u = a.allowedUserIds?.length ?? 0;
  if (g === 0 && u === 0) return 'Restrito — sem grupos nem pessoas';
  const parts: string[] = [];
  if (g) parts.push(`${g} grupo(s)`);
  if (u) parts.push(`${u} pessoa(s)`);
  return `Visível: ${parts.join(' · ')}`;
}

export interface IViewModeAccessSectionProps {
  value: IListViewModeAccessConfig | undefined;
  onChange: (next: IListViewModeAccessConfig | undefined) => void;
  pageWebServerRelativeUrl: string;
  listWebServerRelativeUrl?: string;
  disabled?: boolean;
}

export const ViewModeAccessSection: React.FC<IViewModeAccessSectionProps> = ({
  value,
  onChange,
  pageWebServerRelativeUrl,
  listWebServerRelativeUrl,
  disabled = false,
}) => {
  const pageNorm = useMemo(() => normWebPath(pageWebServerRelativeUrl), [pageWebServerRelativeUrl]);
  const listNorm = useMemo(
    () => (listWebServerRelativeUrl?.trim() ? normWebPath(listWebServerRelativeUrl) : ''),
    [listWebServerRelativeUrl]
  );

  const [subsites, setSubsites] = useState<IWebSummary[]>([]);
  const [groups, setGroups] = useState<IGroupDetails[]>([]);
  const [groupsLoading, setGroupsLoading] = useState(false);
  const [userSearch, setUserSearch] = useState('');
  const [groupListNameFilter, setGroupListNameFilter] = useState('');
  const [userSearchLoading, setUserSearchLoading] = useState(false);
  const [userPicks, setUserPicks] = useState<{ id: number; label: string }[]>([]);

  const selectedWebKey = useMemo((): string => {
    if (!value?.webServerRelativeUrl?.trim()) return PAGE_KEY;
    const v = normWebPath(value.webServerRelativeUrl);
    if (v === pageNorm) return PAGE_KEY;
    if (listNorm && v === listNorm) return LIST_KEY;
    return v;
  }, [value?.webServerRelativeUrl, pageNorm, listNorm]);

  const groupsFilteredForList = useMemo(
    () => filterSiteGroupsByNameQuery(groups, groupListNameFilter),
    [groups, groupListNameFilter]
  );

  const groupIdSet = useMemo(() => new Set(value?.allowedGroupIds ?? []), [value?.allowedGroupIds]);
  const userIdSet = useMemo(() => new Set(value?.allowedUserIds ?? []), [value?.allowedUserIds]);

  const restrictOn = value !== undefined && value !== null;

  useEffect(() => {
    const ws = new WebsService();
    ws
      .getDirectSubsites()
      .then(setSubsites)
      .catch(() => setSubsites([]));
  }, [pageWebServerRelativeUrl]);

  const siteOptions: IDropdownOption[] = useMemo(() => {
    const opts: IDropdownOption[] = [
      { key: PAGE_KEY, text: 'Site desta página' },
    ];
    if (listNorm && listNorm !== pageNorm) {
      opts.push({ key: LIST_KEY, text: 'Site da fonte de dados (lista)' });
    }
    for (let i = 0; i < subsites.length; i++) {
      const s = subsites[i];
      const p = normWebPath(s.ServerRelativeUrl);
      opts.push({ key: p, text: `${s.Title} — ${p}` });
    }
    return opts;
  }, [pageNorm, listNorm, subsites]);

  const loadGroups = useCallback(
    (webKey: string): void => {
      let webArg: string | undefined;
      if (webKey === PAGE_KEY) webArg = undefined;
      else if (webKey === LIST_KEY) webArg = listNorm || undefined;
      else webArg = webKey;
      setGroupsLoading(true);
      const gs = new GroupsService();
      gs
        .getSiteGroups(webArg)
        .then((g) => {
          setGroups(g);
          setGroupsLoading(false);
        })
        .catch(() => {
          setGroups([]);
          setGroupsLoading(false);
        });
    },
    [listNorm]
  );

  useEffect(() => {
    loadGroups(selectedWebKey);
  }, [selectedWebKey, loadGroups]);

  useEffect(() => {
    if (!value?.allowedUserIds?.length) {
      setUserPicks([]);
      return;
    }
    const us = new UsersService();
    let cancelled = false;
    void Promise.all(
      value.allowedUserIds.map((id) =>
        us.getUserById(id).then(
          (u) => ({ id, label: `${u.Title} (${id})` }),
          () => ({ id, label: `#${id}` })
        )
      )
    ).then((rows) => {
      if (!cancelled) setUserPicks(rows);
    });
    return (): void => {
      cancelled = true;
    };
  }, [value?.allowedUserIds]);

  const pushChange = useCallback(
    (partial: Partial<IListViewModeAccessConfig>): void => {
      const hadRestriction = value !== undefined && value !== null;
      const base: IListViewModeAccessConfig = {
        allowedGroupIds: value?.allowedGroupIds?.slice(),
        allowedUserIds: value?.allowedUserIds?.slice(),
        webServerRelativeUrl: value?.webServerRelativeUrl,
        ...partial,
      };
      const g = base.allowedGroupIds?.filter((x) => isFinite(x) && x > 0) ?? [];
      const u = base.allowedUserIds?.filter((x) => isFinite(x) && x > 0) ?? [];
      const web = base.webServerRelativeUrl?.trim();
      const next: IListViewModeAccessConfig = {};
      if (g.length) next.allowedGroupIds = g;
      if (u.length) next.allowedUserIds = u;
      if (web) next.webServerRelativeUrl = web;
      if (!hadRestriction && g.length === 0 && u.length === 0 && !web) {
        onChange(undefined);
        return;
      }
      onChange(Object.keys(next).length ? next : {});
    },
    [value, onChange]
  );

  const handleRestrictToggle = (_: React.MouseEvent, checked?: boolean): void => {
    if (!checked) {
      onChange(undefined);
      setUserSearch('');
      return;
    }
    onChange({});
  };

  const handleSiteChange = (_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption): void => {
    if (!opt) return;
    const k = String(opt.key);
    let webOut: string | undefined;
    if (k === PAGE_KEY) webOut = undefined;
    else if (k === LIST_KEY) webOut = listNorm || undefined;
    else webOut = k;
    pushChange({
      webServerRelativeUrl: webOut,
      allowedGroupIds: [],
    });
  };

  const toggleGroup = (id: number, checked: boolean): void => {
    const next = new Set(groupIdSet);
    if (checked) next.add(id);
    else next.delete(id);
    pushChange({ allowedGroupIds: Array.from(next) });
  };

  const addUserFromSearch = (): void => {
    const q = userSearch.trim();
    if (!q.length) return;
    setUserSearchLoading(true);
    const us = new UsersService();
    const webForEnsure =
      selectedWebKey === PAGE_KEY ? undefined : selectedWebKey === LIST_KEY ? listNorm || undefined : selectedWebKey;
    us
      .searchUsers(q, 12)
      .then(async (hits) => {
        if (!hits.length) {
          const id = await us.ensureUserAndGetId(q, webForEnsure);
          setUserSearchLoading(false);
          if (id === undefined) return;
          const merged = new Set(userIdSet);
          merged.add(id);
          pushChange({ allowedUserIds: Array.from(merged) });
          setUserSearch('');
          return;
        }
        const first = hits[0];
        const id = await us.ensureUserAndGetId(first.Key, webForEnsure);
        setUserSearchLoading(false);
        if (id === undefined) return;
        const merged = new Set(userIdSet);
        merged.add(id);
        pushChange({ allowedUserIds: Array.from(merged) });
        setUserSearch('');
      })
      .catch(() => setUserSearchLoading(false));
  };

  const removeUser = (id: number): void => {
    const merged = new Set(userIdSet);
    merged.delete(id);
    pushChange({ allowedUserIds: Array.from(merged) });
  };

  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Toggle
        label="Restringir visibilidade deste modo"
        checked={restrictOn}
        onChange={handleRestrictToggle}
        disabled={disabled}
      />
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Só utilizadores nos grupos ou lista abaixo vêem este modo na barra (modo OU entre grupos e pessoas).
      </Text>
      {restrictOn && (
        <>
          <Dropdown
            label="Site dos grupos"
            options={siteOptions}
            selectedKey={selectedWebKey}
            onChange={handleSiteChange}
            disabled={disabled}
          />
          {groupsLoading ? (
            <Spinner size={SpinnerSize.small} label="A carregar grupos..." />
          ) : (
            <Stack tokens={{ childrenGap: 4 }}>
              <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
                Grupos permitidos
              </Text>
              <TextField
                placeholder="Filtrar grupos por nome"
                value={groupListNameFilter}
                onChange={(_: unknown, v?: string) => setGroupListNameFilter(v ?? '')}
                disabled={disabled}
                styles={{ root: { maxWidth: 420 } }}
              />
              <div style={{ maxHeight: 180, overflowY: 'auto', border: '1px solid #edebe9', borderRadius: 6, padding: 8 }}>
                {groupsFilteredForList.map((g) => {
                  const id = typeof g.Id === 'number' ? g.Id : Number(g.Id);
                  return (
                    <Checkbox
                      key={id}
                      label={g.Title ?? String(id)}
                      checked={groupIdSet.has(id)}
                      disabled={disabled}
                      onChange={(_, c) => toggleGroup(id, c === true)}
                    />
                  );
                })}
              </div>
                {groups.length > 0 && !groupsFilteredForList.length && groupListNameFilter.trim() ? (
                  <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                    Nenhum grupo corresponde ao filtro.
                  </Text>
                ) : null}
                {groups.length === 0 && (
                  <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                    Sem grupos neste site ou sem permissão de leitura.
                  </Text>
                )}
            </Stack>
          )}
          <Stack tokens={{ childrenGap: 6 }}>
            <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
              Pessoas permitidas
            </Text>
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end" wrap>
              <TextField
                placeholder="Nome ou e-mail para pesquisar"
                value={userSearch}
                onChange={(_: unknown, v?: string) => setUserSearch(v ?? '')}
                disabled={disabled || userSearchLoading}
                styles={{ root: { flex: '1 1 220px', minWidth: 160 } }}
              />
              <DefaultButton
                text="Adicionar"
                disabled={disabled || userSearchLoading || !userSearch.trim()}
                onClick={addUserFromSearch}
              />
            </Stack>
            {userPicks.map((p) => (
              <Stack horizontal verticalAlign="center" key={p.id} tokens={{ childrenGap: 8 }}>
                <Text variant="small">{p.label}</Text>
                <DefaultButton text="Remover" onClick={() => removeUser(p.id)} disabled={disabled} />
              </Stack>
            ))}
          </Stack>
          {restrictOn &&
            (value?.allowedGroupIds?.length ?? 0) === 0 &&
            (value?.allowedUserIds?.length ?? 0) === 0 && (
              <MessageBar messageBarType={MessageBarType.warning}>
                Selecione pelo menos um grupo ou uma pessoa; caso contrário o modo não aparece a ninguém.
              </MessageBar>
            )}
        </>
      )}
    </Stack>
  );
};

export { accessSummary };
