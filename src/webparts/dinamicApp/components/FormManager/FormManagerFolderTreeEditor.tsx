import * as React from 'react';
import { useState, useMemo, useEffect } from 'react';
import {
  Stack,
  Text,
  TextField,
  IconButton,
  DefaultButton,
  Dropdown,
  Checkbox,
  MessageBar,
  MessageBarType,
  Spinner,
  type IDropdownOption,
} from '@fluentui/react';
import { filterSiteGroupsByNameQuery, type IGroupDetails } from '../../../../services';
import type {
  IAttachmentLibraryFolderTreeNode,
  TFormConditionOp,
  TFormManagerFormMode,
} from '../../core/config/types/formManager';
import {
  folderTemplateLiteralInvalidReason,
  sanitizeFolderNameTemplatePreservingPlaceholders,
} from '../../core/formManager/attachmentFolderNameTemplate';
import {
  MAX_ATTACHMENT_FOLDER_TREE_NODES,
  addChild,
  addRootSibling,
  addSiblingAfter,
  countNodesInTree,
  removeNodeById,
  patchNodeName,
  patchNodeShowUploaderStepIds,
  setUploadTargetById,
  updateAttachmentFolderNode,
  flattenFolderTreeNodes,
} from '../../core/formManager/attachmentFolderTree';
import {
  CONDITION_OP_OPTIONS,
  whenNodeToUi,
  whenUiToNode,
  type IWhenUi,
} from '../../core/formManager/formManagerVisualModel';
import { FormManagerCollapseSection } from './FormManagerComponentsTab';

export interface IFolderVisibilityEditorProps {
  fieldOptions: IDropdownOption[];
  defaultConditionFieldName: string;
  siteGroupsSorted: IGroupDetails[];
  siteGroups: IGroupDetails[];
  siteGroupsLoading: boolean;
  siteGroupsErr?: string;
  onReloadSiteGroups: () => void;
}

export interface IFormManagerFolderTreeEditorProps {
  nodes: IAttachmentLibraryFolderTreeNode[];
  onChange: (next: IAttachmentLibraryFolderTreeNode[]) => void;
  disabled?: boolean;
  folderStepOptions: { id: string; title: string }[];
  showFolderStepPicker: boolean;
  folderVisibilityEditor?: IFolderVisibilityEditorProps;
}

const ALL_MODES: TFormManagerFormMode[] = ['create', 'edit', 'view'];

const MODE_LABELS: { key: TFormManagerFormMode; label: string }[] = [
  { key: 'create', label: 'Criar' },
  { key: 'edit', label: 'Editar' },
  { key: 'view', label: 'Ver' },
];

function normFolderRuleGroupTitle(s: string): string {
  return s.trim().toLowerCase();
}

function defaultFolderWhenUi(fieldName: string): IWhenUi {
  return { field: fieldName, op: 'eq', compareKind: 'literal', compareValue: '' };
}

function folderPathLabel(
  tree: IAttachmentLibraryFolderTreeNode[],
  targetId: string
): string {
  function walk(ns: IAttachmentLibraryFolderTreeNode[], acc: string[]): string | undefined {
    for (let i = 0; i < ns.length; i++) {
      const n = ns[i];
      const label = n.nameTemplate?.trim() || '(sem nome)';
      const next = acc.concat([label]);
      if (n.id === targetId) return next.join(' / ');
      if (n.children?.length) {
        const d = walk(n.children, next);
        if (d) return d;
      }
    }
    return undefined;
  }
  return walk(tree, []) ?? targetId;
}

function FolderStepSelect(props: {
  node: IAttachmentLibraryFolderTreeNode;
  disabled: boolean;
  folderStepOptions: { id: string; title: string }[];
  onPatchStepIds: (id: string, stepIds: string[]) => void;
}): JSX.Element {
  const { node, disabled, folderStepOptions, onPatchStepIds } = props;
  const selectedId = node.showUploaderInStepIds?.[0];
  const dropdownOptions: IDropdownOption[] = [
    { key: '', text: '—' },
    ...folderStepOptions.map((st) => ({ key: st.id, text: st.title })),
  ];
  return (
    <Dropdown
      key={`step-dd-${node.id}-${selectedId ?? ''}`}
      placeholder="Etapa…"
      options={dropdownOptions}
      selectedKey={selectedId ?? ''}
      disabled={disabled}
      onChange={(_, o) => {
        if (!o || disabled) return;
        const key = String(o.key);
        onPatchStepIds(node.id, key ? [key] : []);
      }}
      styles={{ root: { width: '100%', maxWidth: '100%', minWidth: 160 } }}
    />
  );
}

function FolderVisibilityRules(props: {
  node: IAttachmentLibraryFolderTreeNode;
  treeNodes: IAttachmentLibraryFolderTreeNode[];
  disabled: boolean;
  onTreeChange: (next: IAttachmentLibraryFolderTreeNode[]) => void;
  editor: IFolderVisibilityEditorProps;
}): JSX.Element {
  const { node, treeNodes, disabled, onTreeChange, editor } = props;
  const {
    fieldOptions,
    defaultConditionFieldName,
    siteGroupsSorted,
    siteGroups,
    siteGroupsLoading,
    siteGroupsErr,
    onReloadSiteGroups,
  } = editor;

  const [folderGroupNameFilter, setFolderGroupNameFilter] = useState('');
  const siteGroupsSortedFiltered = useMemo(
    () => filterSiteGroupsByNameQuery(siteGroupsSorted, folderGroupNameFilter),
    [siteGroupsSorted, folderGroupNameFilter]
  );

  const patchNode = (updater: (n: IAttachmentLibraryFolderTreeNode) => IAttachmentLibraryFolderTreeNode): void => {
    onTreeChange(updateAttachmentFolderNode(treeNodes, node.id, updater));
  };

  const leafWhen = node.showUploaderWhen ? whenNodeToUi(node.showUploaderWhen) : undefined;
  const compositeWhen = node.showUploaderWhen && !leafWhen;

  const modeRowChecked = (m: TFormManagerFormMode): boolean => {
    const cur = node.showUploaderModes;
    if (cur === undefined) return true;
    return cur.indexOf(m) !== -1;
  };

  const toggleMode = (m: TFormManagerFormMode, checked: boolean): void => {
    let next = (node.showUploaderModes?.length ? node.showUploaderModes.slice() : ALL_MODES.slice()) as TFormManagerFormMode[];
    if (checked) {
      if (next.indexOf(m) === -1) next.push(m);
    } else {
      next = next.filter((x) => x !== m);
    }
    patchNode((n) => {
      if (next.length === ALL_MODES.length) {
        const { showUploaderModes: _d, ...rest } = n;
        return rest;
      }
      return { ...n, showUploaderModes: next };
    });
  };

  const patchWhenUi = (partial: Partial<IWhenUi>): void => {
    const baseLeaf = node.showUploaderWhen ? whenNodeToUi(node.showUploaderWhen) : undefined;
    const base: IWhenUi = baseLeaf ?? defaultFolderWhenUi(defaultConditionFieldName);
    const merged: IWhenUi = { ...base, ...partial };
    patchNode((n) => ({ ...n, showUploaderWhen: whenUiToNode(merged) }));
  };

  return (
    <Stack tokens={{ childrenGap: 8 }} styles={{ root: { width: '100%', maxWidth: '100%' } }}>
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
        <Text variant="small" styles={{ root: { color: '#605e5c', minWidth: 44 } }}>
          Modos
        </Text>
        <Text variant="tiny" styles={{ root: { color: '#a19f9d' } }} title="Sem seleção = todos os modos.">
          (todos se vazio)
        </Text>
        {MODE_LABELS.map((x) => (
          <Checkbox
            key={x.key}
            label={x.label}
            checked={modeRowChecked(x.key)}
            disabled={disabled}
            onChange={(_, c) => toggleMode(x.key, !!c)}
          />
        ))}
      </Stack>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Grupos
        <span style={{ fontWeight: 400, color: '#a19f9d' }}> · vazio = qualquer utilizador</span>
      </Text>
      <TextField
        placeholder="Filtrar grupos por nome"
        value={folderGroupNameFilter}
        onChange={(_: unknown, v?: string) => setFolderGroupNameFilter(v ?? '')}
        styles={{ root: { maxWidth: 420 } }}
      />
      {siteGroupsLoading && <Spinner label="Grupos…" />}
      {siteGroupsErr && (
        <Stack tokens={{ childrenGap: 6 }}>
          <MessageBar messageBarType={MessageBarType.warning}>{siteGroupsErr}</MessageBar>
          <DefaultButton text="Repetir" onClick={onReloadSiteGroups} />
        </Stack>
      )}
      {!siteGroupsLoading ? (
        <Stack
          tokens={{ childrenGap: 4 }}
          styles={{
            root: {
              maxHeight: 140,
              overflowY: 'auto',
              border: '1px solid #edebe9',
              borderRadius: 4,
              padding: 6,
            },
          }}
        >
          {(node.showUploaderGroupTitles ?? [])
            .filter((t) => !siteGroups.some((g) => normFolderRuleGroupTitle(g.Title) === normFolderRuleGroupTitle(t)))
            .filter((t) => {
              const q = folderGroupNameFilter.trim().toLowerCase();
              return !q || t.toLowerCase().includes(q);
            })
            .map((t, oi) => (
              <Checkbox
                key={`orphan-grp-${node.id}-${oi}-${t}`}
                label={`${t} (ref.)`}
                checked
                disabled={disabled}
                onChange={(_, c) => {
                  if (c) return;
                  const cur = node.showUploaderGroupTitles ?? [];
                  const n = normFolderRuleGroupTitle(t);
                  const next = cur.filter((x) => normFolderRuleGroupTitle(x) !== n);
                  patchNode((no) => {
                    if (!next.length) {
                      const { showUploaderGroupTitles: _g, ...rest } = no;
                      return rest;
                    }
                    return { ...no, showUploaderGroupTitles: next };
                  });
                }}
              />
            ))}
          {siteGroupsSortedFiltered.map((g) => {
            const cur = node.showUploaderGroupTitles ?? [];
            const n = normFolderRuleGroupTitle(g.Title);
            const checked = cur.some((x) => normFolderRuleGroupTitle(x) === n);
            return (
              <Checkbox
                key={g.Id}
                label={g.Title}
                title={g.Description || undefined}
                checked={checked}
                disabled={disabled}
                onChange={(_, c) => {
                  let next: string[];
                  if (c) {
                    next = checked ? cur : cur.concat([g.Title]);
                  } else {
                    next = cur.filter((x) => normFolderRuleGroupTitle(x) !== n);
                  }
                  patchNode((no) => {
                    if (!next.length) {
                      const { showUploaderGroupTitles: _g, ...rest } = no;
                      return rest;
                    }
                    return { ...no, showUploaderGroupTitles: next };
                  });
                }}
              />
            );
          })}
          {siteGroupsSorted.length > 0 &&
            !siteGroupsSortedFiltered.length &&
            folderGroupNameFilter.trim() && (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Nenhum grupo corresponde ao filtro.
              </Text>
            )}
          {!siteGroupsSorted.length && !(node.showUploaderGroupTitles ?? []).length && (
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Nenhum grupo no site.
            </Text>
          )}
        </Stack>
      ) : null}
      <Checkbox
        label="Condição num campo"
        title="Só mostra o input quando a condição for verdadeira."
        checked={!!node.showUploaderWhen}
        disabled={disabled}
        onChange={(_, c) => {
          if (c) {
            patchNode((n) => ({
              ...n,
              showUploaderWhen: whenUiToNode(defaultFolderWhenUi(defaultConditionFieldName)),
            }));
          } else {
            patchNode((n) => {
              const { showUploaderWhen: _w, ...rest } = n;
              return rest;
            });
          }
        }}
      />
      {compositeWhen && (
        <MessageBar messageBarType={MessageBarType.warning}>
          Condição composta — editar no JSON ou desmarcar a caixa.
        </MessageBar>
      )}
      {node.showUploaderWhen && leafWhen && (
        <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
          <Dropdown
            label="Campo"
            options={fieldOptions}
            selectedKey={leafWhen.field}
            disabled={disabled}
            onChange={(_, o) => o && patchWhenUi({ field: String(o.key) })}
            styles={{ dropdown: { width: 160 } }}
          />
          <Dropdown
            label="Operador"
            options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
            selectedKey={leafWhen.op}
            disabled={disabled}
            onChange={(_, o) => o && patchWhenUi({ op: o.key as TFormConditionOp })}
            styles={{ dropdown: { width: 150 } }}
          />
          <Dropdown
            label="Comparar"
            options={[
              { key: 'literal', text: 'Texto fixo' },
              { key: 'field', text: 'Campo' },
              { key: 'token', text: 'Token' },
            ]}
            selectedKey={leafWhen.compareKind}
            disabled={disabled}
            onChange={(_, o) => o && patchWhenUi({ compareKind: o.key as IWhenUi['compareKind'] })}
            styles={{ dropdown: { width: 112 } }}
          />
          <TextField
            label="Valor"
            value={leafWhen.compareValue}
            disabled={
              disabled ||
              leafWhen.op === 'isEmpty' ||
              leafWhen.op === 'isFilled' ||
              leafWhen.op === 'isTrue' ||
              leafWhen.op === 'isFalse'
            }
            onChange={(_, v) => patchWhenUi({ compareValue: v ?? '' })}
            styles={{ fieldGroup: { minWidth: 120 } }}
          />
        </Stack>
      )}
    </Stack>
  );
}

function FolderStructureRow(props: {
  node: IAttachmentLibraryFolderTreeNode;
  depth: number;
  disabled: boolean;
  onPatchName: (id: string, v: string) => void;
  onAddChild: (id: string) => void;
  onAddSibling: (id: string) => void;
  onRemove: (id: string) => void;
  onSetTarget: (id: string) => void;
  renderChildren: (nodes: IAttachmentLibraryFolderTreeNode[] | undefined, d: number) => React.ReactNode;
}): JSX.Element {
  const { node, depth, disabled, onPatchName, onAddChild, onAddSibling, onRemove, onSetTarget, renderChildren } = props;
  const pad = 12 + depth * 18;
  return (
    <Stack key={node.id} tokens={{ childrenGap: 4 }}>
      <Stack
        horizontal
        verticalAlign="center"
        tokens={{ childrenGap: 6 }}
        styles={{ root: { paddingLeft: pad, flexWrap: 'wrap' } }}
      >
        <IconButton
          iconProps={{ iconName: node.uploadTarget ? 'RadioBtnOn' : 'RadioBtnOff' }}
          onClick={() => onSetTarget(node.id)}
          disabled={disabled}
          title="Destino do upload"
          ariaLabel="Destino do upload"
        />
        <Stack styles={{ root: { flex: '1 1 220px', minWidth: 160, maxWidth: '100%' } }}>
          <TextField
            value={node.nameTemplate}
            onChange={(_, v) => onPatchName(node.id, v ?? '')}
            onBlur={() => {
              const s = sanitizeFolderNameTemplatePreservingPlaceholders(node.nameTemplate);
              if (s !== node.nameTemplate) {
                onPatchName(node.id, s);
              }
            }}
            placeholder="Nome da pasta ou {{Title}}"
            disabled={disabled}
            errorMessage={folderTemplateLiteralInvalidReason(node.nameTemplate)}
            styles={{ root: { width: '100%' } }}
          />
        </Stack>
        <IconButton
          iconProps={{ iconName: 'Add' }}
          title="Subpasta (filho)"
          ariaLabel="Adicionar subpasta"
          disabled={disabled}
          onClick={() => onAddChild(node.id)}
        />
        <IconButton
          iconProps={{ iconName: 'RowInsert' }}
          title="Outra pasta ao mesmo nível (irmã). Na 1.ª linha = várias pastas diretas sob o ID; mais abaixo = irmãs na mesma pasta-pai."
          ariaLabel="Adicionar pasta irmã"
          disabled={disabled}
          onClick={() => onAddSibling(node.id)}
        />
        <IconButton
          iconProps={{ iconName: 'Delete' }}
          title="Remover esta pasta e subpastas"
          ariaLabel="Remover pasta"
          disabled={disabled}
          onClick={() => onRemove(node.id)}
        />
      </Stack>
      {renderChildren(node.children, depth + 1)}
    </Stack>
  );
}

function FolderConfigAccordionItem(props: {
  node: IAttachmentLibraryFolderTreeNode;
  pathLabel: string;
  treeNodes: IAttachmentLibraryFolderTreeNode[];
  isOpen: boolean;
  onToggle: () => void;
  disabled: boolean;
  onTreeChange: (next: IAttachmentLibraryFolderTreeNode[]) => void;
  folderStepOptions: { id: string; title: string }[];
  onPatchStepIds: (id: string, stepIds: string[]) => void;
  folderVisibilityEditor?: IFolderVisibilityEditorProps;
}): JSX.Element {
  const [rulesExpanded, setRulesExpanded] = useState(false);
  const {
    node,
    pathLabel,
    treeNodes,
    isOpen,
    onToggle,
    disabled,
    onTreeChange,
    folderStepOptions,
    onPatchStepIds,
    folderVisibilityEditor,
  } = props;
  const rulesActive =
    node.showUploaderModes !== undefined ||
    (node.showUploaderGroupTitles?.length ?? 0) > 0 ||
    !!node.showUploaderWhen;
  const targetMark = node.uploadTarget ? ' · destino' : '';
  return (
    <FormManagerCollapseSection
      title={`${pathLabel}${targetMark}`}
      isOpen={isOpen}
      onToggle={onToggle}
    >
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { width: '100%', maxWidth: '100%' } }}>
        <Stack
          horizontal
          verticalAlign="center"
          tokens={{ childrenGap: 10 }}
          wrap
          styles={{ root: { width: '100%', alignItems: 'flex-start' } }}
        >
          <Text variant="small" styles={{ root: { color: '#605e5c', fontWeight: 600, paddingTop: 6 } }}>
            Etapa
          </Text>
          <Stack styles={{ root: { flex: 1, minWidth: 160, maxWidth: '100%' } }}>
            <FolderStepSelect
              node={node}
              disabled={disabled}
              folderStepOptions={folderStepOptions}
              onPatchStepIds={onPatchStepIds}
            />
          </Stack>
        </Stack>
        <Text variant="tiny" styles={{ root: { color: '#605e5c' } }}>
          Quantidade (biblioteca + novos). Vazio = sem limite.
        </Text>
        <Stack horizontal verticalAlign="end" tokens={{ childrenGap: 10 }} wrap>
          <TextField
            label="Mín."
            type="number"
            min={0}
            max={500}
            disabled={disabled}
            value={node.minAttachmentCount !== undefined ? String(node.minAttachmentCount) : ''}
            onChange={(_, v) => {
              const t = (v ?? '').trim();
              onTreeChange(
                updateAttachmentFolderNode(treeNodes, node.id, (no) => {
                  if (!t) {
                    const { minAttachmentCount: _m, ...rest } = no;
                    return rest;
                  }
                  const n = Math.min(500, Math.max(0, Math.floor(Number(t)) || 0));
                  if (n <= 0) {
                    const { minAttachmentCount: _m, ...rest } = no;
                    return rest;
                  }
                  const next = { ...no, minAttachmentCount: n };
                  if (next.maxAttachmentCount !== undefined && next.maxAttachmentCount < n) {
                    next.maxAttachmentCount = n;
                  }
                  return next;
                })
              );
            }}
            styles={{ root: { width: 88 } }}
          />
          <TextField
            label="Máx."
            type="number"
            min={0}
            max={500}
            disabled={disabled}
            value={node.maxAttachmentCount !== undefined ? String(node.maxAttachmentCount) : ''}
            onChange={(_, v) => {
              const t = (v ?? '').trim();
              onTreeChange(
                updateAttachmentFolderNode(treeNodes, node.id, (no) => {
                  if (!t) {
                    const { maxAttachmentCount: _x, ...rest } = no;
                    return rest;
                  }
                  const n = Math.min(500, Math.max(0, Math.floor(Number(t)) || 0));
                  const next = { ...no, maxAttachmentCount: n };
                  if (next.minAttachmentCount !== undefined && n < next.minAttachmentCount) {
                    next.minAttachmentCount = n;
                  }
                  return next;
                })
              );
            }}
            styles={{ root: { width: 88 } }}
          />
        </Stack>
        {folderVisibilityEditor && (
          <FormManagerCollapseSection
            title={rulesActive ? 'Regras opcionais · configuradas' : 'Regras opcionais'}
            isOpen={rulesExpanded}
            onToggle={() => setRulesExpanded((o) => !o)}
          >
            <FolderVisibilityRules
              node={node}
              treeNodes={treeNodes}
              disabled={disabled}
              onTreeChange={onTreeChange}
              editor={folderVisibilityEditor}
            />
          </FormManagerCollapseSection>
        )}
      </Stack>
    </FormManagerCollapseSection>
  );
}

export function FormManagerFolderTreeEditor(props: IFormManagerFolderTreeEditorProps): JSX.Element {
  const { nodes, onChange, disabled = false, folderStepOptions, showFolderStepPicker, folderVisibilityEditor } = props;
  const atMax = countNodesInTree(nodes) >= MAX_ATTACHMENT_FOLDER_TREE_NODES;
  const [openConfigFolderId, setOpenConfigFolderId] = useState<string | null>(null);

  const flatFolders = useMemo(() => flattenFolderTreeNodes(nodes), [nodes]);

  useEffect(() => {
    if (openConfigFolderId && !flatFolders.some((f) => f.id === openConfigFolderId)) {
      setOpenConfigFolderId(null);
    }
  }, [flatFolders, openConfigFolderId]);

  const toggleFolderConfig = (id: string): void => {
    setOpenConfigFolderId((prev) => (prev === id ? null : id));
  };

  const renderStructureChildren = (list: IAttachmentLibraryFolderTreeNode[] | undefined, depth: number): React.ReactNode => {
    if (!list?.length) return null;
    return (
      <Stack tokens={{ childrenGap: 4 }}>
        {list.map((node) => (
          <FolderStructureRow
            key={node.id}
            node={node}
            depth={depth}
            disabled={disabled}
            onPatchName={(id, v) => onChange(patchNodeName(nodes, id, v))}
            onAddChild={(id) => onChange(addChild(nodes, id))}
            onAddSibling={(id) => onChange(addSiblingAfter(nodes, id))}
            onRemove={(id) => onChange(removeNodeById(nodes, id))}
            onSetTarget={(id) => onChange(setUploadTargetById(nodes, id))}
            renderChildren={renderStructureChildren}
          />
        ))}
      </Stack>
    );
  };

  const showConfigSection = showFolderStepPicker && folderStepOptions.length > 0;

  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { width: '100%', maxWidth: '100%' } }}>
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { width: '100%', maxWidth: '100%' } }}>
        <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
          Estrutura de pastas
        </Text>
     
        {renderStructureChildren(nodes, 0)}
        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
          {nodes.length === 0 && (
            <DefaultButton
              iconProps={{ iconName: 'CreateNewFolder' }}
              text="Adicionar primeira pasta"
              title="Primeira pasta sob o ID; depois use + ou irmã para mais pastas"
              disabled={disabled || atMax}
              onClick={() => onChange(addRootSibling(nodes))}
            />
          )}
          {atMax && (
            <Text variant="tiny" styles={{ root: { color: '#a4262c' } }}>
              Limite de {MAX_ATTACHMENT_FOLDER_TREE_NODES} pastas.
            </Text>
          )}
        </Stack>
      </Stack>

      {showConfigSection && (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { width: '100%', maxWidth: '100%' } }}>
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
            Configurar cada pasta
          </Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Abra uma pasta de cada vez; ao abrir outra, esta fecha. Etapa e regras aplicam-se só a essa pasta.
          </Text>
          {flatFolders.length === 0 ? (
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Adicione pastas na secção acima.
            </Text>
          ) : (
            <Stack tokens={{ childrenGap: 6 }}>
              {flatFolders.map((fn) => (
                <FolderConfigAccordionItem
                  key={fn.id}
                  node={fn}
                  pathLabel={folderPathLabel(nodes, fn.id)}
                  treeNodes={nodes}
                  isOpen={openConfigFolderId === fn.id}
                  onToggle={() => toggleFolderConfig(fn.id)}
                  disabled={disabled}
                  onTreeChange={onChange}
                  folderStepOptions={folderStepOptions}
                  onPatchStepIds={(id, stepIds) => onChange(patchNodeShowUploaderStepIds(nodes, id, stepIds))}
                  folderVisibilityEditor={folderVisibilityEditor}
                />
              ))}
            </Stack>
          )}
        </Stack>
      )}
    </Stack>
  );
}
