import * as React from 'react';
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
import type { IGroupDetails } from '../../../../services';
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
} from '../../core/formManager/attachmentFolderTree';
import {
  CONDITION_OP_OPTIONS,
  whenNodeToUi,
  whenUiToNode,
  type IWhenUi,
} from '../../core/formManager/formManagerVisualModel';

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
      styles={{ root: { maxWidth: 320 } }}
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
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { maxWidth: 520 } }}>
      <Text variant="tiny" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
        Quem vê este input e em que modos
      </Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Modos do formulário (vazio nas caixas = todos). Desmarque para restringir.
      </Text>
      <Stack horizontal wrap tokens={{ childrenGap: 16 }}>
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
      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Grupos do SharePoint</Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Só utilizadores em pelo menos um dos grupos marcados veem o input desta pasta. Vazio = todos.
      </Text>
      {siteGroupsLoading && <Spinner label="A carregar grupos…" />}
      {siteGroupsErr && (
        <>
          <MessageBar messageBarType={MessageBarType.warning}>{siteGroupsErr}</MessageBar>
          <DefaultButton text="Tentar novamente" onClick={onReloadSiteGroups} />
        </>
      )}
      {!siteGroupsLoading ? (
        <Stack
          tokens={{ childrenGap: 6 }}
          styles={{
            root: {
              maxHeight: 200,
              overflowY: 'auto',
              border: '1px solid #edebe9',
              borderRadius: 4,
              padding: 8,
            },
          }}
        >
          {(node.showUploaderGroupTitles ?? [])
            .filter((t) => !siteGroups.some((g) => normFolderRuleGroupTitle(g.Title) === normFolderRuleGroupTitle(t)))
            .map((t, oi) => (
              <Checkbox
                key={`orphan-grp-${node.id}-${oi}-${t}`}
                label={`${t} (guardado; não na lista)`}
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
          {siteGroupsSorted.map((g) => {
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
          {!siteGroupsSorted.length && !(node.showUploaderGroupTitles ?? []).length && (
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Nenhum grupo no site.
            </Text>
          )}
        </Stack>
      ) : null}
      <Checkbox
        label="Só mostrar quando a condição abaixo for verdadeira"
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
          Condição composta (várias cláusulas). Edição completa: JSON avançado. Desmarque a caixa acima para remover.
        </MessageBar>
      )}
      {node.showUploaderWhen && leafWhen && (
        <Stack tokens={{ childrenGap: 8 }}>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            Condição nos dados do formulário
          </Text>
          <Stack horizontal wrap tokens={{ childrenGap: 8 }} verticalAlign="end">
            <Dropdown
              label="Campo"
              options={fieldOptions}
              selectedKey={leafWhen.field}
              disabled={disabled}
              onChange={(_, o) => o && patchWhenUi({ field: String(o.key) })}
            />
            <Dropdown
              label="Operador"
              options={CONDITION_OP_OPTIONS.map((x) => ({ key: x.key, text: x.text }))}
              selectedKey={leafWhen.op}
              disabled={disabled}
              onChange={(_, o) => o && patchWhenUi({ op: o.key as TFormConditionOp })}
            />
            <Dropdown
              label="Comparar com"
              options={[
                { key: 'literal', text: 'Texto fixo' },
                { key: 'field', text: 'Outro campo' },
                { key: 'token', text: 'Token' },
              ]}
              selectedKey={leafWhen.compareKind}
              disabled={disabled}
              onChange={(_, o) => o && patchWhenUi({ compareKind: o.key as IWhenUi['compareKind'] })}
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
            />
          </Stack>
        </Stack>
      )}
    </Stack>
  );
}

function FolderRow(props: {
  node: IAttachmentLibraryFolderTreeNode;
  treeNodes: IAttachmentLibraryFolderTreeNode[];
  depth: number;
  disabled: boolean;
  onPatchName: (id: string, v: string) => void;
  onAddChild: (id: string) => void;
  onAddSibling: (id: string) => void;
  onRemove: (id: string) => void;
  onSetTarget: (id: string) => void;
  renderChildren: (nodes: IAttachmentLibraryFolderTreeNode[] | undefined, d: number) => React.ReactNode;
  allowSiblingAtDepth: boolean;
  folderStepOptions: { id: string; title: string }[];
  showFolderStepPicker: boolean;
  onPatchStepIds: (id: string, stepIds: string[]) => void;
  onTreeChange: (next: IAttachmentLibraryFolderTreeNode[]) => void;
  folderVisibilityEditor?: IFolderVisibilityEditorProps;
}): JSX.Element {
  const {
    node,
    treeNodes,
    depth,
    disabled,
    onPatchName,
    onAddChild,
    onAddSibling,
    onRemove,
    onSetTarget,
    renderChildren,
    allowSiblingAtDepth,
    folderStepOptions,
    showFolderStepPicker,
    onPatchStepIds,
    onTreeChange,
    folderVisibilityEditor,
  } = props;
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
        <Stack styles={{ root: { flex: '1 1 220px', maxWidth: 420, minWidth: 160 } }}>
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
          title={
            allowSiblingAtDepth
              ? 'Pasta ao mesmo nível (abaixo desta)'
              : 'Só existe uma pasta raiz sob o ID do item; use subpastas ou níveis abaixo.'
          }
          ariaLabel="Adicionar pasta irmã"
          disabled={disabled || !allowSiblingAtDepth}
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
      {showFolderStepPicker && folderStepOptions.length > 0 && (
        <Stack
          styles={{ root: { paddingLeft: pad + 8, paddingTop: 4, maxWidth: 520 } }}
          tokens={{ childrenGap: 8 }}
        >
          <Text variant="tiny" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
            Input de ficheiros nesta pasta na etapa:
          </Text>
          <FolderStepSelect
            node={node}
            disabled={disabled}
            folderStepOptions={folderStepOptions}
            onPatchStepIds={onPatchStepIds}
          />
          {folderVisibilityEditor && (
            <FolderVisibilityRules
              node={node}
              treeNodes={treeNodes}
              disabled={disabled}
              onTreeChange={onTreeChange}
              editor={folderVisibilityEditor}
            />
          )}
        </Stack>
      )}
      {renderChildren(node.children, depth + 1)}
    </Stack>
  );
}

export function FormManagerFolderTreeEditor(props: IFormManagerFolderTreeEditorProps): JSX.Element {
  const { nodes, onChange, disabled = false, folderStepOptions, showFolderStepPicker, folderVisibilityEditor } = props;
  const atMax = countNodesInTree(nodes) >= MAX_ATTACHMENT_FOLDER_TREE_NODES;

  const renderChildren = (list: IAttachmentLibraryFolderTreeNode[] | undefined, depth: number): React.ReactNode => {
    if (!list?.length) return null;
    return (
      <Stack tokens={{ childrenGap: 4 }}>
        {list.map((node) => (
          <FolderRow
            key={node.id}
            node={node}
            treeNodes={nodes}
            depth={depth}
            disabled={disabled}
            onPatchName={(id, v) => onChange(patchNodeName(nodes, id, v))}
            onAddChild={(id) => onChange(addChild(nodes, id))}
            onAddSibling={(id) => onChange(addSiblingAfter(nodes, id))}
            onRemove={(id) => onChange(removeNodeById(nodes, id))}
            onSetTarget={(id) => onChange(setUploadTargetById(nodes, id))}
            renderChildren={renderChildren}
            allowSiblingAtDepth={depth > 0}
            folderStepOptions={folderStepOptions}
            showFolderStepPicker={showFolderStepPicker}
            onPatchStepIds={(id, stepIds) => onChange(patchNodeShowUploaderStepIds(nodes, id, stepIds))}
            onTreeChange={onChange}
            folderVisibilityEditor={folderVisibilityEditor}
          />
        ))}
      </Stack>
    );
  };

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Sob a pasta do ID do item existe no máximo uma pasta raiz; a estrutura fica dentro dela. Por pasta, pode
        indicar em que etapa aparece o input de ficheiros (mesmo layout global); os ficheiros gravam nessa pasta na
        biblioteca. Regras opcionais: modos (criar/editar/ver), grupos e condição em campos.
      </Text>
      {renderChildren(nodes, 0)}
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
        {nodes.length === 0 && (
          <DefaultButton
            iconProps={{ iconName: 'CreateNewFolder' }}
            text="Adicionar primeira pasta (raiz sob o ID)"
            title="Uma única raiz; depois use subpastas e pastas ao mesmo nível nos níveis inferiores"
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
  );
}
