import * as React from 'react';
import { Stack, Text, TextField, IconButton, DefaultButton } from '@fluentui/react';
import type { IAttachmentLibraryFolderTreeNode } from '../../core/config/types/formManager';
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
  setUploadTargetById,
} from '../../core/formManager/attachmentFolderTree';

export interface IFormManagerFolderTreeEditorProps {
  nodes: IAttachmentLibraryFolderTreeNode[];
  onChange: (next: IAttachmentLibraryFolderTreeNode[]) => void;
  disabled?: boolean;
}

function FolderRow(props: {
  node: IAttachmentLibraryFolderTreeNode;
  depth: number;
  disabled: boolean;
  onPatchName: (id: string, v: string) => void;
  onAddChild: (id: string) => void;
  onAddSibling: (id: string) => void;
  onRemove: (id: string) => void;
  onSetTarget: (id: string) => void;
  renderChildren: (nodes: IAttachmentLibraryFolderTreeNode[] | undefined, d: number) => React.ReactNode;
  allowSiblingAtDepth: boolean;
}): JSX.Element {
  const {
    node,
    depth,
    disabled,
    onPatchName,
    onAddChild,
    onAddSibling,
    onRemove,
    onSetTarget,
    renderChildren,
    allowSiblingAtDepth,
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
      {renderChildren(node.children, depth + 1)}
    </Stack>
  );
}

export function FormManagerFolderTreeEditor(props: IFormManagerFolderTreeEditorProps): JSX.Element {
  const { nodes, onChange, disabled = false } = props;
  const atMax = countNodesInTree(nodes) >= MAX_ATTACHMENT_FOLDER_TREE_NODES;

  const renderChildren = (list: IAttachmentLibraryFolderTreeNode[] | undefined, depth: number): React.ReactNode => {
    if (!list?.length) return null;
    return (
      <Stack tokens={{ childrenGap: 4 }}>
        {list.map((node) => (
          <FolderRow
            key={node.id}
            node={node}
            depth={depth}
            disabled={disabled}
            onPatchName={(id, v) => onChange(patchNodeName(nodes, id, v))}
            onAddChild={(id) => onChange(addChild(nodes, id))}
            onAddSibling={(id) => onChange(addSiblingAfter(nodes, id))}
            onRemove={(id) => onChange(removeNodeById(nodes, id))}
            onSetTarget={(id) => onChange(setUploadTargetById(nodes, id))}
            renderChildren={renderChildren}
            allowSiblingAtDepth={depth > 0}
          />
        ))}
      </Stack>
    );
  };

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Sob a pasta do ID do item existe no máximo uma pasta raiz; a estrutura (irmãos e filhos) fica dentro dela, para
        localizar tudo pelo ID da solicitação.
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
