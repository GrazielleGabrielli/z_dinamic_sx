import * as React from 'react';
import { useMemo } from 'react';
import { Stack, Text, DefaultButton, IconButton, Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type {
  IFormLinkedChildFormConfig,
  IFormManagerConfig,
  TFormAttachmentFilePreviewKind,
  TFormAttachmentUploadLayoutKind,
} from '../../core/config/types/formManager';
import type { ILinkedChildRowState } from '../../core/formManager/formLinkedChildSync';
import { flattenFolderTreeNodes, treeHasPerStepFolderUploaders } from '../../core/formManager/attachmentFolderTree';
import { linkedChildAttPendingKey } from '../../core/formManager/linkedChildAttachmentPendingKeys';
import { resolveLinkedChildAttachmentRuntime } from '../../core/formManager/linkedChildAttachmentRuntime';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { isAttachmentFolderUploaderVisible } from '../../core/formManager/formRuleEngine';
import { LinkedChildFormRowFields } from './LinkedChildFormRowFields';
import { FormAttachmentUploader } from './FormAttachmentUploader';

export interface ILinkedChildFormsBlockProps {
  configs: IFormLinkedChildFormConfig[];
  parentItemId: number | undefined;
  formMode: 'create' | 'edit' | 'view';
  rowsByConfigId: Record<string, ILinkedChildRowState[]>;
  onRowsChange: (configId: string, rows: ILinkedChildRowState[]) => void;
  fieldMetaByConfigId: Record<string, IFieldMetadata[]>;
  loadingByConfigId: Record<string, boolean>;
  errorByConfigId: Record<string, string | undefined>;
  userGroupTitles: string[];
  currentUserId: number;
  authorId: number | undefined;
  dynamicContext: IDynamicContext;
  rowErrorsByConfigId?: Record<string, Record<string, string>[]>;
  formManager: IFormManagerConfig;
  linkedPendingFilesByKey: Record<string, File[]>;
  onLinkedPendingFilesChange: (key: string, files: File[]) => void;
  currentParentStepId: string;
  attachmentUploadLayout?: TFormAttachmentUploadLayoutKind;
  attachmentFilePreview?: TFormAttachmentFilePreviewKind;
  attachmentAllowedExtensions?: string[];
}

function newLocalKey(): string {
  return `tmp_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

export const LinkedChildFormsBlock: React.FC<ILinkedChildFormsBlockProps> = ({
  configs,
  parentItemId,
  formMode,
  rowsByConfigId,
  onRowsChange,
  fieldMetaByConfigId,
  loadingByConfigId,
  errorByConfigId,
  userGroupTitles,
  currentUserId,
  authorId,
  dynamicContext,
  rowErrorsByConfigId,
  formManager,
  linkedPendingFilesByKey,
  onLinkedPendingFilesChange,
  currentParentStepId,
  attachmentUploadLayout,
  attachmentFilePreview,
  attachmentAllowedExtensions,
}) => {
  if (!configs.length) return null;

  const folderCtx = useMemo(
    () => ({
      formMode,
      values: {} as Record<string, unknown>,
      submitKind: 'submit' as const,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
    }),
    [formMode, userGroupTitles, currentUserId, authorId, dynamicContext]
  );

  return (
    <Stack tokens={{ childrenGap: 20 }} styles={{ root: { marginTop: 20 } }}>
      {configs.map((cfg) => {
        const rows = rowsByConfigId[cfg.id] ?? [];
        const meta = fieldMetaByConfigId[cfg.id] ?? [];
        const loading = loadingByConfigId[cfg.id] === true;
        const err = errorByConfigId[cfg.id];
        const minR = cfg.minRows ?? 0;
        const maxR = cfg.maxRows;
        const title = (cfg.title ?? cfg.listTitle).trim() || 'Lista vinculada';

        const addRow = (): void => {
          if (maxR !== undefined && rows.length >= maxR) return;
          onRowsChange(cfg.id, [...rows, { localKey: newLocalKey(), values: {} }]);
        };
        const removeRow = (idx: number): void => {
          if (rows.length <= minR && minR > 0) return;
          onRowsChange(
            cfg.id,
            rows.filter((_, j) => j !== idx)
          );
        };
        const moveRow = (from: number, to: number): void => {
          if (to < 0 || to >= rows.length) return;
          const next = rows.slice();
          const [m] = next.splice(from, 1);
          next.splice(to, 0, m);
          onRowsChange(cfg.id, next);
        };
        const patchRow = (idx: number, values: Record<string, unknown>): void => {
          const next = rows.map((r, j) => (j === idx ? { ...r, values } : r));
          onRowsChange(cfg.id, next);
        };

        const attResolved = resolveLinkedChildAttachmentRuntime(cfg, formManager);
        const stepFilterForFolders =
          cfg.childAttachmentStorageKind === 'documentLibraryInheritMain' ? currentParentStepId : 'main';
        const folderNodesForRow = (rowVals: Record<string, unknown>): ReturnType<typeof flattenFolderTreeNodes> => {
          if (attResolved.kind !== 'documentLibrary') return [];
          const tree = attResolved.folderTree;
          if (!tree?.length) return [];
          if (!treeHasPerStepFolderUploaders(tree)) {
            return flattenFolderTreeNodes(tree).filter((n) => n.uploadTarget);
          }
          return flattenFolderTreeNodes(tree).filter(
            (n) =>
              (n.showUploaderInStepIds?.length ?? 0) > 0 &&
              (n.showUploaderInStepIds ?? []).indexOf(stepFilterForFolders) !== -1 &&
              isAttachmentFolderUploaderVisible(n, {
                ...folderCtx,
                values: rowVals,
                attachmentFolderUrl: undefined,
              })
          );
        };

        return (
          <Stack
            key={cfg.id}
            tokens={{ childrenGap: 10 }}
            styles={{ root: { borderTop: '1px solid #edebe9', paddingTop: 16 } }}
          >
            <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
              {title}
            </Text>
            {!parentItemId && formMode === 'create' && (
              <MessageBar messageBarType={MessageBarType.info}>
                As linhas abaixo gravam depois de o registo principal ser guardado (ficam ligadas pelo campo
                Lookup).
              </MessageBar>
            )}
            {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
            {loading && <Spinner label="A carregar lista vinculada…" />}
            {!loading && meta.length === 0 && (
              <Text variant="small" styles={{ root: { color: '#a4262c' } }}>
                Não foi possível carregar campos da lista «{cfg.listTitle}». Verifique o título.
              </Text>
            )}
            {!loading &&
              meta.length > 0 &&
              rows.map((row, ri) => {
                const rowErrRaw = rowErrorsByConfigId?.[cfg.id]?.[ri];
                const blockMsg = rowErrRaw?._block;
                const rowErr: Record<string, string> = { ...(rowErrRaw ?? {}) };
                if (rowErr._block) delete rowErr._block;
                return (
                  <Stack
                    key={row.localKey}
                    tokens={{ childrenGap: 8 }}
                    styles={{
                      root: {
                        border: '1px solid #edebe9',
                        borderRadius: 4,
                        padding: 12,
                        background: '#faf9f8',
                      },
                    }}
                  >
                    <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                      <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                        Linha {ri + 1}
                        {row.sharePointId !== undefined ? ` · #${row.sharePointId}` : ''}
                      </Text>
                      {formMode !== 'view' && (
                        <Stack horizontal tokens={{ childrenGap: 4 }}>
                          <IconButton
                            iconProps={{ iconName: 'Up' }}
                            title="Mover para cima"
                            disabled={ri === 0}
                            onClick={() => moveRow(ri, ri - 1)}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Down' }}
                            title="Mover para baixo"
                            disabled={ri === rows.length - 1}
                            onClick={() => moveRow(ri, ri + 1)}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Remover linha"
                            disabled={rows.length <= minR && minR > 0}
                            onClick={() => removeRow(ri)}
                          />
                        </Stack>
                      )}
                    </Stack>
                    {blockMsg && (
                      <MessageBar messageBarType={MessageBarType.error}>{blockMsg}</MessageBar>
                    )}
                    <LinkedChildFormRowFields
                      childForm={cfg}
                      fieldMetadata={meta}
                      values={row.values}
                      onChange={(v) => patchRow(ri, v)}
                      formMode={formMode}
                      userGroupTitles={userGroupTitles}
                      currentUserId={currentUserId}
                      authorId={authorId}
                      dynamicContext={dynamicContext}
                      localErrors={rowErr}
                    />
                    {formMode !== 'view' && attResolved.kind !== 'none' && (
                      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 8 } }}>
                        <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                          {attResolved.kind === 'itemAttachments'
                            ? 'Anexos (lista filha)'
                            : 'Ficheiros na biblioteca (ligados ao registo desta linha)'}
                        </Text>
                        {(() => {
                          const multiLib =
                            attResolved.kind === 'documentLibrary' &&
                            !!attResolved.folderTree?.length &&
                            treeHasPerStepFolderUploaders(attResolved.folderTree);
                          const nodes = multiLib ? folderNodesForRow(row.values) : [];
                          if (multiLib && nodes.length > 0) {
                            return (
                              <Stack tokens={{ childrenGap: 10 }}>
                                {nodes.map((node) => {
                                  const pk = linkedChildAttPendingKey(cfg.id, row.localKey, node.id);
                                  return (
                                    <FormAttachmentUploader
                                      key={pk}
                                      files={linkedPendingFilesByKey[pk] ?? []}
                                      onFilesChange={(files) => onLinkedPendingFilesChange(pk, files)}
                                      disabled={false}
                                      label={node.nameTemplate?.trim() || 'Pasta'}
                                      layout={attachmentUploadLayout ?? 'default'}
                                      filePreview={attachmentFilePreview ?? 'nameAndSize'}
                                      allowedFileExtensions={attachmentAllowedExtensions}
                                    />
                                  );
                                })}
                              </Stack>
                            );
                          }
                          const flatKey = linkedChildAttPendingKey(cfg.id, row.localKey, '');
                          return (
                            <FormAttachmentUploader
                              files={linkedPendingFilesByKey[flatKey] ?? []}
                              onFilesChange={(files) => onLinkedPendingFilesChange(flatKey, files)}
                              disabled={false}
                              label={
                                attResolved.kind === 'itemAttachments' ? 'Anexos' : 'Ficheiros para enviar'
                              }
                              layout={attachmentUploadLayout ?? 'default'}
                              filePreview={attachmentFilePreview ?? 'nameAndSize'}
                              allowedFileExtensions={attachmentAllowedExtensions}
                            />
                          );
                        })()}
                      </Stack>
                    )}
                  </Stack>
                );
              })}
            {!loading && meta.length > 0 && formMode !== 'view' && (
              <DefaultButton
                text="Adicionar linha"
                disabled={maxR !== undefined && rows.length >= maxR}
                onClick={addRow}
              />
            )}
          </Stack>
        );
      })}
    </Stack>
  );
};
