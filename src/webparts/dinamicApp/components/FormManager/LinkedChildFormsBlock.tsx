import * as React from 'react';
import { Fragment, useMemo, useState } from 'react';
import {
  Stack,
  Text,
  DefaultButton,
  IconButton,
  Spinner,
  MessageBar,
  MessageBarType,
  Icon,
  type IStyle,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type {
  IFormLinkedChildFormConfig,
  IFormManagerConfig,
  TFormAttachmentFilePreviewKind,
  TFormAttachmentUploadLayoutKind,
  TLinkedChildRowsPresentationKind,
} from '../../core/config/types/formManager';
import type { ILinkedChildRowState } from '../../core/formManager/formLinkedChildSync';
import { getLinkedChildOrderedFieldConfigs } from '../../core/formManager/formLinkedChildSync';
import { flattenFolderTreeNodes, treeHasPerStepFolderUploaders } from '../../core/formManager/attachmentFolderTree';
import { linkedChildAttPendingKey } from '../../core/formManager/linkedChildAttachmentPendingKeys';
import { resolveLinkedChildAttachmentRuntime } from '../../core/formManager/linkedChildAttachmentRuntime';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { isAttachmentFolderUploaderVisible } from '../../core/formManager/formRuleEngine';
import { LinkedChildFormRowFields, type TLinkedChildFormRowFieldLayout } from './LinkedChildFormRowFields';
import { FormAttachmentUploader } from './FormAttachmentUploader';
import { attachmentFileKindIconName } from './attachmentFileKindIcon';
import { AttachmentFileDetailModal } from './AttachmentFileDetailModal';

export type ILinkedChildServerAttachmentRow = { fileName: string; fileUrl: string; fileRef?: string };

function LinkedChildServerAttachmentList(props: {
  rows: ILinkedChildServerAttachmentRow[];
  filePreview?: TFormAttachmentFilePreviewKind;
}): JSX.Element | null {
  const { rows, filePreview = 'nameAndSize' } = props;
  const [detailRow, setDetailRow] = useState<ILinkedChildServerAttachmentRow | null>(null);
  if (!rows.length) return null;
  const showIcon =
    filePreview === 'iconAndName' ||
    filePreview === 'thumbnailAndName' ||
    filePreview === 'thumbnailLarge';
  const iconPx = filePreview === 'thumbnailLarge' ? 48 : 20;
  const thumbBox = filePreview === 'thumbnailAndName' || filePreview === 'thumbnailLarge';
  const boxPx = filePreview === 'thumbnailLarge' ? 56 : 40;
  return (
    <>
      <Stack tokens={{ childrenGap: thumbBox ? 8 : 4 }}>
        {rows.map((a, ai) => (
          <Stack
            key={`${a.fileRef ?? a.fileUrl}-${a.fileName}-${ai}`}
            horizontal
            verticalAlign="center"
            tokens={{ childrenGap: 10 }}
            styles={{
              root: thumbBox
                ? {
                    padding: '8px 12px',
                    background: '#faf9f8',
                    borderRadius: 6,
                    border: '1px solid #edebe9',
                  }
                : undefined,
            }}
          >
            {showIcon &&
              (thumbBox ? (
                <div
                  style={{
                    width: boxPx,
                    height: boxPx,
                    borderRadius: 6,
                    background: '#edebe9',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    flexShrink: 0,
                  }}
                >
                  <Icon
                    iconName={attachmentFileKindIconName(a.fileName)}
                    styles={{ root: { fontSize: iconPx, color: '#605e5c' } }}
                  />
                </div>
              ) : (
                <Icon
                  iconName={attachmentFileKindIconName(a.fileName)}
                  styles={{ root: { fontSize: iconPx, color: '#0078d4', flexShrink: 0 } }}
                />
              ))}
            <Text
              variant="small"
              styles={{
                root: {
                  color: '#0078d4',
                  cursor: 'pointer',
                  textDecoration: 'underline',
                  wordBreak: 'break-word',
                },
              }}
              role="button"
              tabIndex={0}
              onClick={() => setDetailRow(a)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  setDetailRow(a);
                }
              }}
            >
              {a.fileName}
            </Text>
          </Stack>
        ))}
      </Stack>
      <AttachmentFileDetailModal
        isOpen={detailRow !== null}
        onDismiss={() => setDetailRow(null)}
        target={
          detailRow
            ? {
                kind: 'server',
                fileName: detailRow.fileName,
                fileUrl: detailRow.fileUrl,
                fileRef: detailRow.fileRef,
              }
            : null
        }
      />
    </>
  );
}

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
  linkedServerAttachmentsByKey: Record<string, ILinkedChildServerAttachmentRow[]>;
}

function newLocalKey(): string {
  return `tmp_${Date.now()}_${Math.random().toString(36).slice(2, 9)}`;
}

function linkedChildRowSurfaceStyle(kind: TLinkedChildRowsPresentationKind): IStyle {
  switch (kind) {
    case 'compact':
      return {
        border: '1px solid #edebe9',
        borderRadius: 4,
        padding: 8,
        background: '#faf9f8',
      };
    case 'cards':
      return {
        border: '1px solid #e1dfdd',
        borderRadius: 8,
        padding: 14,
        background: '#ffffff',
        boxShadow: '0 1.6px 3.6px rgba(0, 0, 0, 0.09)',
      };
    default:
      return {
        border: '1px solid #edebe9',
        borderRadius: 4,
        padding: 12,
        background: '#faf9f8',
      };
  }
}

type ILinkedChildListSectionProps = {
  cfg: IFormLinkedChildFormConfig;
  parentItemId: number | undefined;
  formMode: 'create' | 'edit' | 'view';
  rows: ILinkedChildRowState[];
  onRowsChange: (rows: ILinkedChildRowState[]) => void;
  meta: IFieldMetadata[];
  loading: boolean;
  err: string | undefined;
  rowErrors?: Record<string, string>[];
  formManager: IFormManagerConfig;
  linkedPendingFilesByKey: Record<string, File[]>;
  onLinkedPendingFilesChange: (key: string, files: File[]) => void;
  currentParentStepId: string;
  attachmentUploadLayout?: TFormAttachmentUploadLayoutKind;
  attachmentFilePreview?: TFormAttachmentFilePreviewKind;
  attachmentAllowedExtensions?: string[];
  linkedServerAttachmentsByKey: Record<string, ILinkedChildServerAttachmentRow[]>;
  userGroupTitles: string[];
  currentUserId: number;
  authorId: number | undefined;
  dynamicContext: IDynamicContext;
  folderCtx: {
    formMode: 'create' | 'edit' | 'view';
    values: Record<string, unknown>;
    submitKind: 'submit';
    userGroupTitles: string[];
    currentUserId: number;
    authorId: number | undefined;
    dynamicContext: IDynamicContext;
  };
};

const LinkedChildListSection: React.FC<ILinkedChildListSectionProps> = ({
  cfg,
  parentItemId,
  formMode,
  rows,
  onRowsChange,
  meta,
  loading,
  err,
  rowErrors,
  formManager,
  linkedPendingFilesByKey,
  onLinkedPendingFilesChange,
  currentParentStepId,
  attachmentUploadLayout,
  attachmentFilePreview,
  attachmentAllowedExtensions,
  linkedServerAttachmentsByKey,
  userGroupTitles,
  currentUserId,
  authorId,
  dynamicContext,
  folderCtx,
}) => {
  const presentation = cfg.rowsPresentation ?? 'stack';
  const surfaceKind: TLinkedChildRowsPresentationKind =
    presentation === 'table' ? 'stack' : presentation;
  const stackFieldLayout: TLinkedChildFormRowFieldLayout =
    presentation === 'compact' ? 'compact' : 'stack';

  const title = (cfg.title ?? cfg.listTitle).trim() || 'Lista vinculada';
  const minR = cfg.minRows ?? 0;
  const maxR = cfg.maxRows;

  const addRow = (): void => {
    if (maxR !== undefined && rows.length >= maxR) return;
    onRowsChange([...rows, { localKey: newLocalKey(), values: {} }]);
  };
  const removeRow = (idx: number): void => {
    if (rows.length <= minR && minR > 0) return;
    onRowsChange(rows.filter((_, j) => j !== idx));
  };
  const moveRow = (from: number, to: number): void => {
    if (to < 0 || to >= rows.length) return;
    const next = rows.slice();
    const [m] = next.splice(from, 1);
    next.splice(to, 0, m);
    onRowsChange(next);
  };
  const patchRow = (idx: number, values: Record<string, unknown>): void => {
    onRowsChange(rows.map((r, j) => (j === idx ? { ...r, values } : r)));
  };

  const attResolved = useMemo(
    () => resolveLinkedChildAttachmentRuntime(cfg, formManager),
    [cfg, formManager]
  );
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

  const orderedForTable = useMemo(() => getLinkedChildOrderedFieldConfigs(cfg), [cfg]);
  const tableHeaders = useMemo(() => {
    const byName = new Map(meta.map((m) => [m.InternalName, m]));
    return orderedForTable.map((fc) => {
      const mm = byName.get(fc.internalName);
      return (fc.label ?? mm?.Title ?? fc.internalName).trim();
    });
  }, [meta, orderedForTable]);

  const renderAttachmentBlock = (row: ILinkedChildRowState): JSX.Element | null => {
    if (attResolved.kind === 'none') return null;
    return (
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: presentation === 'table' ? 4 : 8 } }}>
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
                  const serverRows = linkedServerAttachmentsByKey[pk] ?? [];
                  return (
                    <Stack key={pk} tokens={{ childrenGap: 6 }}>
                      <LinkedChildServerAttachmentList rows={serverRows} filePreview={attachmentFilePreview} />
                      {formMode !== 'view' && (
                        <FormAttachmentUploader
                          files={linkedPendingFilesByKey[pk] ?? []}
                          onFilesChange={(files) => onLinkedPendingFilesChange(pk, files)}
                          disabled={false}
                          label={node.nameTemplate?.trim() || 'Pasta'}
                          layout={attachmentUploadLayout ?? 'default'}
                          filePreview={attachmentFilePreview ?? 'nameAndSize'}
                          allowedFileExtensions={attachmentAllowedExtensions}
                          priorFileCount={serverRows.length}
                        />
                      )}
                    </Stack>
                  );
                })}
              </Stack>
            );
          }
          const flatKey = linkedChildAttPendingKey(cfg.id, row.localKey, '');
          const flatServer = linkedServerAttachmentsByKey[flatKey] ?? [];
          return (
            <Stack tokens={{ childrenGap: 6 }}>
              <LinkedChildServerAttachmentList rows={flatServer} filePreview={attachmentFilePreview} />
              {formMode !== 'view' && (
                <FormAttachmentUploader
                  files={linkedPendingFilesByKey[flatKey] ?? []}
                  onFilesChange={(files) => onLinkedPendingFilesChange(flatKey, files)}
                  disabled={false}
                  label={attResolved.kind === 'itemAttachments' ? 'Anexos' : 'Ficheiros para enviar'}
                  layout={attachmentUploadLayout ?? 'default'}
                  filePreview={attachmentFilePreview ?? 'nameAndSize'}
                  allowedFileExtensions={attachmentAllowedExtensions}
                  priorFileCount={flatServer.length}
                />
              )}
            </Stack>
          );
        })()}
      </Stack>
    );
  };

  const thStyle: React.CSSProperties = {
    textAlign: 'left',
    padding: '8px 10px',
    borderBottom: '2px solid #edebe9',
    borderRight: '1px solid #edebe9',
    fontWeight: 600,
    fontSize: 12,
    color: '#323130',
    background: '#f3f2f1',
    whiteSpace: 'nowrap',
  };

  const innerBlocks = (
    <>
      {!parentItemId && formMode === 'create' && (
        <MessageBar messageBarType={MessageBarType.info}>
          As linhas abaixo gravam depois de o registo principal ser guardado (ficam ligadas pelo campo Lookup).
        </MessageBar>
      )}
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      {loading && <Spinner label="A carregar lista vinculada…" />}
      {!loading && meta.length === 0 && (
        <Text variant="small" styles={{ root: { color: '#a4262c' } }}>
          Não foi possível carregar campos da lista «{cfg.listTitle}». Verifique o título.
        </Text>
      )}
      {!loading && meta.length > 0 && presentation === 'table' && (
        <div style={{ overflowX: 'auto' }}>
          <table
            style={{
              width: '100%',
              borderCollapse: 'collapse',
              fontSize: 13,
              border: '1px solid #edebe9',
            }}
          >
            <thead>
              <tr>
                <th style={thStyle}> </th>
                {tableHeaders.map((h, hi) => (
                  <th key={hi} style={thStyle}>
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, ri) => {
                const rowErrRaw = rowErrors?.[ri];
                const blockMsg = rowErrRaw?._block;
                const rowErr: Record<string, string> = { ...(rowErrRaw ?? {}) };
                if (rowErr._block) delete rowErr._block;
                const colspan = 1 + orderedForTable.length;
                return (
                  <Fragment key={row.localKey}>
                    <tr>
                      <td
                        style={{
                          verticalAlign: 'top',
                          padding: '8px 10px',
                          borderBottom: '1px solid #edebe9',
                          borderRight: '1px solid #edebe9',
                          whiteSpace: 'nowrap',
                          background: '#faf9f8',
                        }}
                      >
                        <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }}>
                          <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                            {ri + 1}
                            {row.sharePointId !== undefined ? ` · #${row.sharePointId}` : ''}
                          </Text>
                          {formMode !== 'view' && (
                            <Stack horizontal tokens={{ childrenGap: 2 }}>
                              <IconButton
                                iconProps={{ iconName: 'Up' }}
                                title="Mover para cima"
                                disabled={ri === 0}
                                styles={{ root: { height: 24, width: 28 } }}
                                onClick={() => moveRow(ri, ri - 1)}
                              />
                              <IconButton
                                iconProps={{ iconName: 'Down' }}
                                title="Mover para baixo"
                                disabled={ri === rows.length - 1}
                                styles={{ root: { height: 24, width: 28 } }}
                                onClick={() => moveRow(ri, ri + 1)}
                              />
                              <IconButton
                                iconProps={{ iconName: 'Delete' }}
                                title="Remover linha"
                                disabled={rows.length <= minR && minR > 0}
                                styles={{ root: { height: 24, width: 28 } }}
                                onClick={() => removeRow(ri)}
                              />
                            </Stack>
                          )}
                        </Stack>
                      </td>
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
                        fieldLayout="tableCells"
                        rowPersisted={
                          row.sharePointId !== undefined &&
                          typeof row.sharePointId === 'number' &&
                          isFinite(row.sharePointId)
                        }
                      />
                    </tr>
                    {blockMsg && (
                      <tr>
                        <td colSpan={colspan} style={{ padding: 0, borderBottom: '1px solid #edebe9' }}>
                          <MessageBar messageBarType={MessageBarType.error}>{blockMsg}</MessageBar>
                        </td>
                      </tr>
                    )}
                    {attResolved.kind !== 'none' && (
                      <tr>
                        <td
                          colSpan={colspan}
                          style={{
                            padding: '10px 12px',
                            borderBottom: '1px solid #edebe9',
                            background: '#faf9f8',
                          }}
                        >
                          {renderAttachmentBlock(row)}
                        </td>
                      </tr>
                    )}
                  </Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
      {!loading && meta.length > 0 && presentation !== 'table' &&
        rows.map((row, ri) => {
          const rowErrRaw = rowErrors?.[ri];
          const blockMsg = rowErrRaw?._block;
          const rowErr: Record<string, string> = { ...(rowErrRaw ?? {}) };
          if (rowErr._block) delete rowErr._block;
          return (
            <Stack
              key={row.localKey}
              tokens={{ childrenGap: presentation === 'compact' ? 6 : 8 }}
              styles={{ root: linkedChildRowSurfaceStyle(surfaceKind) }}
            >
              <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
                <Text
                  variant={presentation === 'compact' ? 'small' : 'small'}
                  styles={{ root: { fontWeight: 600, color: '#323130' } }}
                >
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
              {blockMsg && <MessageBar messageBarType={MessageBarType.error}>{blockMsg}</MessageBar>}
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
                fieldLayout={stackFieldLayout}
                rowPersisted={
                  row.sharePointId !== undefined &&
                  typeof row.sharePointId === 'number' &&
                  isFinite(row.sharePointId)
                }
              />
              {renderAttachmentBlock(row)}
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
    </>
  );

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { borderTop: '1px solid #edebe9', paddingTop: 16 } }}>
      <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
        {title}
      </Text>
      {innerBlocks}
    </Stack>
  );
};

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
  linkedServerAttachmentsByKey,
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
      {configs.map((cfg) => (
        <LinkedChildListSection
          key={cfg.id}
          cfg={cfg}
          parentItemId={parentItemId}
          formMode={formMode}
          rows={rowsByConfigId[cfg.id] ?? []}
          onRowsChange={(next) => onRowsChange(cfg.id, next)}
          meta={fieldMetaByConfigId[cfg.id] ?? []}
          loading={loadingByConfigId[cfg.id] === true}
          err={errorByConfigId[cfg.id]}
          rowErrors={rowErrorsByConfigId?.[cfg.id]}
          formManager={formManager}
          linkedPendingFilesByKey={linkedPendingFilesByKey}
          onLinkedPendingFilesChange={onLinkedPendingFilesChange}
          currentParentStepId={currentParentStepId}
          attachmentUploadLayout={attachmentUploadLayout}
          attachmentFilePreview={attachmentFilePreview}
          attachmentAllowedExtensions={attachmentAllowedExtensions}
          linkedServerAttachmentsByKey={linkedServerAttachmentsByKey}
          userGroupTitles={userGroupTitles}
          currentUserId={currentUserId}
          authorId={authorId}
          dynamicContext={dynamicContext}
          folderCtx={folderCtx}
        />
      ))}
    </Stack>
  );
};
