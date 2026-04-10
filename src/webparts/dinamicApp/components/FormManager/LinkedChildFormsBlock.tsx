import * as React from 'react';
import { Stack, Text, DefaultButton, IconButton, Spinner, MessageBar, MessageBarType } from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type { IFormLinkedChildFormConfig } from '../../core/config/types/formManager';
import type { ILinkedChildRowState } from '../../core/formManager/formLinkedChildSync';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { LinkedChildFormRowFields } from './LinkedChildFormRowFields';

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
}) => {
  if (!configs.length) return null;

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
