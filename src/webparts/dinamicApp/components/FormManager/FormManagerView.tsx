import * as React from 'react';
import { useState, useEffect, useMemo, useCallback } from 'react';
import { Stack, MessageBar, MessageBarType, Text, type IStyle } from '@fluentui/react';
import type { IDynamicViewConfig } from '../../core/config/types';
import { getDefaultFormManagerConfig } from '../../core/config/utils';
import { buildDynamicContext, parseQueryString } from '../../core/dynamicTokens';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { FieldsService, ItemsService, UsersService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import { getSP } from '../../../../services/core/sp';
import { DynamicListForm } from './DynamicListForm';
import { FormDataLoadingView, resolveFormDataLoadingKind } from './FormLoadingUi';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  type IFormManagerConfig,
  type TFormManagerFormMode,
  type TFormSubmitKind,
} from '../../core/config/types/formManager';
import { sanitizeFormManagerConfig } from '../../core/formManager/sanitizeFormManagerConfig';
import {
  isFormAttachmentLibraryRuntime,
  uploadFilesToAttachmentLibrary,
  uploadFilesToAttachmentLibraryByFolderNodes,
} from '../../core/formManager/formAttachmentLibrary';
import { resolveFormRootLayoutStyles } from '../../core/formManager/formRootLayout';
import {
  isFormNewModeQuery,
  parseFormItemIdFromQuery,
  resolveFormModeFromQuery,
} from '../../core/formManager/formQueryParams';

export interface IFormManagerViewProps {
  config: IDynamicViewConfig;
}

async function uploadAttachments(
  fm: IFormManagerConfig,
  listTitle: string,
  itemId: number,
  files: File[],
  itemFieldValues: Record<string, unknown>,
  filesByFolderNodeId?: Record<string, File[]>
): Promise<void> {
  const hasBuckets =
    !!filesByFolderNodeId && Object.keys(filesByFolderNodeId).some((k) => filesByFolderNodeId[k].length > 0);
  if (!files.length && !hasBuckets) return;
  if (isFormAttachmentLibraryRuntime(fm)) {
    const lib = fm.attachmentLibrary!;
    const iv = { ...itemFieldValues, Id: itemId };
    if (hasBuckets) {
      await uploadFilesToAttachmentLibraryByFolderNodes(
        lib.libraryTitle!,
        lib.sourceListLookupFieldInternalName!,
        itemId,
        filesByFolderNodeId!,
        lib.folderTree,
        { itemFieldValues: iv }
      );
      return;
    }
    await uploadFilesToAttachmentLibrary(
      lib.libraryTitle!,
      lib.sourceListLookupFieldInternalName!,
      itemId,
      files,
      {
        folderTree: lib.folderTree,
        folderPathSegments: lib.folderPathSegments,
        itemFieldValues: iv,
      }
    );
    return;
  }
  const sp = getSP();
  const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(listTitle);
  const list = isGuid ? sp.web.lists.getById(listTitle) : sp.web.lists.getByTitle(listTitle);
  const item = list.items.getById(itemId) as unknown as {
    attachmentFiles: { add(name: string, content: ArrayBuffer): Promise<unknown> };
  };
  for (let i = 0; i < files.length; i++) {
    const buf = await files[i].arrayBuffer();
    await item.attachmentFiles.add(files[i].name, buf);
  }
}

function formFieldInternalNames(fm: IFormManagerConfig, fieldMeta: IFieldMetadata[]): string[] {
  const skipVirtual = (internalName: string): boolean => internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL;
  if (fm.fields.length > 0) return fm.fields.map((f) => f.internalName).filter(skipVirtual);
  return fieldMeta
    .filter((f) => !f.Hidden && !f.ReadOnlyField && f.InternalName !== 'Id')
    .map((f) => f.InternalName)
    .filter(skipVirtual);
}

function buildSelectExpandForFields(fieldNames: string[], fieldMeta: IFieldMetadata[]): { select: string[]; expand: string[] } {
  const select: string[] = ['Id'];
  const expand: string[] = [];
  const byName = new Map(fieldMeta.map((f) => [f.InternalName, f]));
  for (let i = 0; i < fieldNames.length; i++) {
    const name = fieldNames[i];
    if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) continue;
    const m = byName.get(name);
    const needsExpand = m && ['lookup', 'lookupmulti', 'user', 'usermulti'].indexOf(m.MappedType) !== -1;
    if (needsExpand && m) {
      if (expand.indexOf(name) === -1) expand.push(name);
      const ef = m.LookupField || 'Title';
      if (select.indexOf(`${name}/Id`) === -1) select.push(`${name}/Id`, `${name}/${ef}`);
    } else if (select.indexOf(name) === -1) select.push(name);
  }
  if (select.indexOf('AuthorId') === -1) select.push('AuthorId');
  if (expand.indexOf('Author') === -1) {
    select.push('Author/Id', 'Author/Title');
    expand.push('Author');
  }
  return { select, expand };
}

export const FormManagerView: React.FC<IFormManagerViewProps> = ({ config }) => {
  const fm = useMemo(() => {
    const raw = config.formManager ?? getDefaultFormManagerConfig();
    return sanitizeFormManagerConfig(raw) ?? raw;
  }, [config.formManager]);
  const listTitle = config.dataSource.title;

  const [fieldMeta, setFieldMeta] = useState<IFieldMetadata[]>([]);
  const [metaLoading, setMetaLoading] = useState(true);
  const [dynamicContext, setDynamicContext] = useState<IDynamicContext | undefined>(undefined);
  const [userGroupTitles, setUserGroupTitles] = useState<string[]>([]);
  const [currentUserId, setCurrentUserId] = useState(0);

  const [formMode, setFormMode] = useState<TFormManagerFormMode>('create');
  const [activeItem, setActiveItem] = useState<Record<string, unknown> | null>(null);
  const [formKey, setFormKey] = useState(0);
  const [loadError, setLoadError] = useState<string | undefined>(undefined);
  const [itemLoading, setItemLoading] = useState(false);

  const itemsService = useMemo(() => new ItemsService(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);

  const fieldNames = useMemo(() => formFieldInternalNames(fm, fieldMeta), [fm, fieldMeta]);

  useEffect(() => {
    const usersService = new UsersService();
    usersService
      .getCurrentUser()
      .then((user) => {
        setCurrentUserId(user.Id);
        setDynamicContext(
          buildDynamicContext({
            currentUser: {
              id: user.Id,
              title: user.Title,
              name: user.Title,
              email: user.Email,
              loginName: user.LoginName,
            },
            query: typeof window !== 'undefined' && window.location ? parseQueryString(window.location.search) : undefined,
            now: new Date(),
            list: { title: listTitle },
          })
        );
        return usersService.getUserGroups(user.LoginName);
      })
      .then((groups) => setUserGroupTitles(groups.map((g) => g.Title)))
      .catch(() => {
        setDynamicContext(buildDynamicContext({ now: new Date(), list: { title: listTitle } }));
      });
  }, [listTitle]);

  useEffect(() => {
    if (!listTitle.trim()) return;
    setMetaLoading(true);
    fieldsService
      .getVisibleFields(listTitle.trim())
      .then((f) => {
        setFieldMeta(f);
        setMetaLoading(false);
      })
      .catch(() => {
        setFieldMeta([]);
        setMetaLoading(false);
      });
  }, [listTitle, fieldsService]);

  const loadItemById = useCallback(
    async (itemId: number, modeAfterLoad: TFormManagerFormMode): Promise<void> => {
      if (!listTitle.trim() || !fieldMeta.length) return;
      setItemLoading(true);
      setLoadError(undefined);
      const { select, expand } = buildSelectExpandForFields(fieldNames, fieldMeta);
      try {
        const row = await itemsService.getItemById<Record<string, unknown>>(listTitle, itemId, {
          select,
          expand: expand.length ? expand : undefined,
          fieldMetadata: fieldMeta,
        });
        setActiveItem(row);
        setFormMode(modeAfterLoad);
      } catch (e) {
        setLoadError(e instanceof Error ? e.message : String(e));
        setActiveItem(null);
        setFormMode('create');
      } finally {
        setItemLoading(false);
      }
    },
    [listTitle, fieldMeta, fieldNames, itemsService]
  );

  useEffect(() => {
    if (!fieldMeta.length || !dynamicContext?.query) return;
    const q = dynamicContext.query;
    if (isFormNewModeQuery(q)) {
      setActiveItem(null);
      setFormMode('create');
      setLoadError(undefined);
      return;
    }
    const id = parseFormItemIdFromQuery(q);
    if (id === undefined) return;
    void loadItemById(id, resolveFormModeFromQuery(q, { itemLoaded: true }));
  }, [fieldMeta.length, dynamicContext?.query, loadItemById]);

  const resetToNew = useCallback((): void => {
    setActiveItem(null);
    setFormMode('create');
    setLoadError(undefined);
    setFormKey((k) => k + 1);
  }, []);

  const handleSubmit = async (
    payload: Record<string, unknown>,
    _submitKind: TFormSubmitKind,
    files: File[],
    filesByFolderNodeId?: Record<string, File[]>
  ): Promise<void> => {
    if (formMode === 'view') {
      return;
    }
    const multiLib =
      !!filesByFolderNodeId &&
      Object.keys(filesByFolderNodeId).some((k) => filesByFolderNodeId[k].length > 0);
    if (formMode === 'create') {
      const { id, filesForAttachments } = await itemsService.addItem(
        listTitle,
        payload,
        multiLib ? Object.values(filesByFolderNodeId!).flat() : files
      );
      await uploadAttachments(
        fm,
        listTitle,
        id,
        multiLib ? [] : filesForAttachments,
        { ...payload, Id: id },
        multiLib ? filesByFolderNodeId : undefined
      );
      resetToNew();
      return;
    }
    if (formMode === 'edit' && activeItem) {
      const id = Number(activeItem.Id);
      await itemsService.updateItem(listTitle, id, payload);
      await uploadAttachments(
        fm,
        listTitle,
        id,
        multiLib ? [] : files,
        { ...payload, Id: id },
        multiLib ? filesByFolderNodeId : undefined
      );
      await loadItemById(id, 'edit');
    }
  };

  const dataLoadKind = resolveFormDataLoadingKind(fm);
  const formRootLayoutStyles = useMemo(() => resolveFormRootLayoutStyles(fm), [fm]);

  if (!dynamicContext) {
    return (
      <FormDataLoadingView kind={dataLoadKind} message="A carregar contexto…" />
    );
  }

  if (metaLoading || !fieldMeta.length) {
    return (
      <FormDataLoadingView kind={dataLoadKind} message="A carregar campos da lista…" />
    );
  }

  return (
    <div style={formRootLayoutStyles.outer}>
      <Stack
        tokens={{ childrenGap: 12 }}
        styles={{
          root: {
            marginTop: 8,
            ...formRootLayoutStyles.inner,
          } as IStyle,
        }}
      >
        {loadError && <MessageBar messageBarType={MessageBarType.error}>{loadError}</MessageBar>}
        <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
          {formMode === 'create'
            ? 'Novo registro'
            : formMode === 'view'
              ? `Visualizar #${activeItem?.Id ?? ''}`
              : `Editar #${activeItem?.Id ?? ''}`}
        </Text>
        {itemLoading ? (
          <FormDataLoadingView kind={dataLoadKind} message="A carregar item…" />
        ) : (
          <DynamicListForm
            key={formKey}
            listTitle={listTitle}
            formManager={fm}
            fieldMetadata={fieldMeta}
            formMode={formMode}
            initialItem={activeItem ?? undefined}
            itemId={activeItem ? Number(activeItem.Id) : undefined}
            dynamicContext={dynamicContext}
            userGroupTitles={userGroupTitles}
            currentUserId={currentUserId}
            onSubmit={handleSubmit}
            onDismiss={resetToNew}
            onAfterItemUpdated={async () => {
              if (!activeItem) return;
              const q = dynamicContext?.query ?? {};
              await loadItemById(Number(activeItem.Id), resolveFormModeFromQuery(q, { itemLoaded: true }));
            }}
          />
        )}
      </Stack>
    </div>
  );
};
