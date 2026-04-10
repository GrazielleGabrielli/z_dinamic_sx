import * as React from 'react';
import { useState, useEffect, useLayoutEffect, useMemo, useCallback, useRef } from 'react';
import { flushSync } from 'react-dom';
import {
  Stack,
  Text,
  TextField,
  Toggle,
  Dropdown,
  IDropdownOption,
  DatePicker,
  PrimaryButton,
  DefaultButton,
  IconButton,
  MessageBar,
  MessageBarType,
  Label,
  Link,
  Icon,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type {
  IFormManagerConfig,
  IFormFieldConfig,
  IFormCustomButtonConfig,
  TFormButtonAction,
  TFormCustomButtonOperation,
  TFormManagerFormMode,
  TFormSubmitKind,
  TFormRule,
  TFormSubmitLoadingUiKind,
  TFormAttachmentFilePreviewKind,
} from '../../core/config/types/formManager';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_BANNER_INTERNAL_PREFIX,
  FORM_OCULTOS_STEP_ID,
  FORM_FIXOS_STEP_ID,
  FORM_BUILTIN_HISTORY_BUTTON_ID,
  isFormBannerFieldConfig,
  resolveBannerPlacement,
  resolveBannerWidthPercent,
  resolveBannerHeightPercent,
  resolveFixedPlacement,
  resolveChromePositionMode,
} from '../../core/config/types/formManager';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { isDynamicToken } from '../../core/dynamicTokens';
import {
  buildFormDerivedState,
  collectFormValidationErrors,
  filterValidationErrorsToStepFields,
  pickRequiredStyleStepErrors,
  evaluateCondition,
  evaluateFormValueExpression,
  getDefaultValuesFromRules,
  shouldShowCustomButton,
  shouldShowBuiltinHistoryButton,
  areAllRequiredFieldsFilled,
  isAttachmentFolderUploaderVisible,
  buildFormFieldLabelMap,
  formatValidationSummaryForForm,
  type IFormAttachmentFolderUrlContext,
  type IFormRuleRuntimeContext,
  type IFormValidationAttachmentContext,
} from '../../core/formManager/formRuleEngine';
import { formValuesToSharePointPayload } from '../../core/formManager/formSharePointValues';
import { FormStepNavigation, FormStepPrevNextNav } from './FormStepLayoutUi';
import { FormAttachmentUploader } from './FormAttachmentUploader';
import { runAsyncFormValidations } from '../../core/formManager/formAsyncValidation';
import { interpolateFormButtonRedirectUrl } from '../../core/formManager/formButtonRedirectUrl';
import { appendFormActionLogEntry } from '../../core/formManager/formActionLog';
import { parseAttachmentUiRule } from '../../core/formManager/formManagerVisualModel';
import { ItemsService, UsersService } from '../../../../services';
import { getSP } from '../../../../services/core/sp';
import {
  isFormAttachmentLibraryRuntime,
  uploadFilesToAttachmentLibrary,
  uploadFilesToAttachmentLibraryByFolderNodes,
  loadLibraryAttachmentRowsForMainItem,
  libraryFileRowBelongsToFolderNode,
} from '../../core/formManager/formAttachmentLibrary';
import {
  collectFolderAttachmentLimitErrors,
  filterFolderLimitErrorsToStep,
  flattenFolderTreeNodes,
  FOLDER_ATTACHMENT_LIMIT_ERROR_PREFIX,
  treeHasPerStepFolderUploaders,
} from '../../core/formManager/attachmentFolderTree';
import { FormSubmitLoadingChrome, resolveSubmitLoadingKind } from './FormLoadingUi';
import { FormItemHistoryUi } from './FormItemHistoryUi';
import { LinkedChildFormsBlock } from './LinkedChildFormsBlock';
import { attachmentFileKindIconName } from './attachmentFileKindIcon';
import { stepVisibleInFormMode } from '../../core/formManager/stepFormMode';
import {
  linkedChildFormAsManagerConfig,
  loadLinkedChildRows,
  syncAllLinkedChildLists,
  type ILinkedChildRowState,
} from '../../core/formManager/formLinkedChildSync';
import { FieldsService } from '../../../../services';

export interface IDynamicListFormProps {
  listTitle: string;
  formManager: IFormManagerConfig;
  fieldMetadata: IFieldMetadata[];
  formMode: TFormManagerFormMode;
  initialItem?: Record<string, unknown> | null;
  itemId?: number;
  dynamicContext: IDynamicContext;
  userGroupTitles: string[];
  currentUserId: number;
  onSubmit: (
    payload: Record<string, unknown>,
    submitKind: TFormSubmitKind,
    pendingFiles: File[],
    pendingFilesByFolderNodeId?: Record<string, File[]>
  ) => Promise<number | undefined>;
  onDismiss: () => void;
  /** Chamado após botão «Atualizar» personalizado gravar com sucesso (ex.: recarregar item). */
  onAfterItemUpdated?: () => void | Promise<void>;
}

async function uploadListItemAttachments(
  listTitle: string,
  itemId: number,
  files: File[],
  formManager: IFormManagerConfig,
  itemFieldValues: Record<string, unknown>,
  filesByFolderNodeId?: Record<string, File[]>
): Promise<void> {
  const hasFolderBuckets =
    !!filesByFolderNodeId && Object.keys(filesByFolderNodeId).some((k) => filesByFolderNodeId[k].length > 0);
  if (!files.length && !hasFolderBuckets) return;
  if (isFormAttachmentLibraryRuntime(formManager)) {
    const lib = formManager.attachmentLibrary!;
    const iv = { ...itemFieldValues, Id: itemId };
    if (hasFolderBuckets) {
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

type IServerAttachmentRow = { fileName: string; fileUrl: string; fileRef?: string };

function normalizeSharePointAttachmentFiles(raw: unknown): unknown[] {
  if (Array.isArray(raw)) return raw;
  if (raw && typeof raw === 'object') {
    const o = raw as Record<string, unknown>;
    if (Array.isArray(o.value)) return o.value;
    const d = o.d as Record<string, unknown> | undefined;
    if (d && Array.isArray(d.results)) return d.results as unknown[];
  }
  return [];
}

function mapServerAttachments(rows: unknown[]): IServerAttachmentRow[] {
  const out: IServerAttachmentRow[] = [];
  for (let i = 0; i < rows.length; i++) {
    const a = rows[i];
    if (!a || typeof a !== 'object') continue;
    const r = a as Record<string, unknown>;
    const fn = r.FileName ?? r.fileName;
    if (typeof fn !== 'string' || !fn.trim()) continue;
    const sr = r.ServerRelativeUrl ?? r.serverRelativeUrl;
    let fileUrl = '';
    if (typeof sr === 'string' && sr.trim()) {
      const path = sr.trim();
      fileUrl = /^https?:\/\//i.test(path)
        ? path
        : `${typeof window !== 'undefined' ? window.location.origin : ''}${
            path.startsWith('/') ? '' : '/'
          }${path}`;
    }
    out.push({ fileName: fn.trim(), fileUrl });
  }
  return out;
}

function itemToFormValues(
  item: Record<string, unknown> | undefined,
  names: string[]
): Record<string, unknown> {
  const out: Record<string, unknown> = {};
  if (!item) return out;
  for (let i = 0; i < names.length; i++) {
    const n = names[i];
    out[n] = item[n];
  }
  return out;
}

function formatJoinedFieldValue(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v !== null && 'Title' in (v as object)) {
    return String((v as Record<string, unknown>).Title ?? '');
  }
  return String(v);
}

type IFormButtonFieldOverlay = {
  show: Set<string>;
  hide: Set<string>;
  showOnStepId?: Record<string, string>;
};

function reduceCustomButtonActions(
  actions: TFormButtonAction[],
  startValues: Record<string, unknown>,
  dynamicContext: IDynamicContext,
  baseOverlay: IFormButtonFieldOverlay,
  attachmentFolderUrl?: IFormAttachmentFolderUrlContext
): { mergedValues: Record<string, unknown>; mergedOverlay: IFormButtonFieldOverlay } {
  let next = { ...startValues };
  const mergedOverlay: IFormButtonFieldOverlay = {
    show: cloneStringSet(baseOverlay.show),
    hide: cloneStringSet(baseOverlay.hide),
    ...(baseOverlay.showOnStepId && Object.keys(baseOverlay.showOnStepId).length > 0
      ? { showOnStepId: { ...baseOverlay.showOnStepId } }
      : {}),
  };
  for (let i = 0; i < actions.length; i++) {
    const a = actions[i];
    if (a.when && !evaluateCondition(a.when, next, dynamicContext)) {
      continue;
    }
    if (a.kind === 'setFieldValue') {
      const tplRaw = String(a.valueTemplate ?? '');
      const trimmed = tplRaw.trim();
      let useExpr = trimmed.startsWith('str:') || trimmed.startsWith('attfolder:');
      if (!useExpr) useExpr = isDynamicToken(trimmed);
      const raw = useExpr
        ? evaluateFormValueExpression(tplRaw, next, dynamicContext, attachmentFolderUrl)
        : tplRaw;
      next = { ...next, [a.field]: raw };
    } else if (a.kind === 'joinFields') {
      const tpl = (a.valueTemplate ?? '').trim();
      if (tpl.length > 0) {
        const rawTpl = a.valueTemplate ?? '';
        const interpolated = rawTpl.replace(/\{\{([^}]+)\}\}/g, (_, raw: string) => {
          const name = String(raw).trim();
          return formatJoinedFieldValue(next[name]);
        });
        next = { ...next, [a.targetField]: interpolated };
      } else {
        const parts = a.sourceFields.map((f) => formatJoinedFieldValue(next[f]));
        next = { ...next, [a.targetField]: parts.join(a.separator) };
      }
    } else if (a.kind === 'showFields') {
      const sid = typeof a.displayOnStepId === 'string' ? a.displayOnStepId.trim() : '';
      for (let j = 0; j < a.fields.length; j++) {
        const fn = a.fields[j];
        mergedOverlay.show.add(fn);
        if (sid) {
          if (!mergedOverlay.showOnStepId) mergedOverlay.showOnStepId = {};
          mergedOverlay.showOnStepId[fn] = sid;
        }
      }
    } else if (a.kind === 'hideFields') {
      for (let j = 0; j < a.fields.length; j++) {
        const fn = a.fields[j];
        mergedOverlay.hide.add(fn);
        if (mergedOverlay.showOnStepId && mergedOverlay.showOnStepId[fn]) {
          delete mergedOverlay.showOnStepId[fn];
        }
      }
    }
  }
  return { mergedValues: next, mergedOverlay };
}

function cloneStringSet(s: Set<string>): Set<string> {
  const n = new Set<string>();
  s.forEach((x) => {
    n.add(x);
  });
  return n;
}

const REQ_EMPTY_BORDER = '#a4262c';

function isValueEmptyForRequired(v: unknown, mappedType: string): boolean {
  if (mappedType === 'boolean') {
    return v === undefined || v === null;
  }
  if (mappedType === 'url') {
    if (v === null || v === undefined) return true;
    if (typeof v === 'object' && v !== null && 'Url' in v) {
      return String((v as Record<string, unknown>).Url ?? '').trim() === '';
    }
    if (typeof v === 'string') return v.trim() === '';
    return false;
  }
  if (v === null || v === undefined) return true;
  if (typeof v === 'string' && v.trim() === '') return true;
  if (Array.isArray(v) && v.length === 0) return true;
  if (typeof v === 'object' && v !== null && 'Id' in (v as object)) {
    const id = (v as Record<string, unknown>).Id;
    if (id === null || id === undefined || id === '') return true;
  }
  return false;
}

function stylesTextFieldRequiredEmpty(active: boolean): { fieldGroup?: Record<string, string | number> } | undefined {
  if (!active) return undefined;
  return {
    fieldGroup: {
      borderColor: REQ_EMPTY_BORDER,
      borderWidth: 1,
      borderStyle: 'solid',
      borderRadius: 2,
    },
  };
}

function listRequiredEmptyErrorsInStep(
  stepFieldList: Set<string>,
  values: Record<string, unknown>,
  metaByName: Map<string, IFieldMetadata>,
  fieldVisible: (n: string) => boolean
): Record<string, string> {
  const out: Record<string, string> = {};
  stepFieldList.forEach((name) => {
    if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) return;
    if (name.indexOf(FORM_BANNER_INTERNAL_PREFIX) === 0) return;
    if (!fieldVisible(name)) return;
    const m = metaByName.get(name);
    if (!m?.Required) return;
    if (!isValueEmptyForRequired(values[name], m.MappedType)) return;
    out[name] = 'Obrigatório.';
  });
  return out;
}

function lookupIdFromValue(v: unknown): number | undefined {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'object' && v !== null && 'Id' in v) {
    const id = (v as Record<string, unknown>).Id;
    if (typeof id === 'number') return id;
  }
  return undefined;
}

function normalizeIdTitleArray(v: unknown): { Id: number; Title: string }[] {
  if (v === null || v === undefined) return [];
  if (Array.isArray(v)) {
    const out: { Id: number; Title: string }[] = [];
    for (let i = 0; i < v.length; i++) {
      const x = v[i];
      const id = lookupIdFromValue(x);
      if (id === undefined) continue;
      let title = '';
      if (typeof x === 'object' && x !== null && 'Title' in x) {
        title = String((x as Record<string, unknown>).Title ?? '');
      }
      out.push({ Id: id, Title: title });
    }
    return out;
  }
  if (typeof v === 'object' && v !== null && 'results' in v) {
    const r = (v as { results?: unknown }).results;
    if (Array.isArray(r)) return normalizeIdTitleArray(r);
  }
  const id = lookupIdFromValue(v);
  return id !== undefined ? [{ Id: id, Title: '' }] : [];
}

function mergeOptionsForIds(
  opts: IDropdownOption[],
  entries: { id: number; label: string }[]
): IDropdownOption[] {
  const keys = new Set(opts.map((o) => String(o.key)));
  let out = opts;
  for (let i = 0; i < entries.length; i++) {
    const k = String(entries[i].id);
    if (keys.has(k)) continue;
    keys.add(k);
    out = [...out, { key: k, text: entries[i].label.trim() || `#${entries[i].id}` }];
  }
  return out;
}

function userTitleFromValue(v: unknown): string {
  if (typeof v === 'object' && v !== null && 'Title' in v) {
    return String((v as Record<string, unknown>).Title ?? '');
  }
  return '';
}

function parseUrlFieldValue(v: unknown): { Url: string; Description: string } {
  if (v === null || v === undefined) return { Url: '', Description: '' };
  if (typeof v === 'object' && v !== null && 'Url' in v) {
    const o = v as Record<string, unknown>;
    return { Url: String(o.Url ?? ''), Description: String(o.Description ?? '') };
  }
  const s = String(v);
  const comma = s.indexOf(',');
  if (comma !== -1) {
    return { Url: s.slice(0, comma).trim(), Description: s.slice(comma + 1).trim() };
  }
  return { Url: s, Description: '' };
}

function dropdownReqStyles(showReq: boolean | undefined) {
  return showReq
    ? {
        dropdown: {
          borderColor: REQ_EMPTY_BORDER,
          borderWidth: 1,
          borderStyle: 'solid',
          borderRadius: 2,
        },
      }
    : undefined;
}

interface IFormChromeZoneProps {
  zone: 'top' | 'bottom';
  fields: IFormFieldConfig[];
  renderField: (fc: IFormFieldConfig) => React.ReactNode;
  layoutDeps: unknown;
}

const FormChromeZone: React.FC<IFormChromeZoneProps> = ({ zone, fields, renderField, layoutDeps }) => {
  const rootRef = useRef<HTMLDivElement>(null);
  const [edgeOffsets, setEdgeOffsets] = useState<number[]>(() => fields.map(() => 0));
  const [minH, setMinH] = useState(0);

  useLayoutEffect(() => {
    const root = rootRef.current;
    if (!root) return;
    const children = Array.from(root.children) as HTMLElement[];
    const next: number[] = fields.map(() => 0);
    let total = 0;
    if (zone === 'top') {
      let acc = 0;
      for (let i = 0; i < fields.length; i++) {
        const h = children[i]?.offsetHeight ?? 0;
        if (resolveChromePositionMode(fields[i]) === 'absolute') next[i] = acc;
        acc += h;
      }
      total = acc;
    } else {
      let acc = 0;
      for (let i = fields.length - 1; i >= 0; i--) {
        const h = children[i]?.offsetHeight ?? 0;
        if (resolveChromePositionMode(fields[i]) === 'absolute') next[i] = acc;
        acc += h;
      }
      total = children.reduce((s, el) => s + (el?.offsetHeight ?? 0), 0);
    }
    setEdgeOffsets((prev) => {
      if (prev.length === next.length && next.every((v, idx) => v === next[idx])) return prev;
      return next;
    });
    setMinH(total);
  }, [fields, zone, layoutDeps]);

  if (!fields.length) return null;

  return (
    <Stack
      tokens={{ childrenGap: 0 }}
      styles={{
        root: {
          position: 'relative',
          width: '100%',
          minHeight: minH || undefined,
          marginBottom: zone === 'top' ? 8 : 0,
          marginTop: zone === 'bottom' ? 8 : 0,
          borderBottom: zone === 'top' ? '1px solid #edebe9' : undefined,
          borderTop: zone === 'bottom' ? '1px solid #edebe9' : undefined,
        },
      }}
    >
      <div ref={rootRef}>
        {fields.map((fc, i) => {
          const mode = resolveChromePositionMode(fc);
          const base: React.CSSProperties = {
            width: '100%',
            boxSizing: 'border-box',
          };
          let style: React.CSSProperties = { ...base };
          if (mode === 'flow') {
            style = {
              ...base,
              position: 'relative',
              paddingBottom: zone === 'top' ? 8 : 0,
              paddingTop: zone === 'bottom' ? 8 : 0,
            };
          } else if (mode === 'sticky') {
            style = {
              ...base,
              position: 'sticky',
              ...(zone === 'top' ? { top: 0 } : { bottom: 0 }),
              zIndex: 6,
              background: '#ffffff',
            };
          } else {
            style = {
              ...base,
              position: 'absolute',
              left: 0,
              right: 0,
              ...(zone === 'top' ? { top: edgeOffsets[i] ?? 0 } : { bottom: edgeOffsets[i] ?? 0 }),
              zIndex: 5,
              background: '#ffffff',
            };
          }
          return (
            <div key={fc.internalName} style={style}>
              {renderField(fc)}
            </div>
          );
        })}
      </div>
    </Stack>
  );
};

export const DynamicListForm: React.FC<IDynamicListFormProps> = ({
  listTitle,
  formManager,
  fieldMetadata,
  formMode,
  initialItem,
  itemId,
  dynamicContext,
  userGroupTitles,
  currentUserId,
  onSubmit,
  onDismiss,
  onAfterItemUpdated,
}) => {
  const fieldConfigs: IFormFieldConfig[] =
    formManager.fields.length > 0
      ? formManager.fields
      : fieldMetadata
          .filter((f) => !f.Hidden && !f.ReadOnlyField && f.InternalName !== 'Id')
          .map((f) => ({ internalName: f.InternalName, sectionId: FORM_OCULTOS_STEP_ID }));
  const names = useMemo(
    () =>
      fieldConfigs
        .filter((f) => f.internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL && !isFormBannerFieldConfig(f))
        .map((f) => f.internalName),
    [fieldConfigs]
  );
  const ocultosNullFieldNames = useMemo(
    () =>
      fieldConfigs
        .filter(
          (f) =>
            f.sectionId === FORM_OCULTOS_STEP_ID &&
            f.internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL &&
            !isFormBannerFieldConfig(f)
        )
        .map((f) => f.internalName),
    [fieldConfigs]
  );
  const metaByName = useMemo(() => new Map(fieldMetadata.map((f) => [f.InternalName, f])), [fieldMetadata]);
  const fieldLabelByName = useMemo(
    () => buildFormFieldLabelMap(fieldConfigs, metaByName),
    [fieldConfigs, metaByName]
  );

  const attachmentAllowedExtensions = useMemo(
    () => parseAttachmentUiRule(formManager.rules ?? []).allowedFileExtensions ?? [],
    [formManager.rules]
  );
  const attachmentPreviewKind: TFormAttachmentFilePreviewKind =
    formManager.attachmentFilePreview ?? 'nameAndSize';

  const multiFolderAttachmentMode = useMemo(() => {
    if (formManager.attachmentStorageKind !== 'documentLibrary') return false;
    const t = formManager.attachmentLibrary?.folderTree;
    if (!t?.length) return false;
    return treeHasPerStepFolderUploaders(t);
  }, [formManager.attachmentStorageKind, formManager.attachmentLibrary?.folderTree]);

  const [values, setValues] = useState<Record<string, unknown>>(() =>
    itemToFormValues(initialItem ?? undefined, names)
  );
  const [submitUi, setSubmitUi] = useState<TFormSubmitLoadingUiKind | null>(null);
  const submitting = submitUi !== null;
  const [formError, setFormError] = useState<string | undefined>(undefined);
  const [localErrors, setLocalErrors] = useState<Record<string, string>>({});
  const [lookupOptions, setLookupOptions] = useState<Record<string, IDropdownOption[]>>({});
  const [pendingFiles, setPendingFiles] = useState<File[]>([]);
  const [pendingFilesByFolder, setPendingFilesByFolder] = useState<Record<string, File[]>>({});
  const [attachmentCount, setAttachmentCount] = useState(0);
  const [serverAttachments, setServerAttachments] = useState<IServerAttachmentRow[]>([]);
  const prevByTriggerRef = useRef<Record<string, unknown>>({});
  const [buttonOverlay, setButtonOverlay] = useState<IFormButtonFieldOverlay>(() => ({
    show: new Set<string>(),
    hide: new Set<string>(),
  }));
  const [attachmentLibRootServerRelative, setAttachmentLibRootServerRelative] = useState<string | undefined>(
    undefined
  );

  const authorId = useMemo(() => {
    const a = initialItem?.AuthorId ?? initialItem?.Author;
    if (typeof a === 'number') return a;
    if (a && typeof a === 'object' && 'Id' in (a as object)) return (a as { Id: number }).Id;
    return undefined;
  }, [initialItem]);

  const itemsService = useMemo(() => new ItemsService(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);
  const usersService = useMemo(() => new UsersService(), []);
  const [siteUserOptions, setSiteUserOptions] = useState<IDropdownOption[]>([{ key: '', text: '—' }]);

  const linkedConfigsSorted = useMemo(() => {
    const raw = formManager.linkedChildForms ?? [];
    return raw
      .filter((c) => c.listTitle.trim() && c.parentLookupFieldInternalName.trim())
      .slice()
      .sort((a, b) => (a.order ?? 0) - (b.order ?? 0));
  }, [formManager.linkedChildForms]);

  const [linkedMetaById, setLinkedMetaById] = useState<Record<string, IFieldMetadata[]>>({});
  const [linkedRowsById, setLinkedRowsById] = useState<Record<string, ILinkedChildRowState[]>>({});
  const [linkedBaselineById, setLinkedBaselineById] = useState<Record<string, number[]>>({});
  const [linkedLoadErrById, setLinkedLoadErrById] = useState<Record<string, string>>({});
  const [linkedLoadingById, setLinkedLoadingById] = useState<Record<string, boolean>>({});
  const [linkedRowErrorsById, setLinkedRowErrorsById] = useState<Record<string, Record<string, string>[]>>({});

  const builtinHistoryButtonConfig = useMemo((): IFormCustomButtonConfig => {
    const label = (formManager.historyButtonLabel ?? 'Histórico').trim() || 'Histórico';
    const sub = formManager.historyPanelSubtitle?.trim();
    return {
      id: FORM_BUILTIN_HISTORY_BUTTON_ID,
      label,
      shortDescription: sub || undefined,
      appearance: 'default',
      behavior: 'actionsOnly',
      operation: 'history',
      actions: [],
    };
  }, [formManager.historyButtonLabel, formManager.historyPanelSubtitle]);

  useEffect(() => {
    setValues(itemToFormValues(initialItem ?? undefined, names));
    setButtonOverlay({ show: new Set<string>(), hide: new Set<string>() });
  }, [initialItem, names]);

  useEffect(() => {
    let cancelled = false;
    void usersService
      .getSiteUsers()
      .then((users) => {
        if (cancelled) return;
        const sorted = [...users].sort((a, b) =>
          (a.Title || a.Email || '').localeCompare(b.Title || b.Email || '', undefined, { sensitivity: 'base' })
        );
        setSiteUserOptions([
          { key: '', text: '—' },
          ...sorted.map((u) => ({
            key: String(u.Id),
            text: (u.Title || u.Email || u.LoginName || '').trim() || `#${u.Id}`,
          })),
        ]);
      })
      .catch(() => {
        if (!cancelled) setSiteUserOptions([{ key: '', text: '—' }]);
      });
    return (): void => {
      cancelled = true;
    };
  }, [usersService]);

  useEffect(() => {
    if (!linkedConfigsSorted.length) {
      setLinkedMetaById({});
      setLinkedLoadErrById({});
      return;
    }
    let cancel = false;
    void (async (): Promise<void> => {
      const next: Record<string, IFieldMetadata[]> = {};
      const err: Record<string, string> = {};
      const load: Record<string, boolean> = {};
      for (let i = 0; i < linkedConfigsSorted.length; i++) {
        load[linkedConfigsSorted[i].id] = true;
      }
      setLinkedLoadingById((prev) => ({ ...prev, ...load }));
      for (let i = 0; i < linkedConfigsSorted.length; i++) {
        if (cancel) return;
        const c = linkedConfigsSorted[i];
        try {
          const f = await fieldsService.getVisibleFields(c.listTitle.trim());
          if (!cancel) next[c.id] = f;
        } catch (e) {
          if (!cancel) err[c.id] = e instanceof Error ? e.message : String(e);
        } finally {
          if (!cancel) {
            setLinkedLoadingById((prev) => {
              const p = { ...prev };
              delete p[c.id];
              return p;
            });
          }
        }
      }
      if (!cancel) {
        setLinkedMetaById(next);
        setLinkedLoadErrById(err);
      }
    })();
    return (): void => {
      cancel = true;
    };
  }, [linkedConfigsSorted, fieldsService]);

  useEffect(() => {
    if (formMode !== 'create') return;
    setLinkedRowsById((prev) => {
      const next = { ...prev };
      for (let i = 0; i < linkedConfigsSorted.length; i++) {
        const c = linkedConfigsSorted[i];
        if (next[c.id] !== undefined) continue;
        const min = c.minRows ?? 0;
        const count = Math.max(min, 1);
        next[c.id] = Array.from({ length: count }, (_, j) => ({
          localKey: `new_${c.id}_${j}_${Date.now()}_${Math.random().toString(36).slice(2)}`,
          values: {},
        }));
      }
      return next;
    });
    setLinkedBaselineById({});
  }, [formMode, linkedConfigsSorted]);

  useEffect(() => {
    if (formMode === 'create' || itemId === undefined || itemId === null) return;
    let cancel = false;
    void (async (): Promise<void> => {
      const nextRows: Record<string, ILinkedChildRowState[]> = {};
      const nextBase: Record<string, number[]> = {};
      for (let i = 0; i < linkedConfigsSorted.length; i++) {
        if (cancel) return;
        const c = linkedConfigsSorted[i];
        const meta = linkedMetaById[c.id];
        if (!meta?.length) continue;
        try {
          const rows = await loadLinkedChildRows(itemsService, c, itemId, meta);
          nextRows[c.id] = rows;
          nextBase[c.id] = rows
            .map((r) => r.sharePointId)
            .filter((x): x is number => typeof x === 'number' && isFinite(x));
        } catch {
          nextRows[c.id] = [];
          nextBase[c.id] = [];
        }
      }
      if (!cancel) {
        setLinkedRowsById(nextRows);
        setLinkedBaselineById(nextBase);
      }
    })();
    return (): void => {
      cancel = true;
    };
  }, [formMode, itemId, linkedConfigsSorted, linkedMetaById, itemsService]);

  useEffect(() => {
    if (formMode !== 'create') return;
    setValues((prev) => {
      const merged = getDefaultValuesFromRules(formManager, prev, dynamicContext);
      return merged;
    });
  }, [formManager, formMode, dynamicContext]);

  useEffect(() => {
    if (!isFormAttachmentLibraryRuntime(formManager)) {
      setAttachmentLibRootServerRelative(undefined);
      return;
    }
    const title = formManager.attachmentLibrary!.libraryTitle!.trim();
    let cancelled = false;
    (async () => {
      try {
        const sp = getSP();
        const list = sp.web.lists.getByTitle(title);
        const rf = await list.rootFolder.select('ServerRelativeUrl')();
        const u = (rf as { ServerRelativeUrl?: string }).ServerRelativeUrl;
        if (!cancelled) {
          setAttachmentLibRootServerRelative(typeof u === 'string' && u.trim() ? u.trim() : undefined);
        }
      } catch {
        if (!cancelled) setAttachmentLibRootServerRelative(undefined);
      }
    })().catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [formManager.attachmentStorageKind, formManager.attachmentLibrary?.libraryTitle]);

  const attachmentFolderUrl = useMemo(
    (): IFormAttachmentFolderUrlContext => ({
      libraryRootServerRelativeUrl: attachmentLibRootServerRelative,
      itemId,
      folderTree: formManager.attachmentLibrary?.folderTree,
    }),
    [attachmentLibRootServerRelative, itemId, formManager.attachmentLibrary?.folderTree]
  );

  const runtimeCtx = useCallback(
    (submitKind?: TFormSubmitKind): IFormRuleRuntimeContext => ({
      formMode,
      values,
      submitKind,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      attachmentFolderUrl,
    }),
    [
      formMode,
      values,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      attachmentFolderUrl,
    ]
  );

  const derived = useMemo(
    () =>
      buildFormDerivedState(formManager, fieldConfigs, runtimeCtx(), {
        show: buttonOverlay.show,
        hide: buttonOverlay.hide,
      }),
    [
      formManager,
      fieldConfigs,
      runtimeCtx,
      values,
      formMode,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      attachmentFolderUrl,
      buttonOverlay,
    ]
  );

  const flatPendingFiles = useMemo(() => {
    if (multiFolderAttachmentMode) return Object.values(pendingFilesByFolder).flat();
    return pendingFiles;
  }, [multiFolderAttachmentMode, pendingFilesByFolder, pendingFiles]);

  const allRequiredFilled = useMemo(
    () =>
      areAllRequiredFieldsFilled(
        formManager,
        fieldConfigs,
        runtimeCtx(),
        metaByName,
        { show: buttonOverlay.show, hide: buttonOverlay.hide },
        {
          attachmentCount,
          pendingFiles: flatPendingFiles.map((f) => ({
            size: f.size,
            type: f.type || 'application/octet-stream',
            name: f.name,
          })),
        }
      ),
    [
      formManager,
      fieldConfigs,
      runtimeCtx,
      values,
      formMode,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      attachmentFolderUrl,
      buttonOverlay,
      attachmentCount,
      flatPendingFiles,
      metaByName,
    ]
  );

  const clearRules = useMemo(
    () => formManager.rules.filter((r): r is Extract<TFormRule, { action: 'clearFields' }> => r.action === 'clearFields'),
    [formManager.rules]
  );

  useEffect(() => {
    for (let i = 0; i < clearRules.length; i++) {
      const rule = clearRules[i];
      if (!rule.triggerField) continue;
      const cur = values[rule.triggerField];
      const prev = prevByTriggerRef.current[rule.triggerField];
      if (prev !== undefined && prev !== cur) {
        setValues((v) => {
          const next = { ...v };
          for (let j = 0; j < rule.fields.length; j++) next[rule.fields[j]] = null;
          return next;
        });
      }
      prevByTriggerRef.current[rule.triggerField] = cur;
    }
  }, [values, clearRules]);

  const loadLookupOptions = useCallback(
    async (fieldName: string, odataFilter?: string): Promise<void> => {
      const m = metaByName.get(fieldName);
      if (!m?.LookupList) return;
      try {
        const filter = odataFilter?.trim() ? odataFilter : undefined;
        const rows = await itemsService.getItems<Record<string, unknown>>(m.LookupList, {
          select: ['Id', m.LookupField || 'Title'],
          filter,
          top: 200,
        });
        const lf = m.LookupField || 'Title';
        const opts: IDropdownOption[] = [
          { key: '', text: '—' },
          ...rows.map((row) => ({
            key: String(row.Id),
            text: String(row[lf] ?? row.Id),
          })),
        ];
        setLookupOptions((o) => ({ ...o, [fieldName]: opts }));
      } catch {
        setLookupOptions((o) => ({ ...o, [fieldName]: [] }));
      }
    },
    [itemsService, metaByName]
  );

  const lookupFetchKey = useMemo(() => {
    const parts: string[] = [];
    for (let i = 0; i < fieldConfigs.length; i++) {
      const fn = fieldConfigs[i].internalName;
      const m = metaByName.get(fn);
      if (m?.MappedType !== 'lookup' && m?.MappedType !== 'lookupmulti') continue;
      const listId = String(m.LookupList ?? '');
      const disp = String(m.LookupField || 'Title');
      const lf = derived.lookupFilters[fn];
      if (lf) {
        const pid = lookupIdFromValue(values[lf.parentField]);
        parts.push(
          `${fn}\t${listId}\t${disp}\t${lf.parentField}\t${lf.odataFilterTemplate}\t${pid === undefined ? '' : String(pid)}`
        );
      } else {
        parts.push(`${fn}\t${listId}\t${disp}\t`);
      }
    }
    parts.sort();
    return parts.join('\n');
  }, [fieldConfigs, metaByName, derived.lookupFilters, values]);

  useEffect(() => {
    let cancelled = false;
    void (async (): Promise<void> => {
      for (let i = 0; i < fieldConfigs.length; i++) {
        if (cancelled) return;
        const fn = fieldConfigs[i].internalName;
        const m = metaByName.get(fn);
        if (m?.MappedType === 'lookup' || m?.MappedType === 'lookupmulti') {
          const lf = derived.lookupFilters[fn];
          let filter: string | undefined;
          if (lf) {
            const pid = lookupIdFromValue(values[lf.parentField]);
            if (pid !== undefined) filter = lf.odataFilterTemplate.split('{parent}').join(String(pid));
          }
          await loadLookupOptions(fn, filter);
        }
      }
    })();
    return (): void => {
      cancelled = true;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- fieldConfigs, metaByName, derived, values entram via lookupFetchKey (conteúdo estável).
  }, [lookupFetchKey, loadLookupOptions]);

  useEffect(() => {
    if (formMode === 'create' || !itemId) {
      setAttachmentCount(0);
      setServerAttachments([]);
      return;
    }
    const libMode = isFormAttachmentLibraryRuntime(formManager);
    let cancelled = false;
    void (async (): Promise<void> => {
      try {
        if (libMode) {
          const lib = formManager.attachmentLibrary!;
          const mapped = await loadLibraryAttachmentRowsForMainItem(
            lib.libraryTitle!,
            lib.sourceListLookupFieldInternalName!,
            itemId
          );
          if (!cancelled) {
            setAttachmentCount(mapped.length);
            setServerAttachments(mapped);
          }
          return;
        }
        const sp = getSP();
        const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(listTitle);
        const list = isGuid ? sp.web.lists.getById(listTitle) : sp.web.lists.getByTitle(listTitle);
        const item = list.items.getById(itemId) as unknown as { attachmentFiles(): Promise<unknown> };
        const raw = await item.attachmentFiles();
        const rows = normalizeSharePointAttachmentFiles(raw);
        const mapped = mapServerAttachments(rows);
        if (!cancelled) {
          setAttachmentCount(mapped.length);
          setServerAttachments(mapped);
        }
      } catch {
        if (!cancelled) {
          setAttachmentCount(0);
          setServerAttachments([]);
        }
      }
    })();
    return (): void => {
      cancelled = true;
    };
  }, [listTitle, itemId, formMode, initialItem, formManager]);

  const updateField = (name: string, v: unknown): void => {
    setValues((prev) => ({ ...prev, [name]: v }));
  };

  const reloadLinkedRowsForParent = useCallback(
    async (parentId: number) => {
      const nextRows: Record<string, ILinkedChildRowState[]> = {};
      const nextBase: Record<string, number[]> = {};
      for (let i = 0; i < linkedConfigsSorted.length; i++) {
        const c = linkedConfigsSorted[i];
        const meta = linkedMetaById[c.id];
        if (!meta?.length) continue;
        try {
          const rows = await loadLinkedChildRows(itemsService, c, parentId, meta);
          nextRows[c.id] = rows;
          nextBase[c.id] = rows
            .map((r) => r.sharePointId)
            .filter((x): x is number => typeof x === 'number' && isFinite(x));
        } catch {
          nextRows[c.id] = [];
          nextBase[c.id] = [];
        }
      }
      setLinkedRowsById(nextRows);
      setLinkedBaselineById(nextBase);
    },
    [linkedConfigsSorted, linkedMetaById, itemsService]
  );

  const performLinkedSync = useCallback(
    async (parentId: number) => {
      if (!linkedConfigsSorted.length) return;
      await syncAllLinkedChildLists(
        itemsService,
        linkedConfigsSorted,
        parentId,
        linkedRowsById,
        linkedMetaById,
        linkedBaselineById
      );
      await reloadLinkedRowsForParent(parentId);
      setLinkedRowErrorsById({});
    },
    [
      linkedConfigsSorted,
      itemsService,
      linkedRowsById,
      linkedMetaById,
      linkedBaselineById,
      reloadLinkedRowsForParent,
    ]
  );

  const validate = async (
    submitKind: TFormSubmitKind,
    opts?: {
      values?: Record<string, unknown>;
      buttonOverlay?: IFormButtonFieldOverlay;
    }
  ): Promise<Record<string, string> | undefined> => {
    const vals = opts?.values ?? values;
    const ov = opts?.buttonOverlay ?? buttonOverlay;
    const att: IFormValidationAttachmentContext = {
      attachmentCount,
      pendingFiles: flatPendingFiles.map((f) => ({
        size: f.size,
        type: f.type || 'application/octet-stream',
        name: f.name,
      })),
    };
    const ctx: IFormRuleRuntimeContext = {
      formMode,
      values: vals,
      submitKind,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      attachmentFolderUrl,
    };
    const sync = collectFormValidationErrors(formManager, fieldConfigs, ctx, att, {
      show: ov.show,
      hide: ov.hide,
    });
    let mergedErr = { ...sync };
    if (submitKind !== 'draft' && multiFolderAttachmentMode && isFormAttachmentLibraryRuntime(formManager)) {
      const t = formManager.attachmentLibrary?.folderTree;
      if (t?.length) {
        mergedErr = {
          ...mergedErr,
          ...collectFolderAttachmentLimitErrors(t, {
            pendingByFolder: pendingFilesByFolder,
            libraryCountByNodeId: (nodeId) => {
              if (!itemId) return 0;
              let c = 0;
              for (let i = 0; i < serverAttachments.length; i++) {
                const row = serverAttachments[i];
                const fr = row.fileRef;
                if (typeof fr !== 'string' || !fr.trim()) continue;
                if (libraryFileRowBelongsToFolderNode(fr.trim(), nodeId, t, itemId, vals)) c++;
              }
              return c;
            },
            isFolderUploaderVisible: (n) =>
              isAttachmentFolderUploaderVisible(n, {
                formMode,
                values: vals,
                submitKind,
                userGroupTitles,
                currentUserId,
                authorId,
                dynamicContext,
                attachmentFolderUrl,
              }),
          }),
        };
      }
    }
    setLocalErrors(mergedErr);
    if (Object.keys(mergedErr).length > 0) return mergedErr;
    const asyncErr = await runAsyncFormValidations(formManager, vals, itemsService, listTitle, itemId, submitKind);
    if (Object.keys(asyncErr).length > 0) {
      setLocalErrors(asyncErr);
      return asyncErr;
    }
    if (submitKind === 'submit' && linkedConfigsSorted.length > 0) {
      const rowErr: Record<string, Record<string, string>[]> = {};
      let anyMsg = false;
      for (let ci = 0; ci < linkedConfigsSorted.length; ci++) {
        const cfg = linkedConfigsSorted[ci];
        const meta = linkedMetaById[cfg.id];
        if (!meta?.length) continue;
        const shell = linkedChildFormAsManagerConfig(cfg);
        const rows = linkedRowsById[cfg.id] ?? [];
        const minR = cfg.minRows ?? 0;
        if (rows.length < minR) {
          anyMsg = true;
          if (!rowErr[cfg.id]) rowErr[cfg.id] = [];
          rowErr[cfg.id][0] = { _block: `Mínimo ${minR} linha(s).` };
        }
        const maxR = cfg.maxRows;
        if (maxR !== undefined && rows.length > maxR) {
          anyMsg = true;
          if (!rowErr[cfg.id]) rowErr[cfg.id] = [];
          rowErr[cfg.id][0] = { ...rowErr[cfg.id][0], _block: `Máximo ${maxR} linha(s).` };
        }
        for (let ri = 0; ri < rows.length; ri++) {
          const row = rows[ri];
          const ctxL: IFormRuleRuntimeContext = {
            formMode,
            values: row.values,
            submitKind,
            userGroupTitles,
            currentUserId,
            authorId,
            dynamicContext,
          };
          const attEmpty: IFormValidationAttachmentContext = {
            attachmentCount: 0,
            pendingFiles: [],
          };
          const syncL = collectFormValidationErrors(shell, cfg.fields, ctxL, attEmpty, undefined);
          const asyncL = await runAsyncFormValidations(
            shell,
            row.values,
            itemsService,
            cfg.listTitle.trim(),
            row.sharePointId,
            submitKind
          );
          const mergedL = { ...syncL, ...asyncL };
          if (Object.keys(mergedL).length > 0) {
            anyMsg = true;
            if (!rowErr[cfg.id]) rowErr[cfg.id] = [];
            while (rowErr[cfg.id].length <= ri) rowErr[cfg.id].push({});
            rowErr[cfg.id][ri] = { ...rowErr[cfg.id][ri], ...mergedL };
          }
        }
      }
      setLinkedRowErrorsById(rowErr);
      if (anyMsg) {
        const flat: Record<string, string> = { _linked: 'Corrija as listas vinculadas.' };
        setLocalErrors(flat);
        return flat;
      }
    }
    setLinkedRowErrorsById({});
    setLocalErrors({});
    return undefined;
  };

  const handleSave = async (
    submitKind: TFormSubmitKind,
    opts?: {
      valuesOverride?: Record<string, unknown>;
      buttonOverlayOverride?: IFormButtonFieldOverlay;
      submitLoadingFromButton?: IFormCustomButtonConfig;
    }
  ): Promise<boolean> => {
    const vals = opts?.valuesOverride ?? values;
    const ov = opts?.buttonOverlayOverride ?? buttonOverlay;
    setFormError(undefined);
    const validationErrors = await validate(submitKind, { values: vals, buttonOverlay: ov });
    if (validationErrors) {
      setFormError(formatValidationSummaryForForm(validationErrors, fieldLabelByName));
      return false;
    }
    setSubmitUi(resolveSubmitLoadingKind(formManager, opts?.submitLoadingFromButton));
    try {
      const payload = formValuesToSharePointPayload(fieldMetadata, vals, names, {
        nullWhenEmptyFieldNames: ocultosNullFieldNames,
      });
      const savedId = await onSubmit(
        payload,
        submitKind,
        flatPendingFiles,
        multiFolderAttachmentMode ? pendingFilesByFolder : undefined
      );
      if (submitKind === 'submit' && linkedConfigsSorted.length > 0) {
        const parentId = savedId ?? itemId;
        if (parentId !== undefined && parentId !== null && typeof parentId === 'number' && isFinite(parentId)) {
          try {
            await performLinkedSync(parentId);
          } catch (le) {
            setFormError(
              `Registo principal gravado, mas as listas vinculadas falharam: ${
                le instanceof Error ? le.message : String(le)
              }`
            );
            return false;
          }
        }
      }
      return true;
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
      return false;
    } finally {
      setSubmitUi(null);
    }
  };

  const stepsAll = formManager.steps?.length ? formManager.steps : null;
  const visibleStepsForUi = useMemo(() => {
    if (!stepsAll) return null;
    const nonSpecial = stepsAll.filter(
      (s) => s.id !== FORM_OCULTOS_STEP_ID && s.id !== FORM_FIXOS_STEP_ID
    );
    const forMode = nonSpecial.filter((s) => stepVisibleInFormMode(s, formMode));
    return forMode.length > 0 ? forMode : nonSpecial;
  }, [stepsAll, formMode]);

  const fixosStepConfig = stepsAll?.find((s) => s.id === FORM_FIXOS_STEP_ID);
  const fixosChromeActive =
    fixosStepConfig === undefined || stepVisibleInFormMode(fixosStepConfig, formMode);
  const [stepIndex, setStepIndex] = useState(0);
  const [historyBtn, setHistoryBtn] = useState<IFormCustomButtonConfig | null>(null);
  useEffect(() => {
    if (!visibleStepsForUi?.length) return;
    setStepIndex((i) => Math.min(i, visibleStepsForUi.length - 1));
  }, [visibleStepsForUi]);

  const linkedConfigsForCurrentMainStep = useMemo(() => {
    if (!linkedConfigsSorted.length) return [];
    const vis = visibleStepsForUi;
    if (!vis?.length) return linkedConfigsSorted;
    const curId = vis[stepIndex]?.id;
    if (!curId) return linkedConfigsSorted;
    const visibleIds = new Set(vis.map((s) => s.id));
    const defaultId = vis[0].id;
    return linkedConfigsSorted.filter((c) => {
      const raw = (c.mainFormStepId ?? '').trim();
      const resolved = raw && visibleIds.has(raw) ? raw : defaultId;
      return resolved === curId;
    });
  }, [linkedConfigsSorted, visibleStepsForUi, stepIndex]);

  const runCustomButton = async (btn: IFormCustomButtonConfig): Promise<void> => {
    const op: TFormCustomButtonOperation = btn.operation ?? 'legacy';
    if (op === 'history') {
      if (formManager.historyEnabled !== true) {
        setFormError('Ative o histórico na aba Lista de logs do gestor de formulário.');
        return;
      }
      if (itemId === undefined || itemId === null || formMode === 'create') {
        setFormError('O histórico só está disponível quando o item já existe na lista.');
        return;
      }
      setFormError(undefined);
      setHistoryBtn(btn);
      return;
    }
    const actions = op === 'redirect' ? [] : btn.actions ?? [];
    const { mergedValues, mergedOverlay } = reduceCustomButtonActions(
      actions,
      values,
      dynamicContext,
      buttonOverlay,
      attachmentFolderUrl
    );
    if (op !== 'redirect') {
      flushSync(() => {
        setValues(mergedValues);
        setButtonOverlay(mergedOverlay);
      });
      const behaviorEarly = btn.behavior ?? 'actionsOnly';
      if (
        behaviorEarly === 'actionsOnly' &&
        mergedOverlay.show.size > 0 &&
        visibleStepsForUi &&
        visibleStepsForUi.length > 1
      ) {
        const shownNames = Array.from(mergedOverlay.show);
        for (let si = 0; si < shownNames.length; si++) {
          const name = shownNames[si];
          const stepId = mergedOverlay.showOnStepId?.[name];
          if (!stepId) continue;
          const idx = visibleStepsForUi.findIndex((s) => s.id === stepId);
          if (idx >= 0) {
            setStepIndex(idx);
            break;
          }
        }
      }
    }

    if (op === 'redirect') {
      const tpl = (btn.redirectUrlTemplate ?? '').trim();
      if (!tpl) {
        setFormError('Configure o URL de redirecionamento no gestor de formulário.');
        return;
      }
      const url = interpolateFormButtonRedirectUrl(tpl, mergedValues, { itemId, formMode });
      try {
        await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
          sourceListTitle: listTitle,
          sourceItemId: itemId,
          formMode,
        });
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
        return;
      }
      window.location.assign(url);
      return;
    }

    if (op === 'add') {
      setFormError(undefined);
      const validationErrors = await validate('submit', { values: mergedValues, buttonOverlay: mergedOverlay });
      if (validationErrors) {
        setFormError(formatValidationSummaryForForm(validationErrors, fieldLabelByName));
        return;
      }
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      try {
        const payload = formValuesToSharePointPayload(fieldMetadata, mergedValues, names, {
          nullWhenEmptyFieldNames: ocultosNullFieldNames,
        });
        const { id: newId, filesForAttachments } = await itemsService.addItem(
          listTitle,
          payload,
          multiFolderAttachmentMode ? flatPendingFiles : pendingFiles
        );
        await uploadListItemAttachments(
          listTitle,
          newId,
          multiFolderAttachmentMode ? [] : filesForAttachments,
          formManager,
          {
            ...mergedValues,
            Id: newId,
          },
          multiFolderAttachmentMode ? pendingFilesByFolder : undefined
        );
        try {
          await performLinkedSync(newId);
        } catch (le) {
          setFormError(
            `Item criado, mas as listas vinculadas falharam: ${le instanceof Error ? le.message : String(le)}`
          );
          setSubmitUi(null);
          return;
        }
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: newId,
            formMode,
          });
        } catch (le) {
          setFormError(
            `Item criado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
          );
        }
        onDismiss();
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      return;
    }

    if (op === 'update') {
      if (!itemId || formMode === 'create') {
        setFormError('Atualizar requer um item aberto (parâmetros Form / FormID na página).');
        return;
      }
      setFormError(undefined);
      const validationErrors = await validate('submit', { values: mergedValues, buttonOverlay: mergedOverlay });
      if (validationErrors) {
        setFormError(formatValidationSummaryForForm(validationErrors, fieldLabelByName));
        return;
      }
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      try {
        const payload = formValuesToSharePointPayload(fieldMetadata, mergedValues, names, {
          nullWhenEmptyFieldNames: ocultosNullFieldNames,
        });
        await itemsService.updateItem(listTitle, itemId, payload);
        await uploadListItemAttachments(
          listTitle,
          itemId,
          multiFolderAttachmentMode ? [] : pendingFiles,
          formManager,
          {
            ...mergedValues,
            Id: itemId,
          },
          multiFolderAttachmentMode ? pendingFilesByFolder : undefined
        );
        try {
          await performLinkedSync(itemId);
        } catch (le) {
          setFormError(
            `Gravado, mas as listas vinculadas falharam: ${le instanceof Error ? le.message : String(le)}`
          );
          setSubmitUi(null);
          return;
        }
        await onAfterItemUpdated?.();
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: itemId,
            formMode,
          });
        } catch (le) {
          setFormError(
            `Gravado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
          );
        }
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      return;
    }

    if (op === 'delete') {
      if (!itemId || formMode === 'create') {
        setFormError('Eliminar só está disponível ao editar ou ver um item existente.');
        return;
      }
      if (!window.confirm('Eliminar este item permanentemente?')) return;
      setFormError(undefined);
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      try {
        await itemsService.deleteItem(listTitle, itemId);
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: itemId,
            formMode,
          });
        } catch (le) {
          setFormError(
            `Eliminado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
          );
        }
        onDismiss();
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      return;
    }

    const behavior = btn.behavior ?? 'actionsOnly';
    if (behavior === 'close') {
      try {
        await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
          sourceListTitle: listTitle,
          sourceItemId: itemId,
          formMode,
        });
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
        return;
      }
      onDismiss();
      return;
    }
    if (behavior === 'draft') {
      const saved = await handleSave('draft', {
        valuesOverride: mergedValues,
        buttonOverlayOverride: mergedOverlay,
        submitLoadingFromButton: btn,
      });
      if (saved) {
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: itemId,
            formMode,
          });
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
        }
      }
    } else if (behavior === 'submit') {
      const saved = await handleSave('submit', {
        valuesOverride: mergedValues,
        buttonOverlayOverride: mergedOverlay,
        submitLoadingFromButton: btn,
      });
      if (saved) {
        try {
          await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
            sourceListTitle: listTitle,
            sourceItemId: itemId,
            formMode,
          });
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
        }
      }
    } else if (behavior === 'actionsOnly') {
      try {
        await appendFormActionLogEntry(itemsService, formManager.actionLog, btn, {
          sourceListTitle: listTitle,
          sourceItemId: itemId,
          formMode,
        });
      } catch (e) {
        setFormError(e instanceof Error ? e.message : String(e));
      }
    }
  };

  const currentStepFieldSet = useMemo(() => {
    if (!visibleStepsForUi?.length) return null;
    const s = visibleStepsForUi[stepIndex];
    if (!s) return null;
    return new Set(s.fieldNames);
  }, [visibleStepsForUi, stepIndex]);

  const foldersForCurrentAttachmentStep = useMemo(() => {
    if (!multiFolderAttachmentMode) return [];
    const tree = formManager.attachmentLibrary?.folderTree;
    if (!tree?.length) return [];
    const curId = visibleStepsForUi?.[stepIndex]?.id;
    if (!curId) return [];
    const ctx = runtimeCtx();
    return flattenFolderTreeNodes(tree).filter(
      (n) => n.showUploaderInStepIds?.includes(curId) && isAttachmentFolderUploaderVisible(n, ctx)
    );
  }, [
    multiFolderAttachmentMode,
    formManager.attachmentLibrary?.folderTree,
    visibleStepsForUi,
    stepIndex,
    runtimeCtx,
    formMode,
    values,
    userGroupTitles,
    currentUserId,
    authorId,
    dynamicContext,
  ]);

  const libraryRowsForFolderNode = useCallback(
    (nodeId: string): IServerAttachmentRow[] => {
      if (!itemId) return [];
      const tree = formManager.attachmentLibrary?.folderTree;
      if (!tree?.length) return [];
      const out: IServerAttachmentRow[] = [];
      for (let i = 0; i < serverAttachments.length; i++) {
        const row = serverAttachments[i];
        const fr = row.fileRef;
        if (
          typeof fr === 'string' &&
          fr.trim() &&
          libraryFileRowBelongsToFolderNode(fr.trim(), nodeId, tree, itemId, values)
        ) {
          out.push(row);
        }
      }
      return out;
    },
    [itemId, formManager.attachmentLibrary?.folderTree, serverAttachments, values]
  );

  const tryGoToStep = useCallback(
    (targetIndex: number) => {
      if (!visibleStepsForUi?.length) return;
      const max = visibleStepsForUi.length - 1;
      const t = Math.max(0, Math.min(max, targetIndex));
      if (t === stepIndex) return;
      if (formMode === 'view') {
        setStepIndex(t);
        return;
      }
      const sn = formManager.stepNavigation;
      const requireBlock =
        sn?.requireFilledRequiredToAdvance === true ||
        (sn as { requireFilledRequiredToAdvance?: unknown } | undefined)?.requireFilledRequiredToAdvance ===
          'true';
      const fullVal = sn?.fullValidationOnAdvance === true;
      const allowBackFree = sn?.allowBackWithoutValidation !== false;
      const attCtx: IFormValidationAttachmentContext = {
        attachmentCount,
        pendingFiles: flatPendingFiles.map((f) => ({
          size: f.size,
          type: f.type || 'application/octet-stream',
          name: f.name,
        })),
      };
      const ctx: IFormRuleRuntimeContext = {
        formMode,
        values,
        submitKind: 'submit',
        userGroupTitles,
        currentUserId,
        authorId,
        dynamicContext,
        attachmentFolderUrl,
      };
      const overlay = { show: buttonOverlay.show, hide: buttonOverlay.hide };
      const syncErrorsForStep = (stepFieldList: Set<string>, stepIdx: number): Record<string, string> | null => {
        const derivedStep = buildFormDerivedState(formManager, fieldConfigs, ctx, overlay);
        const fv = (n: string): boolean => derivedStep.fieldVisible[n] !== false;
        const sync = collectFormValidationErrors(formManager, fieldConfigs, ctx, attCtx, overlay);
        let rel = filterValidationErrorsToStepFields(sync, stepFieldList);
        if (!fullVal) rel = pickRequiredStyleStepErrors(rel);
        const listReq = listRequiredEmptyErrorsInStep(stepFieldList, values, metaByName, fv);
        let merged: Record<string, string> = { ...rel, ...listReq };
        if (multiFolderAttachmentMode && isFormAttachmentLibraryRuntime(formManager)) {
          const tree = formManager.attachmentLibrary?.folderTree;
          if (tree?.length) {
            const folderAll = collectFolderAttachmentLimitErrors(tree, {
              pendingByFolder: pendingFilesByFolder,
              libraryCountByNodeId: (nodeId) => {
                if (!itemId) return 0;
                let c = 0;
                for (let i = 0; i < serverAttachments.length; i++) {
                  const row = serverAttachments[i];
                  const fr = row.fileRef;
                  if (typeof fr !== 'string' || !fr.trim()) continue;
                  if (libraryFileRowBelongsToFolderNode(fr.trim(), nodeId, tree, itemId, values)) c++;
                }
                return c;
              },
              isFolderUploaderVisible: (n) =>
                isAttachmentFolderUploaderVisible(n, {
                  formMode,
                  values,
                  submitKind: 'submit',
                  userGroupTitles,
                  currentUserId,
                  authorId,
                  dynamicContext,
                  attachmentFolderUrl,
                }),
            });
            const sid = visibleStepsForUi[stepIdx]?.id;
            merged = { ...merged, ...filterFolderLimitErrorsToStep(folderAll, tree, sid) };
          }
        }
        if (Object.keys(merged).length === 0) return null;
        return merged;
      };
      if (!requireBlock) {
        setFormError(undefined);
        setStepIndex(t);
        return;
      }
      if (t < stepIndex) {
        if (allowBackFree) {
          setFormError(undefined);
          setStepIndex(t);
          return;
        }
        const cur = visibleStepsForUi[stepIndex];
        const curSet = new Set(cur?.fieldNames ?? []);
        const bad = syncErrorsForStep(curSet, stepIndex);
        if (bad) {
          setLocalErrors(bad);
          setFormError(formatValidationSummaryForForm(bad, fieldLabelByName));
          return;
        }
        setFormError(undefined);
        setStepIndex(t);
        return;
      }
      for (let k = stepIndex; k < t; k++) {
        const st = visibleStepsForUi[k];
        const fieldSet = new Set(st?.fieldNames ?? []);
        const bad = syncErrorsForStep(fieldSet, k);
        if (bad) {
          setStepIndex(k);
          setLocalErrors(bad);
          setFormError(formatValidationSummaryForForm(bad, fieldLabelByName));
          return;
        }
      }
      setFormError(undefined);
      setStepIndex(t);
    },
    [
      visibleStepsForUi,
      stepIndex,
      formMode,
      formManager,
      fieldConfigs,
      fieldLabelByName,
      values,
      userGroupTitles,
      currentUserId,
      authorId,
      dynamicContext,
      attachmentCount,
      flatPendingFiles,
      buttonOverlay.show,
      buttonOverlay.hide,
      buttonOverlay.showOnStepId,
      metaByName,
      multiFolderAttachmentMode,
      pendingFilesByFolder,
      serverAttachments,
      itemId,
      attachmentFolderUrl,
      formManager.attachmentLibrary?.folderTree,
    ]
  );

  const [modalOpen, setModalOpen] = useState(false);
  const modalGroupIds = useMemo(() => {
    const seen: Record<string, boolean> = {};
    const ids: string[] = [];
    for (let i = 0; i < fieldConfigs.length; i++) {
      const gid = fieldConfigs[i].modalGroupId;
      if (gid && !seen[gid]) {
        seen[gid] = true;
        ids.push(gid);
      }
    }
    return ids;
  }, [fieldConfigs]);

  const renderServerAttachmentList = (rows: IServerAttachmentRow[]): React.ReactNode => {
    if (rows.length === 0) return null;
    const showIcon =
      attachmentPreviewKind === 'iconAndName' ||
      attachmentPreviewKind === 'thumbnailAndName' ||
      attachmentPreviewKind === 'thumbnailLarge';
    const iconPx = attachmentPreviewKind === 'thumbnailLarge' ? 48 : 20;
    const thumbBox =
      attachmentPreviewKind === 'thumbnailAndName' || attachmentPreviewKind === 'thumbnailLarge';
    const boxPx = attachmentPreviewKind === 'thumbnailLarge' ? 56 : 40;
    return (
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
            {a.fileUrl ? (
              <Link href={a.fileUrl} target="_blank" rel="noopener noreferrer">
                {a.fileName}
              </Link>
            ) : (
              <Text variant="small">{a.fileName}</Text>
            )}
          </Stack>
        ))}
      </Stack>
    );
  };

  const renderBannerVisual = (fc: IFormFieldConfig): React.ReactNode => {
    const name = fc.internalName;
    if (derived.fieldVisible[name] === false) return null;
    const url = (fc.bannerImageUrl ?? '').trim();
    const bannerLabel = (fc.label ?? 'Banner').trim() || 'Banner';
    const wPct = resolveBannerWidthPercent(fc);
    const hPct = resolveBannerHeightPercent(fc);
    const imgStyle: React.CSSProperties = {
      width: `${wPct}%`,
      maxWidth: '100%',
      height: 'auto',
      display: 'block',
      margin: '0 auto',
      borderRadius: 2,
      objectFit: 'contain',
      ...(hPct !== undefined ? { maxHeight: `${hPct}vh` } : {}),
    };
    return (
      <>
        {url ? <img src={url} alt={bannerLabel} style={imgStyle} /> : null}
        {fc.helpText && (
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            {fc.helpText}
          </Text>
        )}
      </>
    );
  };

  const renderFieldControl = (fc: IFormFieldConfig): React.ReactNode => {
    const name = fc.internalName;
    if (derived.fieldVisible[name] === false) return null;
    if (name === FORM_ATTACHMENTS_FIELD_INTERNAL) {
      const readOnly = formMode === 'view' || derived.fieldDisabled[name] === true;
      const attErr = localErrors._attachments;
      const attReq = derived.fieldRequired[name] === true;
      const attachmentSatisfied =
        flatPendingFiles.length > 0 || (formMode !== 'create' && attachmentCount > 0);
      const attReqEmpty = attReq && !readOnly && !attachmentSatisfied;
      if (
        multiFolderAttachmentMode &&
        foldersForCurrentAttachmentStep.length === 0 &&
        formMode !== 'view'
      ) {
        return null;
      }
      if (formMode === 'view') {
        if (multiFolderAttachmentMode && foldersForCurrentAttachmentStep.length === 0) {
          return null;
        }
        if (multiFolderAttachmentMode && foldersForCurrentAttachmentStep.length > 0) {
          let stepCount = 0;
          for (let si = 0; si < foldersForCurrentAttachmentStep.length; si++) {
            stepCount += libraryRowsForFolderNode(foldersForCurrentAttachmentStep[si].id).length;
          }
          return (
            <Stack key={name} tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 12 } }}>
              <Label required={attReq}>{fc.label ?? 'Anexos ao item'}</Label>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {stepCount} anexo(s) nesta etapa. Não é possível adicionar novos em modo ver.
              </Text>
              {foldersForCurrentAttachmentStep.map((node) => {
                const folderRows = libraryRowsForFolderNode(node.id);
                return (
                  <Stack key={node.id} tokens={{ childrenGap: 6 }}>
                    <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                      {node.nameTemplate?.trim() || 'Pasta'}
                    </Text>
                    {folderRows.length > 0 ? (
                      renderServerAttachmentList(folderRows)
                    ) : (
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        Nenhum ficheiro nesta pasta.
                      </Text>
                    )}
                  </Stack>
                );
              })}
              {fc.helpText && (
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  {fc.helpText}
                </Text>
              )}
            </Stack>
          );
        }
        return (
          <Stack key={name} tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 12 } }}>
            <Label required={attReq}>{fc.label ?? 'Anexos ao item'}</Label>
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              {attachmentCount} anexo(s) no item. Não é possível adicionar novos em modo ver.
            </Text>
            {serverAttachments.length > 0 ? (
              renderServerAttachmentList(serverAttachments)
            ) : (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                Nenhum ficheiro na pasta de anexos do item.
              </Text>
            )}
            {fc.helpText && (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {fc.helpText}
              </Text>
            )}
          </Stack>
        );
      }
      if (multiFolderAttachmentMode && foldersForCurrentAttachmentStep.length > 0) {
        return (
          <Stack key={name} tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 12 } }}>
            {foldersForCurrentAttachmentStep.map((node) => {
              const folderRows = libraryRowsForFolderNode(node.id);
              const libCount = folderRows.length;
              const minA = node.minAttachmentCount;
              const maxA = node.maxAttachmentCount;
              const limHint =
                minA !== undefined && minA > 0 && maxA !== undefined
                  ? `Entre ${minA} e ${maxA} ficheiro(s) nesta pasta.`
                  : minA !== undefined && minA > 0
                    ? `Mínimo ${minA} ficheiro(s) nesta pasta.`
                    : maxA !== undefined
                      ? `Máximo ${maxA} ficheiro(s) nesta pasta.`
                      : undefined;
              const desc = [limHint, fc.helpText].filter(Boolean).join(' · ') || undefined;
              const folderLimKey = `${FOLDER_ATTACHMENT_LIMIT_ERROR_PREFIX}${node.id}`;
              const folderLimErr = localErrors[folderLimKey];
              return (
                <Stack key={node.id} tokens={{ childrenGap: 6 }}>
                  {folderRows.length > 0 && (
                    <Stack tokens={{ childrenGap: 6 }}>
                      <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                        {isFormAttachmentLibraryRuntime(formManager)
                          ? 'Ficheiros já na biblioteca'
                          : 'Anexos já no item'}
                      </Text>
                      {renderServerAttachmentList(folderRows)}
                    </Stack>
                  )}
                  <FormAttachmentUploader
                    files={pendingFilesByFolder[node.id] ?? []}
                    onFilesChange={(files) => {
                      setPendingFilesByFolder((prev) => ({ ...prev, [node.id]: files }));
                    }}
                    disabled={readOnly}
                    label={node.nameTemplate?.trim() || 'Pasta'}
                    description={desc}
                    errorMessage={folderLimErr || attErr}
                    required={attReq}
                    requiredEmptyHighlight={attReqEmpty}
                    layout={formManager.attachmentUploadLayout ?? 'default'}
                    filePreview={attachmentPreviewKind}
                    allowedFileExtensions={
                      attachmentAllowedExtensions.length > 0 ? attachmentAllowedExtensions : undefined
                    }
                    priorFileCount={libCount}
                    maxTotalAttachmentCount={maxA}
                  />
                </Stack>
              );
            })}
          </Stack>
        );
      }
      return (
        <Stack key={name} tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
          {serverAttachments.length > 0 && (
            <Stack tokens={{ childrenGap: 6 }}>
              <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                {isFormAttachmentLibraryRuntime(formManager)
                  ? 'Ficheiros já na biblioteca (ligados a este item)'
                  : 'Anexos já no item'}
              </Text>
              {renderServerAttachmentList(serverAttachments)}
            </Stack>
          )}
          <FormAttachmentUploader
            files={pendingFiles}
            onFilesChange={setPendingFiles}
            disabled={readOnly}
            label={fc.label ?? 'Anexos ao item'}
            description={fc.helpText}
            errorMessage={attErr}
            required={attReq}
            requiredEmptyHighlight={attReqEmpty}
            layout={formManager.attachmentUploadLayout ?? 'default'}
            filePreview={attachmentPreviewKind}
            allowedFileExtensions={
              attachmentAllowedExtensions.length > 0 ? attachmentAllowedExtensions : undefined
            }
          />
        </Stack>
      );
    }
    if (isFormBannerFieldConfig(fc)) {
      if (fc.sectionId === FORM_FIXOS_STEP_ID) {
        return (
          <Stack key={name} tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 12 } }}>
            {renderBannerVisual(fc)}
          </Stack>
        );
      }
      if (resolveBannerPlacement(fc) !== 'inStep') return null;
      return (
        <Stack key={name} tokens={{ childrenGap: 6 }} styles={{ root: { marginBottom: 12 } }}>
          {renderBannerVisual(fc)}
        </Stack>
      );
    }
    const m = metaByName.get(name);
    if (!m) return null;
    const disabled = formMode === 'view' || derived.fieldDisabled[name] === true;
    const readOnly = derived.fieldReadOnly[name] === true || disabled;
    const err = localErrors[name];
    const label = fc.label ?? m.Title;
    const help = derived.dynamicHelpByField[name] ?? fc.helpText;
    const isRequired = derived.fieldRequired[name] === true || m.Required === true;
    const canFill = formMode !== 'view' && !readOnly;
    const showReqEmpty = isRequired && canFill && isValueEmptyForRequired(values[name], m.MappedType);
    const comp = derived.computedDisplay[name];
    if (comp !== undefined) {
      return (
        <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 12 } }}>
          <Label required={isRequired}>{label}</Label>
          <Text styles={{ root: { color: '#323130' } }}>{String(comp)}</Text>
          {help && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
        </Stack>
      );
    }

    const common = { disabled: readOnly, errorMessage: err };

    switch (m.MappedType) {
      case 'boolean':
        return (
          <Stack
            key={name}
            tokens={{ childrenGap: 6 }}
            styles={{
              root: {
                marginBottom: 12,
                ...(showReqEmpty
                  ? { borderLeft: `3px solid ${REQ_EMPTY_BORDER}`, paddingLeft: 8, paddingTop: 2, paddingBottom: 2 }
                  : {}),
              },
            }}
          >
            <Label required={isRequired}>{label}</Label>
            <Toggle
              ariaLabel={label}
              onText="Sim"
              offText="Não"
              checked={values[name] === true || values[name] === 1}
              onChange={(_, c) => updateField(name, !!c)}
              disabled={readOnly}
            />
          </Stack>
        );
      case 'number':
      case 'currency':
        return (
          <TextField
            key={name}
            label={label}
            type="number"
            placeholder={fc.placeholder}
            value={values[name] !== null && values[name] !== undefined ? String(values[name]) : ''}
            onChange={(_, v) => updateField(name, v === '' ? null : Number(v))}
            required={isRequired}
            {...common}
            description={help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
          />
        );
      case 'datetime':
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 8 } }}>
            <Label required={isRequired}>{label}</Label>
            <DatePicker
              value={values[name] ? new Date(String(values[name])) : undefined}
              onSelectDate={(d) => updateField(name, d ? d.toISOString() : null)}
              disabled={readOnly}
              textField={{
                errorMessage: err,
                styles: stylesTextFieldRequiredEmpty(showReqEmpty),
              }}
            />
            {help && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
          </Stack>
        );
      case 'choice': {
        const raw = (m.Choices ?? []).map((c) => ({ key: c, text: c }));
        const opts: IDropdownOption[] = !isRequired ? [{ key: '', text: '—' }, ...raw] : raw;
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            options={opts}
            selectedKey={
              values[name] !== undefined && values[name] !== null && String(values[name]) !== ''
                ? String(values[name])
                : ''
            }
            onChange={(_, o) => o && updateField(name, o.key === '' ? null : o.key)}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'multichoice': {
        const selected: string[] = Array.isArray(values[name])
          ? (values[name] as string[])
          : typeof values[name] === 'string'
            ? String(values[name]).split(';').map((s) => s.trim()).filter(Boolean)
            : [];
        const opts: IDropdownOption[] = (m.Choices ?? []).map((c) => ({
          key: c,
          text: c,
          selected: selected.indexOf(c) !== -1,
        }));
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            multiSelect
            options={opts}
            selectedKeys={selected}
            onChange={(_, o) => {
              if (!o) return;
              const k = String(o.key);
              const next = selected.indexOf(k) !== -1 ? selected.filter((x) => x !== k) : [...selected, k];
              updateField(name, next);
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'lookup': {
        const id = lookupIdFromValue(values[name]);
        const baseOpts = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const opts =
          id !== undefined && id > 0
            ? mergeOptionsForIds(baseOpts, [{ id, label: userTitleFromValue(values[name]) }])
            : baseOpts;
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            options={opts}
            selectedKey={id !== undefined ? String(id) : ''}
            onChange={(_, o) => {
              if (!o || o.key === '') updateField(name, null);
              else updateField(name, { Id: Number(o.key), Title: String(o.text ?? '') });
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'lookupmulti': {
        const selected = normalizeIdTitleArray(values[name]);
        const baseRaw = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const baseOpts = baseRaw.filter((o) => o.key !== '');
        const extra = selected.map((x) => ({ id: x.Id, label: x.Title }));
        const opts = mergeOptionsForIds(baseOpts, extra);
        const keys = selected.map((x) => String(x.Id));
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            multiSelect
            options={opts}
            selectedKeys={keys}
            onChange={(_, o) => {
              if (!o || o.key === '') return;
              const k = String(o.key);
              const hit = selected.findIndex((x) => String(x.Id) === k);
              const next =
                hit === -1
                  ? [...selected, { Id: Number(o.key), Title: String(o.text ?? '') }]
                  : selected.filter((_, i) => i !== hit);
              updateField(name, next);
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'user': {
        const id = lookupIdFromValue(values[name]);
        const baseOpts = !isRequired ? siteUserOptions : siteUserOptions.filter((o) => o.key !== '');
        const opts =
          id !== undefined && id > 0
            ? mergeOptionsForIds(baseOpts, [{ id, label: userTitleFromValue(values[name]) }])
            : baseOpts;
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            options={opts}
            selectedKey={id !== undefined ? String(id) : ''}
            onChange={(_, o) => {
              if (!o || o.key === '') updateField(name, null);
              else updateField(name, { Id: Number(o.key), Title: String(o.text ?? '') });
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'usermulti': {
        const selected = normalizeIdTitleArray(values[name]);
        const baseOpts = siteUserOptions.filter((o) => o.key !== '');
        const extra = selected.map((x) => ({ id: x.Id, label: x.Title }));
        const opts = mergeOptionsForIds(baseOpts, extra);
        const keys = selected.map((x) => String(x.Id));
        return (
          <Dropdown
            key={name}
            label={label}
            placeholder={fc.placeholder}
            multiSelect
            options={opts}
            selectedKeys={keys}
            onChange={(_, o) => {
              if (!o || o.key === '') return;
              const k = String(o.key);
              const hit = selected.findIndex((x) => String(x.Id) === k);
              const next =
                hit === -1
                  ? [...selected, { Id: Number(o.key), Title: String(o.text ?? '') }]
                  : selected.filter((_, i) => i !== hit);
              updateField(name, next);
            }}
            required={isRequired}
            errorMessage={err}
            disabled={readOnly}
            styles={dropdownReqStyles(showReqEmpty)}
          />
        );
      }
      case 'url': {
        const uv = parseUrlFieldValue(values[name]);
        return (
          <Stack key={name} tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
            <Label required={isRequired}>{label}</Label>
            <TextField
              label="Endereço web"
              placeholder="https://"
              value={uv.Url}
              onChange={(_, v) =>
                updateField(name, { Url: v ?? '', Description: uv.Description })
              }
              disabled={readOnly}
              errorMessage={err}
              styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
            />
            <TextField
              label="Descrição a apresentar"
              value={uv.Description}
              onChange={(_, v) =>
                updateField(name, { Url: uv.Url, Description: v ?? '' })
              }
              disabled={readOnly}
            />
            {help && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text>}
          </Stack>
        );
      }
      case 'multiline':
        return (
          <TextField
            key={name}
            label={label}
            multiline
            rows={4}
            placeholder={fc.placeholder}
            value={values[name] !== null && values[name] !== undefined ? String(values[name]) : ''}
            onChange={(_, v) => updateField(name, v ?? '')}
            required={isRequired}
            {...common}
            description={help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
          />
        );
      default:
        return (
          <TextField
            key={name}
            label={label}
            placeholder={fc.placeholder}
            value={values[name] !== null && values[name] !== undefined ? String(values[name]) : ''}
            onChange={(_, v) => updateField(name, v ?? '')}
            required={isRequired}
            {...common}
            description={help}
            styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
          />
        );
    }
  };

  const renderFields = (scope: 'main' | 'modal'): React.ReactNode => {
    const bySection = new Map<string, IFormFieldConfig[]>();
    for (let i = 0; i < fieldConfigs.length; i++) {
      const fc = fieldConfigs[i];
      const inModal = !!fc.modalGroupId;
      if (scope === 'modal' && !inModal) continue;
      if (scope === 'main' && inModal) continue;
      if (fc.sectionId === FORM_FIXOS_STEP_ID) continue;
      if (
        isFormBannerFieldConfig(fc) &&
        (resolveBannerPlacement(fc) === 'topFixed' || resolveBannerPlacement(fc) === 'bottomFixed')
      ) {
        continue;
      }
      if (scope === 'main' && currentStepFieldSet) {
        const name = fc.internalName;
        if (!currentStepFieldSet.has(name)) {
          const allowMultiAttach =
            name === FORM_ATTACHMENTS_FIELD_INTERNAL &&
            multiFolderAttachmentMode &&
            foldersForCurrentAttachmentStep.length > 0;
          if (!allowMultiAttach) {
            if (!buttonOverlay.show.has(name)) continue;
            const vis = visibleStepsForUi;
            if (!vis?.length) continue;
            const sid = buttonOverlay.showOnStepId?.[name];
            const fallbackSingle = vis.length === 1 ? vis[0].id : undefined;
            const target = sid || fallbackSingle;
            const curId = vis[stepIndex]?.id;
            if (!target || !curId || target !== curId) continue;
          }
        }
      }
      let sid =
        derived.effectiveSectionByField[fc.internalName] ?? fc.sectionId ?? formManager.sections[0]?.id ?? 'main';
      if (
        scope === 'main' &&
        fc.sectionId === FORM_OCULTOS_STEP_ID &&
        buttonOverlay.show.has(fc.internalName) &&
        visibleStepsForUi?.length
      ) {
        const mapSid = buttonOverlay.showOnStepId?.[fc.internalName];
        const fallbackSingle = visibleStepsForUi.length === 1 ? visibleStepsForUi[0].id : undefined;
        const stepTarget = mapSid || fallbackSingle;
        if (stepTarget && formManager.sections.some((x) => x.id === stepTarget)) {
          sid = stepTarget;
        }
      }
      const arr = bySection.get(sid) ?? [];
      arr.push(fc);
      bySection.set(sid, arr);
    }
    const out: React.ReactNode[] = [];
    for (let s = 0; s < formManager.sections.length; s++) {
      const sec = formManager.sections[s];
      if (sec.id === FORM_OCULTOS_STEP_ID || sec.id === FORM_FIXOS_STEP_ID) continue;
      if (derived.sectionVisible[sec.id] === false) continue;
      const fields = bySection.get(sec.id);
      if (!fields?.length) continue;
      out.push(
        <Stack key={sec.id} tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 16 } }}>
          <Text variant="large" styles={{ root: { fontWeight: 600 } }}>{sec.title}</Text>
          {fields.map((fc) => renderFieldControl(fc))}
        </Stack>
      );
    }
    return <>{out}</>;
  };

  const topChromeFields = fieldConfigs.filter((fc) => {
    if (derived.fieldVisible[fc.internalName] === false) return false;
    if (fc.sectionId === FORM_FIXOS_STEP_ID) {
      if (!fixosChromeActive) return false;
      return resolveFixedPlacement(fc) === 'top';
    }
    return isFormBannerFieldConfig(fc) && resolveBannerPlacement(fc) === 'topFixed';
  });
  const bottomChromeFields = fieldConfigs.filter((fc) => {
    if (derived.fieldVisible[fc.internalName] === false) return false;
    if (fc.sectionId === FORM_FIXOS_STEP_ID) {
      if (!fixosChromeActive) return false;
      return resolveFixedPlacement(fc) === 'bottom';
    }
    return isFormBannerFieldConfig(fc) && resolveBannerPlacement(fc) === 'bottomFixed';
  });

  const submitMsg = 'A gravar…';

  return (
    <Stack tokens={{ childrenGap: 16 }} styles={{ root: { paddingTop: 8 } }}>
      <FormSubmitLoadingChrome kind="infoBar" active={submitUi === 'infoBar'} message={submitMsg} />
      <FormSubmitLoadingChrome kind="topProgress" active={submitUi === 'topProgress'} message={submitMsg} />
      <Stack
        styles={{
          root: {
            position: 'relative',
            minHeight: submitUi === 'overlay' || submitUi === 'formShimmer' ? 160 : undefined,
          },
        }}
        tokens={{ childrenGap: 16 }}
      >
        {formError && (
          <MessageBar messageBarType={MessageBarType.error} styles={{ root: { whiteSpace: 'pre-line' } }}>
            {formError}
          </MessageBar>
        )}
        {localErrors._attachments && (
          <MessageBar messageBarType={MessageBarType.error}>{localErrors._attachments}</MessageBar>
        )}
        {localErrors._async && <MessageBar messageBarType={MessageBarType.error}>{localErrors._async}</MessageBar>}
        {derived.messages.map((m) => (
          <MessageBar
            key={m.ruleId}
            messageBarType={
              m.variant === 'error'
                ? MessageBarType.error
                : m.variant === 'warning'
                  ? MessageBarType.warning
                  : MessageBarType.info
            }
          >
            {m.text}
          </MessageBar>
        ))}
        {topChromeFields.length > 0 && (
          <FormChromeZone
            zone="top"
            fields={topChromeFields}
            renderField={(fc) => renderFieldControl(fc)}
            layoutDeps={values}
          />
        )}
        {visibleStepsForUi && visibleStepsForUi.length > 1 && (
          <FormStepNavigation
            steps={visibleStepsForUi}
            activeIndex={stepIndex}
            onStepSelect={(i) => tryGoToStep(i)}
            layout={formManager.stepLayout ?? 'segmented'}
          />
        )}
        {modalGroupIds.length > 0 && formMode !== 'view' && (
          <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
            {modalGroupIds.map((gid: string) => (
              <DefaultButton key={gid} text={`Editar ${gid}`} onClick={() => setModalOpen(true)} />
            ))}
          </Stack>
        )}
        {renderFields('main')}
        {visibleStepsForUi && visibleStepsForUi.length > 1 && (
          <Stack
            horizontal
            verticalAlign="center"
            tokens={{ childrenGap: 8 }}
            styles={{ root: { width: '100%' } }}
            wrap
          >
            <FormStepPrevNextNav
              variant={formManager.stepNavButtons ?? 'fluent'}
              stepIndex={stepIndex}
              stepCount={visibleStepsForUi.length}
              onPrev={() => tryGoToStep(stepIndex - 1)}
              onNext={() => tryGoToStep(stepIndex + 1)}
              disabled={submitting}
            />
          </Stack>
        )}
        {bottomChromeFields.length > 0 && (
          <FormChromeZone
            zone="bottom"
            fields={bottomChromeFields}
            renderField={(fc) => renderFieldControl(fc)}
            layoutDeps={values}
          />
        )}
        {linkedConfigsForCurrentMainStep.length > 0 && (
          <LinkedChildFormsBlock
            configs={linkedConfigsForCurrentMainStep}
            parentItemId={itemId}
            formMode={formMode}
            rowsByConfigId={linkedRowsById}
            onRowsChange={(configId, rows) =>
              setLinkedRowsById((prev) => ({ ...prev, [configId]: rows }))
            }
            fieldMetaByConfigId={linkedMetaById}
            loadingByConfigId={linkedLoadingById}
            errorByConfigId={linkedLoadErrById}
            userGroupTitles={userGroupTitles}
            currentUserId={currentUserId}
            authorId={authorId}
            dynamicContext={dynamicContext}
            rowErrorsByConfigId={linkedRowErrorsById}
          />
        )}
        <Stack horizontal tokens={{ childrenGap: 8 }} wrap>
          {(formManager.customButtons ?? [])
            .filter((b) =>
              formManager.historyEnabled === true ? (b.operation ?? 'legacy') !== 'history' : true
            )
            .filter((b) =>
              shouldShowCustomButton(b, runtimeCtx(), {
                allRequiredFilled: allRequiredFilled,
                historyEnabledInConfig: formManager.historyEnabled === true,
                historyItemId: itemId,
              })
            )
            .map((b) =>
              b.appearance === 'primary' ? (
                <PrimaryButton
                  key={b.id}
                  text={b.label}
                  title={b.shortDescription || undefined}
                  onClick={() => void runCustomButton(b)}
                  disabled={submitting}
                />
              ) : (
                <DefaultButton
                  key={b.id}
                  text={b.label}
                  title={b.shortDescription || undefined}
                  onClick={() => void runCustomButton(b)}
                  disabled={submitting}
                />
              )
            )}
          {formManager.historyEnabled === true &&
            shouldShowBuiltinHistoryButton({
              historyEnabledInConfig: true,
              historyItemId: itemId,
              historyGroupTitles: formManager.historyGroupTitles,
              userGroupTitles,
            }) &&
            (() => {
              const hk = formManager.historyButtonKind ?? 'text';
              const hLabel = (formManager.historyButtonLabel ?? 'Histórico').trim() || 'Histórico';
              const hIcon = (formManager.historyButtonIcon ?? 'History').trim() || 'History';
              const hSub = formManager.historyPanelSubtitle?.trim();
              const onH = (): void => {
                void runCustomButton(builtinHistoryButtonConfig);
              };
              if (hk === 'icon') {
                return (
                  <IconButton
                    key="__builtin_history"
                    iconProps={{ iconName: hIcon }}
                    title={hSub || hLabel}
                    ariaLabel={hLabel}
                    onClick={onH}
                    disabled={submitting}
                  />
                );
              }
              if (hk === 'iconAndText') {
                return (
                  <DefaultButton
                    key="__builtin_history"
                    text={hLabel}
                    iconProps={{ iconName: hIcon }}
                    title={hSub}
                    onClick={onH}
                    disabled={submitting}
                  />
                );
              }
              return (
                <DefaultButton
                  key="__builtin_history"
                  text={hLabel}
                  title={hSub}
                  onClick={onH}
                  disabled={submitting}
                />
              );
            })()}
        </Stack>
        {historyBtn &&
          itemId !== undefined &&
          itemId !== null &&
          formManager.historyEnabled === true && (
            <FormItemHistoryUi
              actionLog={formManager.actionLog}
              sourceItemId={itemId}
              presentationKind={formManager.historyPresentationKind ?? 'panel'}
              layoutKind={formManager.historyLayoutKind ?? 'list'}
              isOpen={true}
              onDismiss={() => setHistoryBtn(null)}
              title={historyBtn.label}
              subtitle={historyBtn.shortDescription}
            />
          )}
        <FormSubmitLoadingChrome kind="belowButtons" active={submitUi === 'belowButtons'} message={submitMsg} />
        <FormSubmitLoadingChrome kind="overlay" active={submitUi === 'overlay'} message={submitMsg} />
        <FormSubmitLoadingChrome kind="formShimmer" active={submitUi === 'formShimmer'} message={submitMsg} />
        {modalOpen && (
          <Stack styles={{ root: { borderTop: '1px solid #edebe9', paddingTop: 16 } }} tokens={{ childrenGap: 12 }}>
            <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>Campos adicionais</Text>
            {renderFields('modal')}
            <DefaultButton text="Fechar modal" onClick={() => setModalOpen(false)} />
          </Stack>
        )}
      </Stack>
    </Stack>
  );
};
