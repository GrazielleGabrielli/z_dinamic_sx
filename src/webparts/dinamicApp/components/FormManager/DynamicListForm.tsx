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
  Icon,
  Dialog,
  DialogFooter,
  DialogType,
  Spinner,
  SpinnerSize,
  useTheme,
} from '@fluentui/react';
import type { IFieldMetadata } from '../../../../services';
import type {
  IFormManagerConfig,
  IFormFieldConfig,
  IFormCustomButtonConfig,
  IFormLinkedChildFormConfig,
  TFormCustomButtonConfirmKind,
  TFormButtonAction,
  TFormCustomButtonOperation,
  TFormManagerFormMode,
  TFormSubmitKind,
  TFormRule,
  TFormSubmitLoadingUiKind,
  TFormAttachmentFilePreviewKind,
  IFormManagerActionLogConfig,
} from '../../core/config/types/formManager';
import { IMaskInput } from 'react-imask';
import { resolveTextInputMaskOptions } from '../../core/formManager/formTextInputMasks';
import { FLUENT_DATE_PICKER_PT_BR } from '../../core/formManager/fluentDatePickerPtBr';
import { applyTextTransformsToRecordValues } from '../../core/formManager/formTextValueTransform';
import {
  buildLookupDropdownSelectRaw,
  buildLookupODataFilter,
  hasConfiguredLookupFilter,
  isParentValueReadyForLookupFilter,
  lookupRowToOptionText,
  resolveLookupFormLabelInternalName,
} from '../../core/formManager/lookupFormDropdownHelpers';
import {
  FORM_ATTACHMENTS_FIELD_INTERNAL,
  FORM_OCULTOS_STEP_ID,
  FORM_FIXOS_STEP_ID,
  FORM_BUILTIN_HISTORY_BUTTON_ID,
  isFormBannerFieldConfig,
  resolveBannerPlacement,
  resolveBannerWidthPercent,
  resolveBannerHeightPercent,
  resolveFixedPlacement,
  resolveChromePositionMode,
  resolveTextareaRows,
} from '../../core/config/types/formManager';
import type { IDynamicContext } from '../../core/dynamicTokens/types';
import { isDynamicToken } from '../../core/dynamicTokens';
import {
  buildFormDerivedState,
  collectFormValidationErrors,
  evaluateValidateDateRulesForField,
  filterValidationErrorsToStepFields,
  pickRequiredStyleStepErrors,
  evaluateCondition,
  evaluateFormValueExpression,
  getDefaultValuesFromRules,
  getMergedValidateValueLengthBounds,
  getMergedValidateValueNumberBounds,
  shouldShowCustomButton,
  shouldShowBuiltinHistoryButton,
  areAllRequiredFieldsFilled,
  isAttachmentFolderUploaderVisible,
  withRuleRuntimeDynamicContext,
  buildFormFieldLabelMap,
  buildValidationModalSections,
  formatValidationSummaryForForm,
  findEnabledSetComputedRule,
  resolveSetComputedDisplayValue,
  buildPostCreateItemIdComputedPatch,
  type IFormAttachmentFolderUrlContext,
  type IFormRuleRuntimeContext,
  type IFormValidationAttachmentContext,
} from '../../core/formManager/formRuleEngine';
import { collectFormManagerReferencedPayloadFieldNames } from '../../core/formManager/collectFormManagerReferencedPayloadFieldNames';
import { formValuesToSharePointPayload } from '../../core/formManager/formSharePointValues';
import { FormStepNavigation, FormStepPrevNextNav } from './FormStepLayoutUi';
import { FormAttachmentUploader } from './FormAttachmentUploader';
import { AttachmentFileDetailModal } from './AttachmentFileDetailModal';
import { runAsyncFormValidations } from '../../core/formManager/formAsyncValidation';
import { interpolateFormButtonRedirectUrl } from '../../core/formManager/formButtonRedirectUrl';
import { ensureAbsoluteSharePointUrl, parseUrlFieldValue } from '../../core/formManager/formUrlUtils';
import {
  appendFormActionLogEntry,
  type IFormActionLogRuntimeContext,
} from '../../core/formManager/formActionLog';
import { parseAttachmentUiRule } from '../../core/formManager/formManagerVisualModel';
import {
  initConfirmPromptEditor,
  confirmPromptEditorToValue,
  confirmPromptEditorIsFilled,
  isConfirmPromptEligibleField,
  type IConfirmPromptEditorState,
} from '../../core/formManager/confirmPromptFieldHelpers';
import { ItemsService, UsersService } from '../../../../services';
import { getSP, getSPForWeb } from '../../../../services/core/sp';
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
import { MultilineReadonlyHtml } from './MultilineReadonlyHtml';
import { multiSelectDropdownStyles, renderMultiSelectDropdownTitle } from './formMultiSelectDropdownUi';
import { attachmentFileKindIconName } from './attachmentFileKindIcon';
import { shouldRenderMultilineNoteAsHtml } from '../../core/formManager/sharePointNoteHtml';
import { stepVisibleInFormMode } from '../../core/formManager/stepFormMode';
import {
  linkedChildFormAsManagerConfig,
  loadLinkedChildRows,
  syncAllLinkedChildLists,
  type ILinkedChildRowState,
} from '../../core/formManager/formLinkedChildSync';
import { linkedChildAttPendingKey } from '../../core/formManager/linkedChildAttachmentPendingKeys';
import {
  buildMinimalFormManagerForLinkedLibraryUpload,
  resolveLinkedChildAttachmentRuntime,
} from '../../core/formManager/linkedChildAttachmentRuntime';
import { FieldsService } from '../../../../services';
import {
  getFilledPaletteButtonStyles,
  resolveActionLogPaletteAccentHex,
  resolveFormCustomButtonPaletteSlot,
  resolveStepUiAccentColor,
} from '../../core/formManager/formCustomButtonTheme';
import { applyFormManagerPermissionBreak } from '../../core/formManager/applyFormPermissionBreak';

function pickMainAuthorId(
  values: Record<string, unknown>,
  initialItem: Record<string, unknown> | null | undefined,
  currentUserId: number
): number | undefined {
  const raw = values.AuthorId ?? values.authorId;
  if (typeof raw === 'number' && raw > 0) return raw;
  const init = initialItem?.AuthorId ?? initialItem?.authorId;
  if (typeof init === 'number' && init > 0) return init;
  if (typeof currentUserId === 'number' && currentUserId > 0) return currentUserId;
  return undefined;
}

function validateValueCharLimitHint(
  len: number,
  minLength: number,
  maxLength: number
): { text: string; color: string } {
  if (len > maxLength) {
    return { text: `${len} de ${maxLength} caracteres (máximo)`, color: '#a4262c' };
  }
  if (len < minLength) {
    return { text: `${len} de ${minLength} caracteres (mínimo)`, color: '#a4262c' };
  }
  return { text: `${len} de ${maxLength} caracteres (máximo)`, color: '#107c10' };
}

export interface IDynamicListFormProps {
  listTitle: string;
  listWebServerRelativeUrl?: string;
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
  filesByFolderNodeId?: Record<string, File[]>,
  onUploadProgress?: (info: { folderLabel: string; fileName: string }) => void,
  listWebServerRelativeUrl?: string
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
        { itemFieldValues: iv, onUploadFileStart: onUploadProgress }
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
        onUploadFileStart: onUploadProgress,
      }
    );
    return;
  }
  const sp = getSPForWeb(listWebServerRelativeUrl);
  const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(listTitle);
  const list = isGuid ? sp.web.lists.getById(listTitle) : sp.web.lists.getByTitle(listTitle);
  const item = list.items.getById(itemId) as unknown as {
    attachmentFiles: { add(name: string, content: ArrayBuffer): Promise<unknown> };
  };
  for (let i = 0; i < files.length; i++) {
    onUploadProgress?.({ folderLabel: 'Anexos ao item', fileName: files[i].name });
    const buf = await files[i].arrayBuffer();
    await item.attachmentFiles.add(files[i].name, buf);
  }
}

async function uploadLinkedChildPendingAfterSync(
  configs: IFormLinkedChildFormConfig[],
  formManager: IFormManagerConfig,
  syncedRowsByConfigId: Record<string, ILinkedChildRowState[]>,
  pendingByKey: Record<string, File[]>
): Promise<string[]> {
  const keysToClear: string[] = [];
  for (let ci = 0; ci < configs.length; ci++) {
    const cfg = configs[ci];
    const resolved = resolveLinkedChildAttachmentRuntime(cfg, formManager);
    if (resolved.kind === 'none') continue;
    const listTitle = cfg.listTitle.trim();
    if (!listTitle) continue;
    const rows = syncedRowsByConfigId[cfg.id] ?? [];
    for (let ri = 0; ri < rows.length; ri++) {
      const row = rows[ri];
      const sid = row.sharePointId;
      if (sid === undefined || !isFinite(sid)) continue;
      const iv = { ...row.values, Id: sid };
      if (resolved.kind === 'itemAttachments') {
        const key = linkedChildAttPendingKey(cfg.id, row.localKey, '');
        const files = pendingByKey[key] ?? [];
        if (!files.length) continue;
        await uploadListItemAttachments(
          listTitle,
          sid,
          files,
          linkedChildFormAsManagerConfig(cfg),
          iv
        );
        keysToClear.push(key);
        continue;
      }
      const minimal = buildMinimalFormManagerForLinkedLibraryUpload(resolved);
      const tree = resolved.folderTree;
      const multi = !!(tree?.length && treeHasPerStepFolderUploaders(tree));
      if (multi && tree) {
        const byFolder: Record<string, File[]> = {};
        for (const node of flattenFolderTreeNodes(tree)) {
          const key = linkedChildAttPendingKey(cfg.id, row.localKey, node.id);
          const files = pendingByKey[key] ?? [];
          if (files.length) byFolder[node.id] = files;
        }
        if (!Object.keys(byFolder).some((k) => byFolder[k].length)) continue;
        await uploadListItemAttachments(listTitle, sid, [], minimal, iv, byFolder);
        for (const nid of Object.keys(byFolder)) {
          if (byFolder[nid].length) keysToClear.push(linkedChildAttPendingKey(cfg.id, row.localKey, nid));
        }
      } else {
        const key = linkedChildAttPendingKey(cfg.id, row.localKey, '');
        const files = pendingByKey[key] ?? [];
        if (!files.length) continue;
        await uploadListItemAttachments(listTitle, sid, files, minimal, iv);
        keysToClear.push(key);
      }
    }
  }
  return keysToClear;
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
      fileUrl = ensureAbsoluteSharePointUrl(sr);
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
      if (!useExpr && dynamicContext && trimmed.indexOf('[') !== -1) useExpr = true;
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

function confirmKindToIconSpec(kind: TFormCustomButtonConfirmKind | undefined): { iconName: string; color: string } {
  switch (kind) {
    case 'success':
      return { iconName: 'CompletedSolid', color: '#107c10' };
    case 'warning':
      return { iconName: 'WarningSolid', color: '#d83b01' };
    case 'error':
      return { iconName: 'StatusErrorFull', color: '#a4262c' };
    case 'blocked':
      return { iconName: 'Blocked2Solid', color: '#605e5c' };
    case 'info':
    default:
      return { iconName: 'InfoSolid', color: '#0078d4' };
  }
}

/** Modal centrado: círculo do ícone + botão principal destaque (vermelho) para ações críticas. */
function confirmKindToCenteredModalPalette(kind: TFormCustomButtonConfirmKind | undefined): {
  circleBg: string;
  iconName: string;
  iconColor: string;
  confirmDanger: boolean;
} {
  const base = confirmKindToIconSpec(kind);
  switch (kind) {
    case 'warning':
      return { circleBg: '#FEE2E2', iconName: base.iconName, iconColor: '#DC2626', confirmDanger: true };
    case 'error':
      return { circleBg: '#FEE2E2', iconName: base.iconName, iconColor: '#DC2626', confirmDanger: true };
    case 'success':
      return { circleBg: '#DCFCE7', iconName: base.iconName, iconColor: '#15803D', confirmDanger: false };
    case 'blocked':
      return { circleBg: '#F3F4F6', iconName: base.iconName, iconColor: '#4B5563', confirmDanger: false };
    case 'info':
    default:
      return { circleBg: '#DBEAFE', iconName: base.iconName, iconColor: '#2563EB', confirmDanger: false };
  }
}

type TRunTimelineStepStatus = 'pending' | 'running' | 'done' | 'error';

interface IRunTimelineDialogState {
  open: boolean;
  failed: boolean;
  title: string;
  steps: Array<{ id: string; label: string; status: TRunTimelineStepStatus }>;
  /** Detalhe do passo em execução (ex.: pasta e nome do ficheiro no upload). */
  runningDetail?: string;
}

function actionLogWillRunForTimeline(al: IFormManagerActionLogConfig | undefined): boolean {
  return (
    al?.captureEnabled === true &&
    !!(al.listTitle ?? '').trim() &&
    !!(al.actionFieldInternalName ?? '').trim()
  );
}

function finishStepLabelForTimeline(btn: IFormCustomButtonConfig): string {
  const f = btn.finishAfterRun;
  if (!f) return 'Fechar painel';
  if (f.kind === 'redirect') return 'Redirecionar para o destino';
  if (f.kind === 'clearForm') return 'Limpar formulário';
  return 'Concluir';
}

interface IFormButtonRunTimelineCtx {
  hasLinkedChildren: boolean;
  actionLogWillRun: boolean;
  hasPendingAttachments: boolean;
  permissionBreakWillRun: boolean;
}

function buildCustomButtonRunTimelineLabels(
  btn: IFormCustomButtonConfig,
  ctx: IFormButtonRunTimelineCtx
): string[] {
  const out: string[] = [];
  const op: TFormCustomButtonOperation = btn.operation ?? 'legacy';
  const actions = op === 'redirect' ? [] : btn.actions ?? [];

  if (op === 'history') {
    out.push('Abrir histórico do item');
    return out;
  }

  if (actions.length > 0) {
    out.push('Aplicar alterações nos campos (ações do botão)');
  }

  if (op === 'redirect') {
    if (ctx.actionLogWillRun) out.push('Registar ação no histórico');
    out.push('Redirecionar para o destino');
    return out;
  }

  if (op === 'add') {
    out.push('Validar dados do formulário');
    out.push('Criar item na lista');
    if (ctx.hasPendingAttachments) out.push('Enviar anexos');
    if (ctx.hasLinkedChildren) out.push('Sincronizar listas vinculadas');
    if (ctx.permissionBreakWillRun) out.push('Aplicar quebra de permissões');
    if (ctx.actionLogWillRun) out.push('Registar ação no histórico');
    out.push(finishStepLabelForTimeline(btn));
    return out;
  }

  if (op === 'update') {
    out.push('Validar dados do formulário');
    out.push('Atualizar item na lista');
    if (ctx.hasPendingAttachments) out.push('Enviar anexos');
    if (ctx.hasLinkedChildren) out.push('Sincronizar listas vinculadas');
    if (ctx.permissionBreakWillRun) out.push('Aplicar quebra de permissões');
    if (ctx.actionLogWillRun) out.push('Registar ação no histórico');
    out.push(finishStepLabelForTimeline(btn));
    return out;
  }

  if (op === 'delete') {
    out.push('Eliminar item na lista');
    if (ctx.actionLogWillRun) out.push('Registar ação no histórico');
    out.push(finishStepLabelForTimeline(btn));
    return out;
  }

  const behavior = btn.behavior ?? 'actionsOnly';
  if (behavior === 'close') {
    if (ctx.actionLogWillRun) out.push('Registar ação no histórico');
    out.push('Fechar painel');
    return out;
  }
  if (behavior === 'draft') {
    out.push('Validar e guardar rascunho');
    if (ctx.actionLogWillRun) out.push('Registar ação no histórico');
    out.push(finishStepLabelForTimeline(btn));
    return out;
  }
  if (behavior === 'submit') {
    out.push(
      ctx.permissionBreakWillRun
        ? 'Validar, gravar e aplicar quebra de permissões'
        : 'Validar e guardar envio'
    );
    if (ctx.actionLogWillRun) out.push('Registar ação no histórico');
    out.push(finishStepLabelForTimeline(btn));
    return out;
  }
  if (ctx.actionLogWillRun) out.push('Registar ação no histórico');
  out.push(finishStepLabelForTimeline(btn));
  return out;
}

interface IRunTimelineCtl {
  enter: (i: number) => void;
  ok: (i: number) => void;
  err: (i: number) => void;
  closeSuccess: () => void;
  closeError: () => void;
  setRunningDetail: (detail: string | undefined) => void;
}

function createRunTimelineController(
  setState: React.Dispatch<React.SetStateAction<IRunTimelineDialogState | null>>,
  labels: string[],
  title: string
): IRunTimelineCtl | null {
  if (labels.length === 0) return null;
  const patch = (i: number, phase: 'enter' | 'ok' | 'err'): void => {
    setState((prev) => {
      if (!prev) return prev;
      const steps = prev.steps.map((s, idx) => {
        if (phase === 'enter') {
          if (idx < i) return { ...s, status: 'done' as const };
          if (idx === i) return { ...s, status: 'running' as const };
          return { ...s, status: 'pending' as const };
        }
        if (idx === i) {
          return { ...s, status: phase === 'ok' ? ('done' as const) : ('error' as const) };
        }
        return s;
      });
      return { ...prev, steps, runningDetail: undefined };
    });
  };
  flushSync(() => {
    setState({
      open: true,
      failed: false,
      title,
      runningDetail: undefined,
      steps: labels.map((label, idx) => ({
        id: `tl_${idx}`,
        label,
        status: 'pending' as const,
      })),
    });
  });
  return {
    enter: (i) => patch(i, 'enter'),
    ok: (i) => patch(i, 'ok'),
    err: (i) => patch(i, 'err'),
    setRunningDetail: (detail: string | undefined): void => {
      setState((prev) => (prev ? { ...prev, runningDetail: detail } : prev));
    },
    closeSuccess: () => {
      setState(null);
    },
    closeError: () => {
      setState((prev) => (prev ? { ...prev, failed: true, runningDetail: undefined } : prev));
    },
  };
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

function ConfirmPromptFieldEditor(props: {
  meta: IFieldMetadata;
  editor: IConfirmPromptEditorState;
  onChange: (next: IConfirmPromptEditorState) => void;
  /** Campos alinhados ao modal centrado (largura total, entre texto e botões). */
  modalSurface?: boolean;
}): React.ReactElement {
  const { meta, editor, onChange, modalSurface } = props;
  const theme = useTheme();
  const tfStyles = {
    root: {
      marginBottom: 0,
      ...(modalSurface ? { width: '100%' } : {}),
    },
    label: {
      root: {
        fontWeight: '600',
        color: theme.palette.neutralPrimary,
        marginBottom: 6,
      },
    },
    fieldGroup: {
      borderRadius: 10,
      border: `1px solid ${theme.palette.neutralQuaternaryAlt}`,
      backgroundColor: theme.palette.white,
      ':hover': { borderColor: theme.palette.neutralTertiaryAlt },
      selectors: {
        '&.ms-TextField-fieldGroup': { borderRadius: 10 },
      },
    },
    field: { borderRadius: 10 },
    ...(modalSurface ? { wrapper: { width: '100%' } } : {}),
  };
  const ddStyles = {
    dropdown: {
      borderRadius: 10,
      border: `1px solid ${theme.palette.neutralQuaternaryAlt}`,
    },
    title: {
      root: {
        fontWeight: '600',
        color: theme.palette.neutralPrimary,
      },
    },
  };
  const wrapModal = (node: React.ReactElement): React.ReactElement =>
    modalSurface ? (
      <Stack styles={{ root: { width: '100%' } }} tokens={{ childrenGap: 10 }}>
        {node}
      </Stack>
    ) : (
      node
    );
  switch (meta.MappedType) {
    case 'boolean':
      return wrapModal(
        <Stack
          horizontal
          verticalAlign="center"
          horizontalAlign="space-between"
          styles={{
            root: {
              padding: '12px 14px',
              borderRadius: 10,
              border: `1px solid ${theme.palette.neutralLight}`,
              backgroundColor: theme.palette.white,
              width: modalSurface ? '100%' : undefined,
              boxSizing: 'border-box',
            },
          }}
        >
          <Text variant="medium" styles={{ root: { fontWeight: 600 } }}>
            {meta.Title}
          </Text>
          <Toggle
            checked={editor.bool}
            onText="Sim"
            offText="Não"
            onChange={(_, c) => onChange({ ...editor, bool: !!c })}
          />
        </Stack>
      );
    case 'number':
    case 'currency':
      return wrapModal(
        <TextField
          label={meta.Title}
          type="number"
          value={editor.text}
          onChange={(_, v) => onChange({ ...editor, text: v ?? '' })}
          styles={tfStyles}
        />
      );
    case 'datetime':
      return wrapModal(
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { width: modalSurface ? '100%' : undefined } }}>
          <Label styles={{ root: tfStyles.label.root }}>{meta.Title}</Label>
          <DatePicker
            {...FLUENT_DATE_PICKER_PT_BR}
            value={editor.dateIso ? new Date(editor.dateIso) : undefined}
            onSelectDate={(d) => onChange({ ...editor, dateIso: d ? d.toISOString() : null })}
            textField={{
              styles: {
                fieldGroup: tfStyles.fieldGroup,
                field: tfStyles.field,
                wrapper: { width: '100%' },
              },
            }}
          />
        </Stack>
      );
    case 'choice': {
      const raw = (meta.Choices ?? []).map((c) => ({ key: c, text: c }));
      const opts: IDropdownOption[] = raw.length ? raw : [{ key: '', text: '—' }];
      return wrapModal(
        <Dropdown
          label={meta.Title}
          options={opts}
          selectedKey={editor.choiceKey ? editor.choiceKey : undefined}
          onChange={(_, o) => onChange({ ...editor, choiceKey: o ? String(o.key) : '' })}
          styles={{
            ...ddStyles,
            dropdown: { ...ddStyles.dropdown, width: '100%' },
          }}
        />
      );
    }
    case 'multiline':
      return wrapModal(
        <TextField
          label={meta.Title}
          multiline
          autoAdjustHeight
          resizable={false}
          rows={4}
          value={editor.text}
          onChange={(_, v) => onChange({ ...editor, text: v ?? '' })}
          styles={{
            ...tfStyles,
            fieldGroup: { ...tfStyles.fieldGroup, alignItems: 'stretch' },
          }}
        />
      );
    default:
      return wrapModal(
        <TextField
          label={meta.Title}
          value={editor.text}
          onChange={(_, v) => onChange({ ...editor, text: v ?? '' })}
          styles={tfStyles}
        />
      );
  }
}

export const DynamicListForm: React.FC<IDynamicListFormProps> = ({
  listTitle,
  listWebServerRelativeUrl,
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
  const listWeb = listWebServerRelativeUrl?.trim() || undefined;
  const theme = useTheme();
  const stepAccentHex = useMemo(
    () => resolveStepUiAccentColor(theme, formManager.stepAccentPaletteSlot),
    [theme, formManager.stepAccentPaletteSlot]
  );
  const fieldConfigs: IFormFieldConfig[] =
    formManager.fields.length > 0
      ? formManager.fields
      : fieldMetadata
          .filter((f) => !f.Hidden && !f.ReadOnlyField && f.InternalName !== 'Id')
          .map((f) => ({ internalName: f.InternalName, sectionId: FORM_OCULTOS_STEP_ID }));
  const referencedPayloadOnlyNames = useMemo(
    () => collectFormManagerReferencedPayloadFieldNames(formManager),
    [formManager]
  );
  const names = useMemo(() => {
    const base = fieldConfigs
      .filter((f) => f.internalName !== FORM_ATTACHMENTS_FIELD_INTERNAL && !isFormBannerFieldConfig(f))
      .map((f) => f.internalName);
    const baseSet = new Set(base);
    const extras = referencedPayloadOnlyNames.filter((n) => !baseSet.has(n));
    if (extras.length === 0) return base;
    extras.sort();
    return [...base, ...extras];
  }, [fieldConfigs, referencedPayloadOnlyNames]);
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
  const isDateTimeFieldFromMeta = useCallback(
    (internalName: string): boolean => metaByName.get(internalName)?.MappedType === 'datetime',
    [metaByName]
  );
  const fieldConfigByInternalName = useMemo(
    () => new Map(fieldConfigs.map((fc) => [fc.internalName, fc])),
    [fieldConfigs]
  );
  const lookupDestMetaCacheRef = useRef<Record<string, IFieldMetadata[]>>({});
  useEffect(() => {
    lookupDestMetaCacheRef.current = {};
  }, [listWeb]);
  const fieldLabelByName = useMemo(
    () => buildFormFieldLabelMap(fieldConfigs, metaByName),
    [fieldConfigs, metaByName]
  );

  const setComputedExprSnapRef = useRef<{ openKey: string; snap: Record<string, string> }>({
    openKey: '',
    snap: {},
  });
  const setComputedItemOpenKey =
    formMode !== 'create' &&
    itemId !== undefined &&
    itemId !== null &&
    typeof itemId === 'number' &&
    isFinite(itemId)
      ? `${listTitle}\t${listWebServerRelativeUrl ?? ''}\t${itemId}\t${formMode}`
      : '';
  if (setComputedItemOpenKey !== setComputedExprSnapRef.current.openKey) {
    const snap: Record<string, string> = {};
    if (setComputedItemOpenKey) {
      const rules = formManager.rules ?? [];
      for (let i = 0; i < rules.length; i++) {
        const r = rules[i];
        if (r.enabled === false) continue;
        if (r.action !== 'setComputed') continue;
        const f = r.field.trim();
        if (!f) continue;
        snap[f] = r.expression.trim();
      }
    }
    setComputedExprSnapRef.current = { openKey: setComputedItemOpenKey, snap };
  }

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
  const [requiredValidationModalSections, setRequiredValidationModalSections] = useState<
    ReturnType<typeof buildValidationModalSections> | null
  >(null);
  const [lookupOptions, setLookupOptions] = useState<Record<string, IDropdownOption[]>>({});
  const [pendingFiles, setPendingFiles] = useState<File[]>([]);
  const [pendingFilesByFolder, setPendingFilesByFolder] = useState<Record<string, File[]>>({});
  const [attachmentCount, setAttachmentCount] = useState(0);
  const [serverAttachments, setServerAttachments] = useState<IServerAttachmentRow[]>([]);
  const prevByTriggerRef = useRef<Record<string, unknown>>({});
  const prevLinkedParentItemIdRef = useRef<number | undefined>(undefined);
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
  const buildActionLogRuntimeCtx = useCallback(
    (btn: IFormCustomButtonConfig, sourceItemId: number | null | undefined): IFormActionLogRuntimeContext => ({
      sourceListTitle: listTitle,
      sourceItemId,
      formMode,
      logEntryAccentHex: resolveActionLogPaletteAccentHex(
        theme,
        formManager.actionLog?.descriptionPaletteSlotByButtonId?.[btn.id]
      ),
    }),
    [listTitle, formMode, theme, formManager.actionLog]
  );
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
  const [linkedPendingByKey, setLinkedPendingByKey] = useState<Record<string, File[]>>({});
  const [linkedServerAttachmentsByKey, setLinkedServerAttachmentsByKey] = useState<
    Record<string, IServerAttachmentRow[]>
  >({});

  const historyLogEntryPaletteContext = useMemo(
    () => ({
      slotByButtonId: formManager.actionLog?.descriptionPaletteSlotByButtonId ?? {},
      customButtons: (formManager.customButtons ?? []).map((b) => ({ id: b.id, label: b.label })),
      historyButtonLabel: (formManager.historyButtonLabel ?? 'Histórico').trim() || 'Histórico',
    }),
    [
      formManager.actionLog?.descriptionPaletteSlotByButtonId,
      formManager.customButtons,
      formManager.historyButtonLabel,
    ]
  );

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
    if (formMode === 'create') {
      prevLinkedParentItemIdRef.current = undefined;
      return;
    }
    if (itemId === undefined || itemId === null) {
      prevLinkedParentItemIdRef.current = undefined;
      return;
    }
    if (prevLinkedParentItemIdRef.current !== itemId) {
      setLinkedRowsById({});
      setLinkedBaselineById({});
      prevLinkedParentItemIdRef.current = itemId;
    }
  }, [formMode, itemId]);

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
        const validIds = new Set(linkedConfigsSorted.map((c) => c.id));
        setLinkedRowsById((prev) => {
          const merged = { ...prev };
          for (const id of Object.keys(nextRows)) {
            merged[id] = nextRows[id];
          }
          for (const k of Object.keys(merged)) {
            if (!validIds.has(k)) delete merged[k];
          }
          return merged;
        });
        setLinkedBaselineById((prev) => {
          const merged = { ...prev };
          for (const id of Object.keys(nextBase)) {
            merged[id] = nextBase[id];
          }
          for (const k of Object.keys(merged)) {
            if (!validIds.has(k)) delete merged[k];
          }
          return merged;
        });
      }
    })();
    return (): void => {
      cancel = true;
    };
  }, [formMode, itemId, linkedConfigsSorted, linkedMetaById, itemsService]);

  useEffect(() => {
    if (linkedConfigsSorted.length === 0) {
      setLinkedServerAttachmentsByKey({});
      return;
    }
    let cancelled = false;
    void (async (): Promise<void> => {
      const out: Record<string, IServerAttachmentRow[]> = {};
      const sp = getSP();
      for (let ci = 0; ci < linkedConfigsSorted.length; ci++) {
        if (cancelled) return;
        const cfg = linkedConfigsSorted[ci];
        const resolved = resolveLinkedChildAttachmentRuntime(cfg, formManager);
        if (resolved.kind === 'none') continue;
        const rows = linkedRowsById[cfg.id] ?? [];
        const childList = cfg.listTitle.trim();
        if (!childList) continue;
        for (let ri = 0; ri < rows.length; ri++) {
          if (cancelled) return;
          const row = rows[ri];
          const sid = row.sharePointId;
          if (sid === undefined || sid === null || !Number.isFinite(sid)) continue;
          if (resolved.kind === 'itemAttachments') {
            const flatKey = linkedChildAttPendingKey(cfg.id, row.localKey, '');
            try {
              const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(childList);
              const list = isGuid ? sp.web.lists.getById(childList) : sp.web.lists.getByTitle(childList);
              const item = list.items.getById(sid) as unknown as { attachmentFiles(): Promise<unknown> };
              const raw = await item.attachmentFiles();
              const mapped = mapServerAttachments(normalizeSharePointAttachmentFiles(raw));
              out[flatKey] = mapped;
            } catch {
              out[flatKey] = [];
            }
            continue;
          }
          if (resolved.kind === 'documentLibrary') {
            try {
              const all = await loadLibraryAttachmentRowsForMainItem(
                resolved.libraryTitle,
                resolved.lookupFieldInternalName,
                sid
              );
              const mapped: IServerAttachmentRow[] = all.map((x) => ({
                fileName: x.fileName,
                fileUrl: x.fileUrl,
                fileRef: x.fileRef,
              }));
              const tree = resolved.folderTree;
              const multiLib = !!tree?.length && treeHasPerStepFolderUploaders(tree);
              if (multiLib) {
                const nodes = flattenFolderTreeNodes(tree!).filter((n) => n.uploadTarget);
                if (nodes.length > 0) {
                  for (let ni = 0; ni < nodes.length; ni++) {
                    const node = nodes[ni];
                    const pk = linkedChildAttPendingKey(cfg.id, row.localKey, node.id);
                    out[pk] = mapped.filter(
                      (m) =>
                        typeof m.fileRef === 'string' &&
                        m.fileRef.trim() &&
                        libraryFileRowBelongsToFolderNode(m.fileRef.trim(), node.id, tree!, sid, row.values)
                    );
                  }
                } else {
                  out[linkedChildAttPendingKey(cfg.id, row.localKey, '')] = mapped;
                }
              } else {
                out[linkedChildAttPendingKey(cfg.id, row.localKey, '')] = mapped;
              }
            } catch {
              out[linkedChildAttPendingKey(cfg.id, row.localKey, '')] = [];
            }
          }
        }
      }
      if (!cancelled) setLinkedServerAttachmentsByKey(out);
    })();
    return (): void => {
      cancelled = true;
    };
  }, [linkedConfigsSorted, linkedRowsById, formManager]);

  useEffect(() => {
    if (formMode !== 'create') return;
    setValues((prev) => {
      const merged = getDefaultValuesFromRules(formManager, prev, dynamicContext, {
        isDateTimeField: isDateTimeFieldFromMeta,
      });
      return merged;
    });
  }, [formManager, formMode, dynamicContext, isDateTimeFieldFromMeta]);

  useEffect(() => {
    setValues((prev) => applyTextTransformsToRecordValues(prev, fieldConfigs, metaByName));
  }, [values, fieldConfigs, metaByName]);

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
      dynamicContext: withRuleRuntimeDynamicContext(dynamicContext, currentUserId),
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
      buildFormDerivedState(
        formManager,
        fieldConfigs,
        runtimeCtx(),
        {
          show: buttonOverlay.show,
          hide: buttonOverlay.hide,
        },
        metaByName
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
      metaByName,
    ]
  );

  const validateValueLengthMergedByField = useMemo(() => {
    const rules = formManager.rules ?? [];
    const vis = derived.fieldVisible;
    const out: Record<string, { minLength?: number; maxLength?: number }> = {};
    const ctxSlice = {
      formMode,
      values,
      userGroupTitles,
      dynamicContext,
    };
    for (let i = 0; i < fieldConfigs.length; i++) {
      const n = fieldConfigs[i].internalName;
      if (vis[n] === false) continue;
      const b = getMergedValidateValueLengthBounds(rules, n, ctxSlice, vis);
      if (b && (b.minLength !== undefined || b.maxLength !== undefined)) {
        out[n] = { minLength: b.minLength, maxLength: b.maxLength };
      }
    }
    return out;
  }, [formManager.rules, fieldConfigs, formMode, values, userGroupTitles, dynamicContext, derived.fieldVisible]);

  const validateValueNumberMergedByField = useMemo(() => {
    const rules = formManager.rules ?? [];
    const vis = derived.fieldVisible;
    const out: Record<string, { minNumber?: number; maxNumber?: number }> = {};
    const ctxSlice = {
      formMode,
      values,
      userGroupTitles,
      dynamicContext,
    };
    for (let i = 0; i < fieldConfigs.length; i++) {
      const n = fieldConfigs[i].internalName;
      if (vis[n] === false) continue;
      const b = getMergedValidateValueNumberBounds(rules, n, ctxSlice, vis);
      if (b && (b.minNumber !== undefined || b.maxNumber !== undefined)) {
        out[n] = { minNumber: b.minNumber, maxNumber: b.maxNumber };
      }
    }
    return out;
  }, [formManager.rules, fieldConfigs, formMode, values, userGroupTitles, dynamicContext, derived.fieldVisible]);

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

  const visibleCustomButtons = useMemo(() => {
    return (formManager.customButtons ?? [])
      .filter((b) =>
        formManager.historyEnabled === true ? (b.operation ?? 'legacy') !== 'history' : true
      )
      .filter((b) =>
        shouldShowCustomButton(b, runtimeCtx(), {
          allRequiredFilled,
          historyEnabledInConfig: formManager.historyEnabled === true,
          historyItemId: itemId,
        })
      );
  }, [
    formManager.customButtons,
    formManager.historyEnabled,
    runtimeCtx,
    allRequiredFilled,
    itemId,
  ]);

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
    async (fieldName: string, lf?: { parentField: string; childField?: string; filterOperator?: string; odataFilterTemplate?: string }): Promise<void> => {
      const m = metaByName.get(fieldName);
      if (!m?.LookupList) return;
      const fc = fieldConfigByInternalName.get(fieldName);
      try {
        let filter: string | undefined;

        const listGuid = m.LookupList;
        let fieldMetaList: IFieldMetadata[] | undefined = lookupDestMetaCacheRef.current[listGuid];
        if (!fieldMetaList?.length) {
          try {
            const fetched = await fieldsService.getFields(listGuid, listWeb);
            lookupDestMetaCacheRef.current = {
              ...lookupDestMetaCacheRef.current,
              [listGuid]: fetched,
            };
            fieldMetaList = fetched;
          } catch {
            fieldMetaList = undefined;
          }
        }

        if (lf) {
          if (lf.childField && lf.filterOperator) {
            const parentVal = values[lf.parentField];
            const parentMeta = metaByName.get(lf.parentField);
            const childFieldMeta = fieldMetaList?.find((x) => x.InternalName === lf.childField);
            filter = buildLookupODataFilter(lf.childField, lf.filterOperator, parentVal, parentMeta, childFieldMeta);
          } else if (lf.odataFilterTemplate) {
            const pid = lookupIdFromValue(values[lf.parentField]);
            if (pid !== undefined) filter = lf.odataFilterTemplate.split('{parent}').join(String(pid));
          }
        }

        const selectRaw = buildLookupDropdownSelectRaw(m, fc ?? {});
        const labelFieldName = resolveLookupFormLabelInternalName(m, fc ?? {});
        const labelMeta = fieldMetaList?.find((x) => x.InternalName === labelFieldName);

        const rows = await itemsService.getItems<Record<string, unknown>>(m.LookupList, {
          select: selectRaw,
          filter,
          top: 200,
          ...(listWeb ? { webServerRelativeUrl: listWeb } : {}),
          ...(fieldMetaList?.length ? { fieldMetadata: fieldMetaList } : {}),
        });

        const opts: IDropdownOption[] = [
          { key: '', text: '—' },
          ...rows.map((row) => ({
            key: String(row.Id),
            text: lookupRowToOptionText(row, labelFieldName, labelMeta, fc?.lookupOptionLabelSubProp),
            data: row,
          })),
        ];
        setLookupOptions((o) => ({ ...o, [fieldName]: opts }));
      } catch {
        setLookupOptions((o) => ({ ...o, [fieldName]: [] }));
      }
    },
    [itemsService, metaByName, listWeb, fieldsService, fieldConfigByInternalName, values]
  );

  const lookupFetchKey = useMemo(() => {
    const parts: string[] = [];
    for (let i = 0; i < fieldConfigs.length; i++) {
      const fn = fieldConfigs[i].internalName;
      const m = metaByName.get(fn);
      if (m?.MappedType !== 'lookup' && m?.MappedType !== 'lookupmulti') continue;
      const listId = String(m.LookupList ?? '');
      const fc = fieldConfigs[i];
      const labelDisp = resolveLookupFormLabelInternalName(m, fc ?? {});
      const extrasSig = JSON.stringify(fc?.lookupOptionExtraSelectFields ?? []);
      const subPropSig = fc?.lookupOptionLabelSubProp ?? '';
      const detailSig = JSON.stringify(fc?.lookupOptionDetailBelowFields ?? []);
      const lf = derived.lookupFilters[fn];
      if (lf) {
        const parentVal = values[lf.parentField];
        const parentId = lookupIdFromValue(parentVal);
        const parentSig = parentId !== undefined ? String(parentId) :
          typeof parentVal === 'string' ? parentVal :
          typeof parentVal === 'number' ? String(parentVal) : '';
        parts.push(
          `${fn}\t${listId}\t${labelDisp}\t${extrasSig}\t${subPropSig}\t${detailSig}\t${lf.parentField}\t${lf.childField ?? ''}\t${lf.filterOperator ?? ''}\t${parentSig}`
        );
      } else {
        parts.push(`${fn}\t${listId}\t${labelDisp}\t${extrasSig}\t${subPropSig}\t${detailSig}\t`);
      }
    }
    parts.sort();
    return parts.join('\n');
  }, [fieldConfigs, metaByName, derived.lookupFilters, values]);

  const lookupDetailSnapshot = useMemo(() => {
    const out: Record<string, Record<string, unknown> | Record<string, unknown>[] | undefined> = {};
    for (let i = 0; i < fieldConfigs.length; i++) {
      const fc = fieldConfigs[i];
      const detailFns = fc.lookupOptionDetailBelowFields ?? [];
      if (!detailFns.length) continue;
      const m = metaByName.get(fc.internalName);
      if (!m || (m.MappedType !== 'lookup' && m.MappedType !== 'lookupmulti')) continue;
      const opts = lookupOptions[fc.internalName] ?? [];
      if (m.MappedType === 'lookup') {
        const id = lookupIdFromValue(values[fc.internalName]);
        if (!id) {
          out[fc.internalName] = undefined;
          continue;
        }
        const opt = opts.find((o) => String(o.key) === String(id));
        const data = opt && typeof opt === 'object' && 'data' in opt ? (opt as { data?: Record<string, unknown> }).data : undefined;
        out[fc.internalName] = data;
      } else {
        const sel = normalizeIdTitleArray(values[fc.internalName]);
        const many: Record<string, unknown>[] = [];
        for (let s = 0; s < sel.length; s++) {
          const opt = opts.find((o) => String(o.key) === String(sel[s].Id));
          const data =
            opt && typeof opt === 'object' && 'data' in opt ? (opt as { data?: Record<string, unknown> }).data : undefined;
          if (data) many.push(data);
        }
        out[fc.internalName] = many.length ? many : undefined;
      }
    }
    return out;
  }, [fieldConfigs, metaByName, lookupOptions, values]);

  useEffect(() => {
    let cancelled = false;
    void (async (): Promise<void> => {
      for (let i = 0; i < fieldConfigs.length; i++) {
        if (cancelled) return;
        const fn = fieldConfigs[i].internalName;
        const m = metaByName.get(fn);
        if (m?.MappedType === 'lookup' || m?.MappedType === 'lookupmulti') {
          const lf = derived.lookupFilters[fn];
          await loadLookupOptions(fn, lf);
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

  const applyDateFieldSelect = useCallback(
    (name: string, d: Date | null | undefined) => {
      if (d === null || d === undefined) {
        updateField(name, null);
        setLocalErrors((prev) => {
          if (!prev[name]) return prev;
          const { [name]: _, ...rest } = prev;
          return rest;
        });
        return;
      }
      const iso = d.toISOString();
      const nextValues = { ...values, [name]: iso };
      const msg = evaluateValidateDateRulesForField(formManager.rules ?? [], name, nextValues, {
        formMode,
        submitKind: undefined,
        userGroupTitles,
        dynamicContext,
        fieldVisible: (fn) => derived.fieldVisible[fn] !== false,
        now: new Date(),
      });
      if (msg) {
        updateField(name, null);
        setLocalErrors((prev) => ({ ...prev, [name]: msg }));
        return;
      }
      updateField(name, iso);
      setLocalErrors((prev) => {
        if (!prev[name]) return prev;
        const { [name]: _, ...rest } = prev;
        return rest;
      });
    },
    [values, formManager.rules, formMode, userGroupTitles, dynamicContext, derived]
  );

  const reloadLinkedRowsForParent = useCallback(
    async (parentId: number): Promise<Record<string, ILinkedChildRowState[]>> => {
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
      const validIds = new Set(linkedConfigsSorted.map((c) => c.id));
      setLinkedRowsById((prev) => {
        const merged = { ...prev };
        for (const id of Object.keys(nextRows)) {
          merged[id] = nextRows[id];
        }
        for (const k of Object.keys(merged)) {
          if (!validIds.has(k)) delete merged[k];
        }
        return merged;
      });
      setLinkedBaselineById((prev) => {
        const merged = { ...prev };
        for (const id of Object.keys(nextBase)) {
          merged[id] = nextBase[id];
        }
        for (const k of Object.keys(merged)) {
          if (!validIds.has(k)) delete merged[k];
        }
        return merged;
      });
      return nextRows;
    },
    [linkedConfigsSorted, linkedMetaById, itemsService]
  );

  const performLinkedSync = useCallback(
    async (parentId: number): Promise<Record<string, ILinkedChildRowState[]> | undefined> => {
      if (!linkedConfigsSorted.length) return undefined;
      const syncedRows = await syncAllLinkedChildLists(
        itemsService,
        linkedConfigsSorted,
        parentId,
        linkedRowsById,
        linkedMetaById,
        linkedBaselineById
      );
      const toClear = await uploadLinkedChildPendingAfterSync(
        linkedConfigsSorted,
        formManager,
        syncedRows,
        linkedPendingByKey
      );
      if (toClear.length) {
        setLinkedPendingByKey((prev) => {
          const next = { ...prev };
          for (let i = 0; i < toClear.length; i++) delete next[toClear[i]];
          return next;
        });
      }
      const reloaded = await reloadLinkedRowsForParent(parentId);
      setLinkedRowErrorsById({});
      return reloaded;
    },
    [
      linkedConfigsSorted,
      itemsService,
      linkedRowsById,
      linkedMetaById,
      linkedBaselineById,
      reloadLinkedRowsForParent,
      formManager,
      linkedPendingByKey,
    ]
  );

  const runPermissionBreakAfterSubmit = useCallback(
    async (
      parentId: number,
      valuesForPerm: Record<string, unknown>,
      linkedRowsSnapshot: Record<string, ILinkedChildRowState[]>,
      onProgress?: (detail: string) => void
    ): Promise<void> => {
      if (!formManager.permissionBreak?.enabled) return;
      await applyFormManagerPermissionBreak({
        formManager,
        listTitle,
        mainListWebServerRelativeUrl: listWeb,
        mainItemId: parentId,
        mainValues: { ...valuesForPerm, Id: parentId },
        mainAuthorId: pickMainAuthorId(valuesForPerm, initialItem, currentUserId),
        linkedConfigsSorted,
        linkedRowsById: linkedRowsSnapshot,
        onProgress,
      });
    },
    [formManager, listTitle, listWeb, linkedConfigsSorted, initialItem, currentUserId]
  );

  type IValidateFailurePayload = {
    errors: Record<string, string>;
    linkedRowErrorsById: Record<string, Record<string, string>[]>;
    submitKind: TFormSubmitKind;
  };

  const validate = async (
    submitKind: TFormSubmitKind,
    opts?: {
      values?: Record<string, unknown>;
      buttonOverlay?: IFormButtonFieldOverlay;
    }
  ): Promise<IValidateFailurePayload | undefined> => {
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
    const sync = collectFormValidationErrors(
      formManager,
      fieldConfigs,
      ctx,
      att,
      {
        show: ov.show,
        hide: ov.hide,
      },
      metaByName
    );
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
    if (Object.keys(mergedErr).length > 0) {
      setLinkedRowErrorsById({});
      return { errors: mergedErr, linkedRowErrorsById: {}, submitKind };
    }
    const asyncErr = await runAsyncFormValidations(formManager, vals, itemsService, listTitle, itemId, submitKind);
    if (Object.keys(asyncErr).length > 0) {
      setLocalErrors(asyncErr);
      setLinkedRowErrorsById({});
      return { errors: asyncErr, linkedRowErrorsById: {}, submitKind };
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
          const linkedMetaMap = new Map(meta.map((m) => [m.InternalName, m]));
          const syncL = collectFormValidationErrors(shell, cfg.fields, ctxL, attEmpty, undefined, linkedMetaMap);
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
        return { errors: flat, linkedRowErrorsById: rowErr, submitKind };
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
    const validationOutcome = await validate(submitKind, { values: vals, buttonOverlay: ov });
    if (validationOutcome) {
      setRequiredValidationModalSections(
        buildValidationModalSections({
          mainErrors: validationOutcome.errors,
          formManager,
          fieldConfigs,
          ctx: {
            formMode,
            values: vals,
            submitKind: validationOutcome.submitKind,
            userGroupTitles,
            currentUserId,
            authorId,
            dynamicContext,
            attachmentFolderUrl,
          },
          buttonOverlay: { show: ov.show, hide: ov.hide },
          fieldLabelByName: fieldLabelByName,
          mainFieldMetaByName: metaByName,
          linkedConfigs: linkedConfigsSorted,
          linkedRowErrorsById: validationOutcome.linkedRowErrorsById,
          linkedRowsById,
          linkedMetaById,
          mainListLabel: listTitle,
        })
      );
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
      let linkedSnap = linkedRowsById;
      if (submitKind === 'submit' && linkedConfigsSorted.length > 0) {
        const parentId = savedId ?? itemId;
        if (parentId !== undefined && parentId !== null && typeof parentId === 'number' && isFinite(parentId)) {
          try {
            const rel = await performLinkedSync(parentId);
            if (rel) linkedSnap = rel;
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
      if (submitKind === 'submit') {
        const parentIdPb = savedId ?? itemId;
        if (
          parentIdPb !== undefined &&
          parentIdPb !== null &&
          typeof parentIdPb === 'number' &&
          isFinite(parentIdPb)
        ) {
          try {
            await runPermissionBreakAfterSubmit(parentIdPb, vals, linkedSnap);
          } catch (pe) {
            setFormError(
              `Gravado, mas a quebra de permissões falhou: ${pe instanceof Error ? pe.message : String(pe)}`
            );
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
  type IConfirmRunResult = { proceed: boolean; valuesBaselinePatch?: Record<string, unknown> };
  const confirmRunResolveRef = useRef<((r: IConfirmRunResult) => void) | null>(null);
  const [confirmDialogOpen, setConfirmDialogOpen] = useState(false);
  const [confirmDialogView, setConfirmDialogView] = useState<{
    kind: TFormCustomButtonConfirmKind;
    message: string;
    title: string;
    promptFieldInternalName?: string;
  } | null>(null);
  const [confirmPromptEditor, setConfirmPromptEditor] = useState<IConfirmPromptEditorState | null>(null);
  const [runTimelineDialog, setRunTimelineDialog] = useState<IRunTimelineDialogState | null>(null);

  const closeButtonConfirmDialog = useCallback(
    (ok: boolean) => {
      const viewSnapshot = confirmDialogView;
      const editorSnapshot = confirmPromptEditor;
      setConfirmDialogOpen(false);
      setConfirmDialogView(null);
      setConfirmPromptEditor(null);
      const r = confirmRunResolveRef.current;
      confirmRunResolveRef.current = null;
      if (!r) return;
      if (!ok) {
        r({ proceed: false });
        return;
      }
      const name = viewSnapshot?.promptFieldInternalName;
      if (name && editorSnapshot) {
        const meta = metaByName.get(name);
        if (meta && isConfirmPromptEligibleField(meta)) {
          if (!confirmPromptEditorIsFilled(meta, editorSnapshot)) {
            r({ proceed: false });
            return;
          }
          const val = confirmPromptEditorToValue(meta, editorSnapshot);
          r({ proceed: true, valuesBaselinePatch: { [name]: val } });
          return;
        }
      }
      r({ proceed: true });
    },
    [confirmDialogView, confirmPromptEditor, metaByName]
  );

  const confirmPromptMetaForDialog = useMemo(() => {
    const name = confirmDialogView?.promptFieldInternalName;
    if (!name) return undefined;
    const m = metaByName.get(name);
    if (!m || !isConfirmPromptEligibleField(m)) return undefined;
    return m;
  }, [confirmDialogView?.promptFieldInternalName, metaByName]);

  const confirmPrimaryDisabled =
    !!(
      confirmPromptMetaForDialog &&
      confirmPromptEditor &&
      !confirmPromptEditorIsFilled(confirmPromptMetaForDialog, confirmPromptEditor)
    );

  const confirmBeforeRunIfNeeded = useCallback(
    (btn: IFormCustomButtonConfig): Promise<IConfirmRunResult> => {
      const c = btn.confirmBeforeRun;
      const msg = (c?.message ?? '').trim();
      const promptField = (c?.promptFieldInternalName ?? '').trim();
      if (c?.enabled !== true || (!msg && !promptField)) {
        return Promise.resolve({ proceed: true });
      }
      const metaP = promptField ? metaByName.get(promptField) : undefined;
      const promptOk = !!(promptField && metaP && isConfirmPromptEligibleField(metaP));
      return new Promise((resolve) => {
        confirmRunResolveRef.current = resolve;
        setConfirmDialogView({
          kind: c?.kind ?? 'info',
          message: msg,
          title: (btn.label || btn.id).trim() || 'Confirmar',
          ...(promptOk && promptField ? { promptFieldInternalName: promptField } : {}),
        });
        if (promptOk && metaP) {
          setConfirmPromptEditor(initConfirmPromptEditor(metaP, values[promptField]));
        } else {
          setConfirmPromptEditor(null);
        }
        setConfirmDialogOpen(true);
      });
    },
    [metaByName, values]
  );

  const runFinishAfterSuccess = useCallback(
    async (
      btn: IFormCustomButtonConfig,
      valuesForUrl: Record<string, unknown>,
      itemIdForUrl?: number
    ): Promise<'none' | 'redirect' | 'cleared'> => {
      const f = btn.finishAfterRun;
      if (!f) return 'none';
      const id = itemIdForUrl !== undefined && itemIdForUrl !== null ? itemIdForUrl : itemId;
      if (f.kind === 'redirect') {
        const tpl = (f.redirectUrlTemplate ?? '').trim();
        if (!tpl) {
          setFormError('Configure o URL em «Último passo» (redirecionar).');
          return 'none';
        }
        const url = interpolateFormButtonRedirectUrl(tpl, valuesForUrl, {
          itemId: id,
          formMode,
          dynamicContext,
        });
        window.location.assign(url);
        return 'redirect';
      }
      if (f.kind === 'clearForm') {
        setFormError(undefined);
        flushSync(() => {
          const empty = itemToFormValues(undefined, names);
          setValues(
            getDefaultValuesFromRules(formManager, empty, dynamicContext, {
              isDateTimeField: isDateTimeFieldFromMeta,
            })
          );
          setButtonOverlay({ show: new Set<string>(), hide: new Set<string>() });
          setPendingFiles([]);
          setPendingFilesByFolder({});
          setLocalErrors({});
          setLinkedPendingByKey({});
          setLinkedRowErrorsById({});
        });
        setStepIndex(0);
        if (formMode === 'create' && linkedConfigsSorted.length > 0) {
          setLinkedRowsById((prev) => {
            const next: Record<string, ILinkedChildRowState[]> = { ...prev };
            for (let i = 0; i < linkedConfigsSorted.length; i++) {
              const c = linkedConfigsSorted[i];
              const min = c.minRows ?? 0;
              const count = Math.max(min, 1);
              next[c.id] = Array.from({ length: count }, (_, j) => ({
                localKey: `new_${c.id}_${j}_${Date.now()}_${Math.random().toString(36).slice(2)}`,
                values: {},
              }));
            }
            return next;
          });
        }
        return 'cleared';
      }
      return 'none';
    },
    [names, formManager, dynamicContext, itemId, formMode, linkedConfigsSorted, isDateTimeFieldFromMeta]
  );

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
    const cr = await confirmBeforeRunIfNeeded(btn);
    if (!cr.proceed) return;
    const baselinePatch = cr.valuesBaselinePatch;
    const baseValues =
      baselinePatch && Object.keys(baselinePatch).length > 0 ? { ...values, ...baselinePatch } : values;
    if (baselinePatch && Object.keys(baselinePatch).length > 0) {
      flushSync(() => {
        setValues((prev) => ({ ...prev, ...baselinePatch }));
      });
    }
    const useRunTimeline =
      btn.confirmBeforeRun?.enabled === true &&
      ((btn.confirmBeforeRun?.message ?? '').trim().length > 0 ||
        (btn.confirmBeforeRun?.promptFieldInternalName ?? '').trim().length > 0);
    const logWillRun = actionLogWillRunForTimeline(formManager.actionLog);
    const hasPendingAttachments =
      flatPendingFiles.length > 0 ||
      Object.keys(pendingFilesByFolder).some((k) => (pendingFilesByFolder[k]?.length ?? 0) > 0);
    const hasLinkedChildren = linkedConfigsSorted.length > 0;
    const tlCtx: IFormButtonRunTimelineCtx = {
      hasLinkedChildren,
      actionLogWillRun: logWillRun,
      hasPendingAttachments,
      permissionBreakWillRun: formManager.permissionBreak?.enabled === true,
    };
    const runTlTitle = `Execução: ${(btn.label || btn.id).trim() || 'Botão'}`;
    let tl: IRunTimelineCtl | null = null;
    let ti = 0;
    const op: TFormCustomButtonOperation = btn.operation ?? 'legacy';

    if (op === 'history') {
      if (useRunTimeline) {
        tl = createRunTimelineController(
          setRunTimelineDialog,
          buildCustomButtonRunTimelineLabels(btn, tlCtx),
          runTlTitle
        );
      }
      if (formManager.historyEnabled !== true) {
        setFormError('Ative o histórico na aba Componentes do gestor de formulário.');
        if (tl) {
          tl.enter(0);
          tl.err(0);
          tl.closeError();
        }
        return;
      }
      if (itemId === undefined || itemId === null || formMode === 'create') {
        setFormError('O histórico só está disponível quando o item já existe na lista.');
        if (tl) {
          tl.enter(0);
          tl.err(0);
          tl.closeError();
        }
        return;
      }
      setFormError(undefined);
      if (tl) tl.enter(0);
      setHistoryBtn(btn);
      if (tl) tl.ok(0);
      tl?.closeSuccess();
      return;
    }

    if (useRunTimeline && op !== 'delete') {
      tl = createRunTimelineController(
        setRunTimelineDialog,
        buildCustomButtonRunTimelineLabels(btn, tlCtx),
        runTlTitle
      );
    }

    const actions = op === 'redirect' ? [] : btn.actions ?? [];
    if (actions.length > 0) {
      if (tl) tl.enter(ti);
    }
    const { mergedValues, mergedOverlay } = reduceCustomButtonActions(
      actions,
      baseValues,
      dynamicContext,
      buttonOverlay,
      attachmentFolderUrl
    );
    if (actions.length > 0 && tl) {
      tl.ok(ti);
      ti++;
    }
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
        if (tl) {
          tl.enter(ti);
          tl.err(ti);
          tl.closeError();
        }
        return;
      }
      const url = interpolateFormButtonRedirectUrl(tpl, mergedValues, {
        itemId,
        formMode,
        dynamicContext,
      });
      if (logWillRun) {
        if (tl) tl.enter(ti);
        try {
          await appendFormActionLogEntry(
            itemsService,
            formManager.actionLog,
            btn,
            buildActionLogRuntimeCtx(btn, itemId)
          );
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
          if (tl) {
            tl.err(ti);
            tl.closeError();
          }
          return;
        }
        if (tl) tl.ok(ti);
        ti++;
      }
      if (tl) tl.enter(ti);
      window.location.assign(url);
      if (tl) tl.ok(ti);
      tl?.closeSuccess();
      return;
    }

    if (op === 'add') {
      setFormError(undefined);
      if (tl) tl.enter(ti);
      const validationOutcome = await validate('submit', { values: mergedValues, buttonOverlay: mergedOverlay });
      if (validationOutcome) {
        setRequiredValidationModalSections(
          buildValidationModalSections({
            mainErrors: validationOutcome.errors,
            formManager,
            fieldConfigs,
            ctx: {
              formMode,
              values: mergedValues,
              submitKind: validationOutcome.submitKind,
              userGroupTitles,
              currentUserId,
              authorId,
              dynamicContext,
              attachmentFolderUrl,
            },
            buttonOverlay: { show: mergedOverlay.show, hide: mergedOverlay.hide },
            fieldLabelByName: fieldLabelByName,
            mainFieldMetaByName: metaByName,
            linkedConfigs: linkedConfigsSorted,
            linkedRowErrorsById: validationOutcome.linkedRowErrorsById,
            linkedRowsById,
            linkedMetaById,
            mainListLabel: listTitle,
          })
        );
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        return;
      }
      if (tl) tl.ok(ti);
      ti++;
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      try {
        if (tl) tl.enter(ti);
        const payload = formValuesToSharePointPayload(fieldMetadata, mergedValues, names, {
          nullWhenEmptyFieldNames: ocultosNullFieldNames,
        });
        const { id: newId, filesForAttachments } = await itemsService.addItem(
          listTitle,
          payload,
          multiFolderAttachmentMode ? flatPendingFiles : pendingFiles,
          listWeb
        );
        if (tl) tl.ok(ti);
        ti++;
        const idComputedPatch = buildPostCreateItemIdComputedPatch({
          cfg: formManager,
          fieldConfigs,
          values: mergedValues,
          dynamicContext,
          attachmentFolderUrl,
          userGroupTitles,
          submitKind: 'submit',
          newItemId: newId,
          fieldMetaByName: metaByName,
        });
        const idPatchFieldNames = Object.keys(idComputedPatch);
        if (idPatchFieldNames.length > 0) {
          if (tl) tl.enter(ti);
          try {
            const mergedPatchValues = { ...mergedValues, ...idComputedPatch };
            const updatePayload = formValuesToSharePointPayload(fieldMetadata, mergedPatchValues, idPatchFieldNames, {
              nullWhenEmptyFieldNames: ocultosNullFieldNames,
            });
            if (Object.keys(updatePayload).length > 0) {
              await itemsService.updateItem(listTitle, newId, updatePayload, listWeb);
            }
          } catch (ue) {
            setFormError(
              `Item criado (ID ${newId}), mas a atualização dos campos calculados com ID falhou: ${
                ue instanceof Error ? ue.message : String(ue)
              }`
            );
            setSubmitUi(null);
            if (tl) {
              tl.err(ti);
              tl.closeError();
            }
            return;
          }
          if (tl) tl.ok(ti);
          ti++;
        }
        if (hasPendingAttachments) {
          if (tl) tl.enter(ti);
          await uploadListItemAttachments(
            listTitle,
            newId,
            multiFolderAttachmentMode ? [] : filesForAttachments,
            formManager,
            {
              ...mergedValues,
              Id: newId,
            },
            multiFolderAttachmentMode ? pendingFilesByFolder : undefined,
            tl
              ? (info): void => {
                  tl!.setRunningDetail(`Pasta: ${info.folderLabel} · ${info.fileName}`);
                }
              : undefined,
            listWeb
          );
          if (tl) tl.ok(ti);
          ti++;
        }
        let linkedSnapAdd = linkedRowsById;
        if (hasLinkedChildren) {
          if (tl) tl.enter(ti);
          try {
            const relAdd = await performLinkedSync(newId);
            if (relAdd) linkedSnapAdd = relAdd;
          } catch (le) {
            setFormError(
              `Item criado, mas as listas vinculadas falharam: ${le instanceof Error ? le.message : String(le)}`
            );
            setSubmitUi(null);
            if (tl) {
              tl.err(ti);
              tl.closeError();
            }
            return;
          }
          if (tl) tl.ok(ti);
          ti++;
        }
        if (formManager.permissionBreak?.enabled) {
          if (tl) tl.enter(ti);
          try {
            await runPermissionBreakAfterSubmit(
              newId,
              mergedValues,
              linkedSnapAdd,
              tl ? (d) => tl!.setRunningDetail(d) : undefined
            );
            if (tl) {
              tl.setRunningDetail(undefined);
              tl.ok(ti);
              ti++;
            }
          } catch (pe) {
            setFormError(
              `Item criado, mas a quebra de permissões falhou: ${pe instanceof Error ? pe.message : String(pe)}`
            );
            if (tl) {
              tl.setRunningDetail(undefined);
              tl.err(ti);
              tl.closeError();
              setSubmitUi(null);
              return;
            }
          }
        }
        if (logWillRun) {
          if (tl) tl.enter(ti);
          try {
            await appendFormActionLogEntry(
              itemsService,
              formManager.actionLog,
              btn,
              buildActionLogRuntimeCtx(btn, newId)
            );
          } catch (le) {
            setFormError(
              `Item criado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
            );
          }
          if (tl) tl.ok(ti);
          ti++;
        }
        if (tl) tl.enter(ti);
        const fin = await runFinishAfterSuccess(btn, { ...mergedValues, Id: newId }, newId);
        if (tl) {
          if (fin === 'none' && btn.finishAfterRun?.kind === 'redirect') {
            tl.err(ti);
            tl.closeError();
          } else {
            tl.ok(ti);
            tl.closeSuccess();
          }
        }
        if (fin !== 'redirect' && fin !== 'cleared') {
          onDismiss();
        }
      } catch (e) {
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      return;
    }

    if (op === 'update') {
      if (!itemId || formMode === 'create') {
        setFormError('Atualizar requer um item aberto (parâmetros Form / FormID na página).');
        if (tl) {
          tl.enter(ti);
          tl.err(ti);
          tl.closeError();
        }
        return;
      }
      setFormError(undefined);
      if (tl) tl.enter(ti);
      const validationOutcome = await validate('submit', { values: mergedValues, buttonOverlay: mergedOverlay });
      if (validationOutcome) {
        setRequiredValidationModalSections(
          buildValidationModalSections({
            mainErrors: validationOutcome.errors,
            formManager,
            fieldConfigs,
            ctx: {
              formMode,
              values: mergedValues,
              submitKind: validationOutcome.submitKind,
              userGroupTitles,
              currentUserId,
              authorId,
              dynamicContext,
              attachmentFolderUrl,
            },
            buttonOverlay: { show: mergedOverlay.show, hide: mergedOverlay.hide },
            fieldLabelByName: fieldLabelByName,
            mainFieldMetaByName: metaByName,
            linkedConfigs: linkedConfigsSorted,
            linkedRowErrorsById: validationOutcome.linkedRowErrorsById,
            linkedRowsById,
            linkedMetaById,
            mainListLabel: listTitle,
          })
        );
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        return;
      }
      if (tl) tl.ok(ti);
      ti++;
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      let updateBusinessOk = false;
      try {
        if (tl) tl.enter(ti);
        const payload = formValuesToSharePointPayload(fieldMetadata, mergedValues, names, {
          nullWhenEmptyFieldNames: ocultosNullFieldNames,
        });
        await itemsService.updateItem(listTitle, itemId, payload, listWeb);
        if (tl) tl.ok(ti);
        ti++;
        if (hasPendingAttachments) {
          if (tl) tl.enter(ti);
          await uploadListItemAttachments(
            listTitle,
            itemId,
            multiFolderAttachmentMode ? [] : pendingFiles,
            formManager,
            {
              ...mergedValues,
              Id: itemId,
            },
            multiFolderAttachmentMode ? pendingFilesByFolder : undefined,
            tl
              ? (info): void => {
                  tl!.setRunningDetail(`Pasta: ${info.folderLabel} · ${info.fileName}`);
                }
              : undefined,
            listWeb
          );
          if (tl) tl.ok(ti);
          ti++;
        }
        let linkedSnapUp = linkedRowsById;
        if (hasLinkedChildren) {
          if (tl) tl.enter(ti);
          try {
            const relUp = await performLinkedSync(itemId);
            if (relUp) linkedSnapUp = relUp;
          } catch (le) {
            setFormError(
              `Gravado, mas as listas vinculadas falharam: ${le instanceof Error ? le.message : String(le)}`
            );
            setSubmitUi(null);
            if (tl) {
              tl.err(ti);
              tl.closeError();
            }
            return;
          }
          if (tl) tl.ok(ti);
          ti++;
        }
        if (formManager.permissionBreak?.enabled) {
          if (tl) tl.enter(ti);
          try {
            await runPermissionBreakAfterSubmit(
              itemId,
              mergedValues,
              linkedSnapUp,
              tl ? (d) => tl!.setRunningDetail(d) : undefined
            );
            if (tl) {
              tl.setRunningDetail(undefined);
              tl.ok(ti);
              ti++;
            }
          } catch (pe) {
            setFormError(
              `Gravado, mas a quebra de permissões falhou: ${pe instanceof Error ? pe.message : String(pe)}`
            );
            if (tl) {
              tl.setRunningDetail(undefined);
              tl.err(ti);
              tl.closeError();
              setSubmitUi(null);
              return;
            }
          }
        }
        await onAfterItemUpdated?.();
        if (logWillRun) {
          if (tl) tl.enter(ti);
          try {
            await appendFormActionLogEntry(
              itemsService,
              formManager.actionLog,
              btn,
              buildActionLogRuntimeCtx(btn, itemId)
            );
          } catch (le) {
            setFormError(
              `Gravado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
            );
          }
          if (tl) tl.ok(ti);
          ti++;
        }
        updateBusinessOk = true;
      } catch (e) {
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      if (updateBusinessOk) {
        if (tl) tl.enter(ti);
        const finUp = await runFinishAfterSuccess(btn, { ...mergedValues, Id: itemId }, itemId);
        if (tl) {
          if (finUp === 'none' && btn.finishAfterRun?.kind === 'redirect') {
            tl.err(ti);
            tl.closeError();
          } else {
            tl.ok(ti);
            tl.closeSuccess();
          }
        }
      }
      return;
    }

    if (op === 'delete') {
      if (!itemId || formMode === 'create') {
        setFormError('Eliminar só está disponível ao editar ou ver um item existente.');
        return;
      }
      const skipNativeDeleteConfirm =
        btn.confirmBeforeRun?.enabled === true &&
        ((btn.confirmBeforeRun?.message ?? '').trim().length > 0 ||
          (btn.confirmBeforeRun?.promptFieldInternalName ?? '').trim().length > 0);
      if (!skipNativeDeleteConfirm && !window.confirm('Eliminar este item permanentemente?')) return;
      if (useRunTimeline) {
        tl = createRunTimelineController(
          setRunTimelineDialog,
          buildCustomButtonRunTimelineLabels(btn, tlCtx),
          runTlTitle
        );
      }
      if (actions.length > 0 && tl) {
        tl.enter(0);
        tl.ok(0);
        ti = 1;
      }
      setFormError(undefined);
      setSubmitUi(resolveSubmitLoadingKind(formManager, btn));
      try {
        if (tl) tl.enter(ti);
        await itemsService.deleteItem(listTitle, itemId, listWeb);
        if (tl) tl.ok(ti);
        ti++;
        if (logWillRun) {
          if (tl) tl.enter(ti);
          try {
            await appendFormActionLogEntry(
              itemsService,
              formManager.actionLog,
              btn,
              buildActionLogRuntimeCtx(btn, itemId)
            );
          } catch (le) {
            setFormError(
              `Eliminado, mas o registo de log falhou: ${le instanceof Error ? le.message : String(le)}`
            );
          }
          if (tl) tl.ok(ti);
          ti++;
        }
        if (tl) tl.enter(ti);
        const finDel = await runFinishAfterSuccess(btn, mergedValues, itemId);
        if (tl) {
          if (finDel === 'none' && btn.finishAfterRun?.kind === 'redirect') {
            tl.err(ti);
            tl.closeError();
          } else {
            tl.ok(ti);
            tl.closeSuccess();
          }
        }
        if (finDel !== 'redirect') {
          onDismiss();
        }
      } catch (e) {
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        setFormError(e instanceof Error ? e.message : String(e));
      } finally {
        setSubmitUi(null);
      }
      return;
    }

    const behavior = btn.behavior ?? 'actionsOnly';
    if (behavior === 'close') {
      if (logWillRun) {
        if (tl) tl.enter(ti);
        try {
          await appendFormActionLogEntry(
            itemsService,
            formManager.actionLog,
            btn,
            buildActionLogRuntimeCtx(btn, itemId)
          );
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
          if (tl) {
            tl.err(ti);
            tl.closeError();
          }
          return;
        }
        if (tl) tl.ok(ti);
        ti++;
      }
      if (tl) tl.enter(ti);
      const finClose = await runFinishAfterSuccess(btn, mergedValues, itemId);
      if (tl) {
        if (finClose === 'none' && btn.finishAfterRun?.kind === 'redirect') {
          tl.err(ti);
          tl.closeError();
        } else {
          tl.ok(ti);
          tl.closeSuccess();
        }
      }
      if (finClose !== 'redirect') {
        onDismiss();
      }
      return;
    }
    if (behavior === 'draft') {
      if (tl) tl.enter(ti);
      const saved = await handleSave('draft', {
        valuesOverride: mergedValues,
        buttonOverlayOverride: mergedOverlay,
        submitLoadingFromButton: btn,
      });
      if (!saved) {
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        return;
      }
      if (tl) tl.ok(ti);
      ti++;
      if (logWillRun) {
        if (tl) tl.enter(ti);
        try {
          await appendFormActionLogEntry(
            itemsService,
            formManager.actionLog,
            btn,
            buildActionLogRuntimeCtx(btn, itemId)
          );
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
          if (tl) {
            tl.err(ti);
            tl.closeError();
          }
          return;
        }
        if (tl) tl.ok(ti);
        ti++;
      }
      if (tl) tl.enter(ti);
      try {
        const finDr = await runFinishAfterSuccess(btn, mergedValues, itemId);
        if (tl) {
          if (finDr === 'none' && btn.finishAfterRun?.kind === 'redirect') {
            tl.err(ti);
            tl.closeError();
          } else {
            tl.ok(ti);
            tl.closeSuccess();
          }
        }
      } catch (e) {
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        setFormError(e instanceof Error ? e.message : String(e));
      }
    } else if (behavior === 'submit') {
      if (tl) tl.enter(ti);
      const saved = await handleSave('submit', {
        valuesOverride: mergedValues,
        buttonOverlayOverride: mergedOverlay,
        submitLoadingFromButton: btn,
      });
      if (!saved) {
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        return;
      }
      if (tl) tl.ok(ti);
      ti++;
      if (logWillRun) {
        if (tl) tl.enter(ti);
        try {
          await appendFormActionLogEntry(
            itemsService,
            formManager.actionLog,
            btn,
            buildActionLogRuntimeCtx(btn, itemId)
          );
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
          if (tl) {
            tl.err(ti);
            tl.closeError();
          }
          return;
        }
        if (tl) tl.ok(ti);
        ti++;
      }
      if (tl) tl.enter(ti);
      try {
        const finSb = await runFinishAfterSuccess(btn, mergedValues, itemId);
        if (tl) {
          if (finSb === 'none' && btn.finishAfterRun?.kind === 'redirect') {
            tl.err(ti);
            tl.closeError();
          } else {
            tl.ok(ti);
            tl.closeSuccess();
          }
        }
      } catch (e) {
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
        setFormError(e instanceof Error ? e.message : String(e));
      }
    } else if (behavior === 'actionsOnly') {
      if (logWillRun) {
        if (tl) tl.enter(ti);
        try {
          await appendFormActionLogEntry(
            itemsService,
            formManager.actionLog,
            btn,
            buildActionLogRuntimeCtx(btn, itemId)
          );
        } catch (e) {
          setFormError(e instanceof Error ? e.message : String(e));
          if (tl) {
            tl.err(ti);
            tl.closeError();
          }
          return;
        }
        if (tl) tl.ok(ti);
        ti++;
      }
      if (tl) tl.enter(ti);
      try {
        const finAo = await runFinishAfterSuccess(btn, mergedValues, itemId);
        if (tl) {
          if (finAo === 'none' && btn.finishAfterRun?.kind === 'redirect') {
            tl.err(ti);
            tl.closeError();
          } else {
            tl.ok(ti);
            tl.closeSuccess();
          }
        }
      } catch (e) {
        if (tl) {
          tl.err(ti);
          tl.closeError();
        }
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
        const sync = collectFormValidationErrors(formManager, fieldConfigs, ctx, attCtx, overlay, metaByName);
        let rel = filterValidationErrorsToStepFields(sync, stepFieldList);
        if (!fullVal) rel = pickRequiredStyleStepErrors(rel);
        let merged: Record<string, string> = { ...rel };
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
  const [attachmentDetailRow, setAttachmentDetailRow] = useState<IServerAttachmentRow | null>(null);
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
              onClick={() => setAttachmentDetailRow(a)}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  setAttachmentDetailRow(a);
                }
              }}
            >
              {a.fileName}
            </Text>
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
    const setComputedRule = findEnabledSetComputedRule(formManager.rules, name);
    const comp = resolveSetComputedDisplayValue({
      derivedComputed: derived.computedDisplay[name],
      formMode,
      itemId,
      fieldName: name,
      expressionSnapAtItemOpenByField: setComputedExprSnapRef.current.snap,
      setComputedRule,
    });
    if (comp !== undefined) {
      const mComp = metaByName.get(name);
      const labelComp = fc.label ?? mComp?.Title ?? name;
      const helpComp = derived.dynamicHelpByField[name] ?? fc.helpText;
      const reqComp = derived.fieldRequired[name] === true || mComp?.Required === true;
      const compShown =
        mComp?.MappedType === 'datetime'
          ? ((): string => {
              const s = String(comp);
              const ms = Date.parse(s);
              return !isNaN(ms) ? new Date(ms).toLocaleDateString('pt-BR') : s;
            })()
          : String(comp);
      return (
        <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 12 } }}>
          <Label required={reqComp}>{labelComp}</Label>
          <Text styles={{ root: { color: '#323130' } }}>{compShown}</Text>
          {helpComp && <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{helpComp}</Text>}
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

    const common = { disabled: readOnly, errorMessage: err };

    const renderLookupDetailsBelow = (fieldName: string, meta: IFieldMetadata): React.ReactNode => {
      const dfn = fc.lookupOptionDetailBelowFields ?? [];
      if (!dfn.length) return null;
      const snap = lookupDetailSnapshot[fieldName];
      if (snap === undefined) return null;
      const listGuid = String(meta.LookupList ?? '');
      const linkedMeta = lookupDestMetaCacheRef.current[listGuid] ?? [];
      const rowsArr = Array.isArray(snap) ? snap : [snap];
      if (rowsArr.length === 0) return null;
      return (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 2, width: '100%' } }}>
          {dfn.map((fn) => {
            const fm = linkedMeta.find((x) => x.InternalName === fn);
            const fieldLabel = fm ? fm.Title : fn;
            const parts = rowsArr.map((row) =>
              lookupRowToOptionText(row as Record<string, unknown>, fn, fm)
            );
            const val = parts.join('; ');
            return (
              <TextField
                key={fn}
                label={fieldLabel}
                value={val}
                readOnly
                disabled
                multiline={val.length > 100}
                rows={val.length > 100 ? 2 : 1}
              />
            );
          })}
        </Stack>
      );
    };

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
      case 'currency': {
        const numBounds = validateValueNumberMergedByField[name];
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
            min={numBounds?.minNumber}
            max={numBounds?.maxNumber}
          />
        );
      }
      case 'datetime':
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 8 } }}>
            <Label required={isRequired}>{label}</Label>
            <DatePicker
              {...FLUENT_DATE_PICKER_PT_BR}
              value={values[name] ? new Date(String(values[name])) : undefined}
              onSelectDate={(d) => applyDateFieldSelect(name, d ?? null)}
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
            onRenderTitle={(opts) => renderMultiSelectDropdownTitle(theme, opts)}
            styles={multiSelectDropdownStyles(showReqEmpty)}
          />
        );
      }
      case 'lookup': {
        const lfDrop = derived.lookupFilters[name];
        const filterActive = hasConfiguredLookupFilter(lfDrop);
        const parentMetaDrop = lfDrop ? metaByName.get(lfDrop.parentField) : undefined;
        const parentReady =
          !filterActive ||
          isParentValueReadyForLookupFilter(
            lfDrop ? values[lfDrop.parentField] : undefined,
            parentMetaDrop
          );
        const lookupBlockedByParent = filterActive && !parentReady;
        const id = lookupIdFromValue(values[name]);
        const baseOpts = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const opts =
          id !== undefined && id > 0
            ? mergeOptionsForIds(baseOpts, [{ id, label: userTitleFromValue(values[name]) }])
            : baseOpts;
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 12 } }}>
            <Dropdown
              label={label}
              placeholder={fc.placeholder}
              options={opts}
              selectedKey={id !== undefined ? String(id) : ''}
              onChange={(_, o) => {
                if (!o || o.key === '') updateField(name, null);
                else {
                  const raw =
                    o && typeof o === 'object' && 'data' in o ? (o as { data?: Record<string, unknown> }).data : undefined;
                  if (raw && typeof raw === 'object' && typeof raw.Id === 'number') {
                    updateField(name, raw);
                  } else {
                    updateField(name, { Id: Number(o.key), Title: String(o.text ?? '') });
                  }
                }
              }}
              required={isRequired}
              errorMessage={err}
              disabled={readOnly || lookupBlockedByParent}
              styles={dropdownReqStyles(showReqEmpty)}
            />
            {help ? (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {help}
              </Text>
            ) : null}
            {renderLookupDetailsBelow(name, m)}
          </Stack>
        );
      }
      case 'lookupmulti': {
        const lfMulti = derived.lookupFilters[name];
        const filterActiveMulti = hasConfiguredLookupFilter(lfMulti);
        const parentMetaMulti = lfMulti ? metaByName.get(lfMulti.parentField) : undefined;
        const parentReadyMulti =
          !filterActiveMulti ||
          isParentValueReadyForLookupFilter(
            lfMulti ? values[lfMulti.parentField] : undefined,
            parentMetaMulti
          );
        const lookupBlockedByParentMulti = filterActiveMulti && !parentReadyMulti;
        const selected = normalizeIdTitleArray(values[name]);
        const baseRaw = lookupOptions[name] ?? [{ key: '', text: '—' }];
        const baseOpts = baseRaw.filter((o) => o.key !== '');
        const extra = selected.map((x) => ({ id: x.Id, label: x.Title }));
        const opts = mergeOptionsForIds(baseOpts, extra);
        const keys = selected.map((x) => String(x.Id));
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 12 } }}>
            <Dropdown
              label={label}
              placeholder={fc.placeholder}
              multiSelect
              options={opts}
              selectedKeys={keys}
              onChange={(_, o) => {
                if (!o || o.key === '') return;
                const k = String(o.key);
                const hit = selected.findIndex((x) => String(x.Id) === k);
                const raw =
                  o && typeof o === 'object' && 'data' in o ? (o as { data?: Record<string, unknown> }).data : undefined;
                const nextItem =
                  raw && typeof raw === 'object' && typeof raw.Id === 'number'
                    ? raw
                    : { Id: Number(o.key), Title: String(o.text ?? '') };
                const next =
                  hit === -1
                    ? [...selected, nextItem]
                    : selected.filter((_, i) => i !== hit);
                updateField(name, next);
              }}
              required={isRequired}
              errorMessage={err}
              disabled={readOnly || lookupBlockedByParentMulti}
              onRenderTitle={(opts) => renderMultiSelectDropdownTitle(theme, opts)}
              styles={multiSelectDropdownStyles(showReqEmpty)}
            />
            {help ? (
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {help}
              </Text>
            ) : null}
            {renderLookupDetailsBelow(name, m)}
          </Stack>
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
            onRenderTitle={(opts) => renderMultiSelectDropdownTitle(theme, opts)}
            styles={multiSelectDropdownStyles(showReqEmpty)}
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
      case 'multiline': {
        const raw =
          values[name] !== null && values[name] !== undefined ? String(values[name]) : '';
        if (readOnly && shouldRenderMultilineNoteAsHtml(m, raw)) {
          return (
            <MultilineReadonlyHtml
              key={name}
              label={label}
              required={isRequired}
              html={raw}
              help={help}
              showReqEmpty={showReqEmpty}
            />
          );
        }
        const mergedMl = validateValueLengthMergedByField[name];
        const maxCapMl = mergedMl?.maxLength;
        const charBoundsMl =
          mergedMl &&
          mergedMl.minLength !== undefined &&
          mergedMl.maxLength !== undefined &&
          !readOnly &&
          comp === undefined
            ? { minLength: mergedMl.minLength, maxLength: mergedMl.maxLength }
            : null;
        const charHintMl = charBoundsMl
          ? validateValueCharLimitHint(raw.length, charBoundsMl.minLength, charBoundsMl.maxLength)
          : null;
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 12 } }}>
            <TextField
              label={label}
              multiline
              rows={resolveTextareaRows(fc, 4)}
              placeholder={fc.placeholder}
              value={raw}
              onChange={(_, v) => {
                let s = v ?? '';
                if (maxCapMl !== undefined && s.length > maxCapMl) s = s.slice(0, maxCapMl);
                updateField(name, s);
              }}
              required={isRequired}
              {...common}
              {...(maxCapMl !== undefined ? { maxLength: maxCapMl } : {})}
              description={help}
              styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
            />
            {charHintMl ? (
              <Text variant="small" styles={{ root: { color: charHintMl.color } }}>{charHintMl.text}</Text>
            ) : null}
          </Stack>
        );
      }
      default: {
        const rawSingle =
          values[name] !== null && values[name] !== undefined ? String(values[name]) : '';
        const mergedTx = validateValueLengthMergedByField[name];
        const maxCapTx = mergedTx?.maxLength;
        const charBoundsTx =
          mergedTx &&
          mergedTx.minLength !== undefined &&
          mergedTx.maxLength !== undefined &&
          !readOnly &&
          comp === undefined
            ? { minLength: mergedTx.minLength, maxLength: mergedTx.maxLength }
            : null;
        const charHintTx = charBoundsTx
          ? validateValueCharLimitHint(rawSingle.length, charBoundsTx.minLength, charBoundsTx.maxLength)
          : null;
        const maskOpts =
          m.MappedType === 'text'
            ? resolveTextInputMaskOptions(fc.textInputMaskKind, fc.textInputMaskCustomPattern)
            : null;
        const inputBorder = err ? theme.semanticColors.errorText : theme.semanticColors.inputBorder;
        const maskInputStyle: React.CSSProperties = {
          width: '100%',
          boxSizing: 'border-box',
          minHeight: 32,
          padding: '0 8px',
          font: 'inherit',
          color: theme.semanticColors.inputText,
          backgroundColor: theme.semanticColors.inputBackground,
          borderWidth: 1,
          borderStyle: 'solid',
          borderColor: inputBorder,
          borderRadius: 2,
          outline: 'none',
        };
        if (maskOpts) {
          return (
            <Stack
              key={name}
              tokens={{ childrenGap: 4 }}
              styles={{
                root: {
                  marginBottom: 12,
                  ...(showReqEmpty
                    ? {
                        borderLeft: `3px solid ${REQ_EMPTY_BORDER}`,
                        paddingLeft: 8,
                        paddingTop: 2,
                        paddingBottom: 2,
                      }
                    : {}),
                },
              }}
            >
              <Label required={isRequired}>{label}</Label>
              <IMaskInput
                {...maskOpts}
                value={rawSingle}
                disabled={readOnly}
                placeholder={fc.placeholder ?? undefined}
                onAccept={(val) => {
                  let s = String(val ?? '');
                  if (maxCapTx !== undefined && s.length > maxCapTx) s = s.slice(0, maxCapTx);
                  updateField(name, s);
                }}
                style={maskInputStyle}
                aria-invalid={err ? true : undefined}
                aria-required={isRequired ? true : undefined}
              />
              {err ? (
                <Text variant="small" styles={{ root: { color: theme.semanticColors.errorText } }}>{err}</Text>
              ) : null}
              {help ? <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{help}</Text> : null}
              {charHintTx ? (
                <Text variant="small" styles={{ root: { color: charHintTx.color } }}>{charHintTx.text}</Text>
              ) : null}
            </Stack>
          );
        }
        return (
          <Stack key={name} tokens={{ childrenGap: 4 }} styles={{ root: { marginBottom: 12 } }}>
            <TextField
              label={label}
              placeholder={fc.placeholder}
              value={rawSingle}
              onChange={(_, v) => {
                let s = v ?? '';
                if (maxCapTx !== undefined && s.length > maxCapTx) s = s.slice(0, maxCapTx);
                updateField(name, s);
              }}
              required={isRequired}
              {...common}
              {...(maxCapTx !== undefined ? { maxLength: maxCapTx } : {})}
              description={help}
              styles={stylesTextFieldRequiredEmpty(showReqEmpty)}
            />
            {charHintTx ? (
              <Text variant="small" styles={{ root: { color: charHintTx.color } }}>{charHintTx.text}</Text>
            ) : null}
          </Stack>
        );
      }
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
    const hideStepHeading = visibleStepsForUi != null && visibleStepsForUi.length === 1;
    for (let s = 0; s < formManager.sections.length; s++) {
      const sec = formManager.sections[s];
      if (sec.id === FORM_OCULTOS_STEP_ID || sec.id === FORM_FIXOS_STEP_ID) continue;
      if (derived.sectionVisible[sec.id] === false) continue;
      const fields = bySection.get(sec.id);
      if (!fields?.length) continue;
      out.push(
        <Stack key={sec.id} tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 16 } }}>
          {!hideStepHeading ? (
            <Text variant="large" styles={{ root: { fontWeight: 600 } }}>
              {sec.title}
            </Text>
          ) : null}
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

  function renderOneCustomButton(b: IFormCustomButtonConfig): React.ReactElement {
    const slot = resolveFormCustomButtonPaletteSlot(b);
    const common = {
      text: b.label,
      title: b.shortDescription || undefined,
      onClick: () => void runCustomButton(b),
      disabled: submitting,
    };
    if (slot === 'outline') {
      return <DefaultButton key={b.id} {...common} />;
    }
    return <DefaultButton key={b.id} {...common} styles={getFilledPaletteButtonStyles(theme, slot)} />;
  }

  const submitMsg = 'A gravar…';

  const customButtonsBarVertical = formManager.customButtonsBarVertical ?? 'bottom';
  const customButtonsBarHorizontal = formManager.customButtonsBarHorizontal ?? 'left';

  function renderCustomButtonsToolbar(): React.ReactNode {
    const showHistory =
      formManager.historyEnabled === true &&
      shouldShowBuiltinHistoryButton({
        historyEnabledInConfig: true,
        historyItemId: itemId,
        historyGroupTitles: formManager.historyGroupTitles,
        userGroupTitles,
      });
    if (!visibleCustomButtons.length && !showHistory) return null;
    const hAlign = customButtonsBarHorizontal === 'left' ? 'start' : 'end';
    return (
      <Stack
        horizontal
        horizontalAlign={hAlign}
        tokens={{ childrenGap: 8 }}
        wrap
        styles={{ root: { width: '100%' } }}
      >
        {visibleCustomButtons.map((b) => renderOneCustomButton(b))}
        {showHistory &&
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
    );
  }

  const mainFormMiddle = (
    <>
      {visibleStepsForUi && visibleStepsForUi.length > 1 && (
        <FormStepNavigation
          steps={visibleStepsForUi}
          activeIndex={stepIndex}
          onStepSelect={(i) => tryGoToStep(i)}
          layout={formManager.stepLayout ?? 'segmented'}
          accentColor={stepAccentHex}
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
          formManager={formManager}
          linkedPendingFilesByKey={linkedPendingByKey}
          onLinkedPendingFilesChange={(key, files) =>
            setLinkedPendingByKey((prev) => ({ ...prev, [key]: files }))
          }
          currentParentStepId={visibleStepsForUi?.[stepIndex]?.id ?? 'main'}
          attachmentUploadLayout={formManager.attachmentUploadLayout}
          attachmentFilePreview={attachmentPreviewKind}
          attachmentAllowedExtensions={
            attachmentAllowedExtensions.length > 0 ? attachmentAllowedExtensions : undefined
          }
          linkedServerAttachmentsByKey={linkedServerAttachmentsByKey}
        />
      )}
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
            accentColor={stepAccentHex}
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
    </>
  );

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
        {customButtonsBarVertical === 'top' && renderCustomButtonsToolbar()}
        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { width: '100%' } }}>
          {mainFormMiddle}
        </Stack>
        {customButtonsBarVertical === 'bottom' && renderCustomButtonsToolbar()}
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
              accentColor={stepAccentHex}
              logEntryPaletteContext={historyLogEntryPaletteContext}
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
        <AttachmentFileDetailModal
          isOpen={attachmentDetailRow !== null}
          onDismiss={() => setAttachmentDetailRow(null)}
          target={
            attachmentDetailRow
              ? {
                  kind: 'server',
                  fileName: attachmentDetailRow.fileName,
                  fileUrl: attachmentDetailRow.fileUrl,
                  fileRef: attachmentDetailRow.fileRef,
                }
              : null
          }
        />
        <Dialog
          hidden={!confirmDialogOpen}
          onDismiss={() => closeButtonConfirmDialog(false)}
          dialogContentProps={{
            type: DialogType.normal,
            title: confirmDialogView?.title ?? 'Confirmar',
            showCloseButton: false,
            styles: {
              title: {
                position: 'absolute',
                width: 1,
                height: 1,
                margin: -1,
                padding: 0,
                overflow: 'hidden',
                clip: 'rect(0,0,0,0)',
                whiteSpace: 'nowrap',
                border: 0,
              },
              inner: { padding: 0 },
              subText: { padding: 0 },
            },
          }}
          modalProps={{
            isBlocking: true,
            styles: {
              root: {
                backgroundColor: 'rgba(15, 23, 42, 0.45)',
                backdropFilter: 'blur(6px)',
              },
            },
          }}
          styles={{
            root: { selectors: { '& .ms-Modal-scrollableContent': { overflow: 'hidden' } } },
            main: {
              maxWidth: 560,
              borderRadius: 16,
              overflow: 'hidden',
              boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.22)',
              border: `1px solid ${theme.palette.neutralLight}`,
            },
          }}
        >
          {confirmDialogView ? (
            (() => {
              const pal = confirmKindToCenteredModalPalette(confirmDialogView.kind);
              const confirmBtnStyles =
                pal.confirmDanger === true
                  ? {
                      flex: '1 1 0',
                      minWidth: 0,
                      borderRadius: 10,
                      height: 44,
                      backgroundColor: '#DC2626',
                      borderColor: '#DC2626',
                      borderWidth: 1,
                    }
                  : {
                      flex: '1 1 0',
                      minWidth: 0,
                      borderRadius: 10,
                      height: 44,
                    };
              const confirmBtnHovered =
                pal.confirmDanger === true
                  ? { backgroundColor: '#B91C1C', borderColor: '#B91C1C' }
                  : undefined;
              const cancelBtnStyles = {
                flex: '1 1 0',
                minWidth: 0,
                borderRadius: 10,
                height: 44,
                border: `1px solid ${theme.palette.neutralQuaternaryAlt}`,
                backgroundColor: theme.palette.white,
              };
              return (
                <Stack
                  horizontalAlign="center"
                  tokens={{ childrenGap: 16 }}
                  styles={{
                    root: {
                      padding: '28px 28px 24px',
                      boxSizing: 'border-box',
                    },
                  }}
                >
                  <div
                    style={{
                      width: 72,
                      height: 72,
                      borderRadius: '50%',
                      background: pal.circleBg,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      flexShrink: 0,
                      marginTop: 4,
                    }}
                  >
                    <Icon
                      iconName={pal.iconName}
                      styles={{ root: { fontSize: 34, lineHeight: 1, color: pal.iconColor } }}
                    />
                  </div>
                  <Text
                    as="h2"
                    variant="xLarge"
                    styles={{
                      root: {
                        fontWeight: 700,
                        textAlign: 'center',
                        color: theme.palette.neutralPrimary,
                        margin: 0,
                        padding: '0 4px',
                        lineHeight: 1.25,
                      },
                    }}
                  >
                    {confirmDialogView.title}
                  </Text>
                  {confirmDialogView.message.trim() ? (
                    <Text
                      variant="medium"
                      styles={{
                        root: {
                          textAlign: 'center',
                          color: theme.palette.neutralSecondary,
                          lineHeight: 1.55,
                          whiteSpace: 'pre-wrap',
                          maxWidth: 520,
                          margin: '0 auto',
                        },
                      }}
                    >
                      {confirmDialogView.message}
                    </Text>
                  ) : !(confirmPromptMetaForDialog && confirmPromptEditor) ? (
                    <Text
                      variant="medium"
                      styles={{
                        root: {
                          textAlign: 'center',
                          color: theme.palette.neutralSecondary,
                          lineHeight: 1.5,
                          maxWidth: 520,
                        },
                      }}
                    >
                      Tem a certeza que deseja continuar?
                    </Text>
                  ) : null}
                  {confirmPromptMetaForDialog && confirmPromptEditor ? (
                    <Stack
                      tokens={{ childrenGap: 10 }}
                      styles={{
                        root: {
                          width: '100%',
                          maxWidth: 520,
                          marginTop: 4,
                          textAlign: 'left',
                        },
                      }}
                    >
                      <ConfirmPromptFieldEditor
                        modalSurface
                        meta={confirmPromptMetaForDialog}
                        editor={confirmPromptEditor}
                        onChange={setConfirmPromptEditor}
                      />
                      {confirmPrimaryDisabled ? (
                        <Text
                          variant="small"
                          styles={{
                            root: { color: theme.palette.redDark, textAlign: 'center' },
                          }}
                        >
                          Preencha o campo acima para continuar.
                        </Text>
                      ) : null}
                    </Stack>
                  ) : null}
                  <Stack
                    horizontal
                    tokens={{ childrenGap: 12 }}
                    styles={{
                      root: {
                        width: '100%',
                        maxWidth: 520,
                        marginTop: 8,
                        justifyContent: 'stretch',
                      },
                    }}
                  >
                    <DefaultButton
                      text="Cancelar"
                      onClick={() => closeButtonConfirmDialog(false)}
                      styles={{
                        root: cancelBtnStyles,
                        flexContainer: { height: 44 },
                        label: { fontWeight: 600 },
                      }}
                    />
                    <PrimaryButton
                      text="Confirmar"
                      onClick={() => closeButtonConfirmDialog(true)}
                      disabled={confirmPrimaryDisabled}
                      styles={{
                        root: confirmBtnStyles,
                        rootHovered: confirmBtnHovered,
                        rootPressed:
                          pal.confirmDanger === true
                            ? { backgroundColor: '#991B1B', borderColor: '#991B1B' }
                            : undefined,
                        flexContainer: { height: 44 },
                        label: { fontWeight: 600 },
                      }}
                    />
                  </Stack>
                </Stack>
              );
            })()
          ) : null}
        </Dialog>
        <Dialog
          hidden={!requiredValidationModalSections?.length}
          onDismiss={() => setRequiredValidationModalSections(null)}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Campos por preencher',
            showCloseButton: true,
          }}
          modalProps={{ isBlocking: true }}
        >
          {requiredValidationModalSections && requiredValidationModalSections.length > 0 ? (
            <>
              <Stack tokens={{ childrenGap: 16 }} styles={{ root: { maxHeight: '60vh', overflowY: 'auto' } }}>
                {requiredValidationModalSections.map((sec, si) => (
                  <Stack key={`${si}_${sec.heading}`} tokens={{ childrenGap: 6 }}>
                    <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
                      {sec.heading}
                    </Text>
                    {sec.lines.map((ln, li) => (
                      <Text key={li} styles={{ root: { paddingLeft: 8 } }}>
                        • {ln}
                      </Text>
                    ))}
                  </Stack>
                ))}
              </Stack>
              <DialogFooter>
                <PrimaryButton text="Fechar" onClick={() => setRequiredValidationModalSections(null)} />
              </DialogFooter>
            </>
          ) : null}
        </Dialog>
        <Dialog
          hidden={!runTimelineDialog?.open}
          onDismiss={
            runTimelineDialog?.failed ? () => setRunTimelineDialog(null) : () => undefined
          }
          dialogContentProps={{
            type: DialogType.largeHeader,
            title: runTimelineDialog?.title ?? 'Execução',
            showCloseButton: !!runTimelineDialog?.failed,
          }}
          modalProps={{ isBlocking: true }}
        >
          {runTimelineDialog ? (
            <Stack tokens={{ childrenGap: 10 }} styles={{ root: { maxHeight: '55vh', overflowY: 'auto' } }}>
              {runTimelineDialog.steps.map((st) => (
                <Stack
                  key={st.id}
                  horizontal
                  verticalAlign="start"
                  tokens={{ childrenGap: 12 }}
                  styles={{
                    root: {
                      padding: '8px 10px',
                      borderRadius: 4,
                      borderLeft:
                        st.status === 'done'
                          ? '4px solid #107c10'
                          : st.status === 'error'
                            ? '4px solid #a4262c'
                            : st.status === 'running'
                              ? '4px solid #0078d4'
                              : '4px solid #edebe9',
                      backgroundColor:
                        st.status === 'done' ? '#f1faf1' : st.status === 'error' ? '#fdf3f4' : 'transparent',
                    },
                  }}
                >
                  <Stack styles={{ root: { paddingTop: st.status === 'running' ? 2 : 0 } }}>
                    {st.status === 'running' ? (
                      <Spinner size={SpinnerSize.small} />
                    ) : (
                      <Icon
                        iconName={
                          st.status === 'done'
                            ? 'CompletedSolid'
                            : st.status === 'error'
                              ? 'StatusErrorFull'
                              : 'LocationCircle'
                        }
                        styles={{
                          root: {
                            fontSize: 22,
                            color:
                              st.status === 'done'
                                ? '#107c10'
                                : st.status === 'error'
                                  ? '#a4262c'
                                  : '#a19f9d',
                          },
                        }}
                      />
                    )}
                  </Stack>
                  <Stack tokens={{ childrenGap: 4 }} styles={{ root: { flex: 1 } }}>
                    <Text
                      styles={{
                        root: {
                          color: st.status === 'done' ? '#107c10' : st.status === 'pending' ? '#605e5c' : undefined,
                          fontWeight: st.status === 'running' ? 600 : 400,
                        },
                      }}
                    >
                      {st.label}
                    </Text>
                    {st.status === 'running' && runTimelineDialog.runningDetail ? (
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        {runTimelineDialog.runningDetail}
                      </Text>
                    ) : null}
                  </Stack>
                </Stack>
              ))}
              {runTimelineDialog.failed ? (
                <DialogFooter>
                  <PrimaryButton text="Fechar" onClick={() => setRunTimelineDialog(null)} />
                </DialogFooter>
              ) : null}
            </Stack>
          ) : null}
        </Dialog>
      </Stack>
    </Stack>
  );
};
