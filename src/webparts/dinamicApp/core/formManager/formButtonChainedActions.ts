import type { IDynamicContext } from '../dynamicTokens/types';
import type { TFormButtonAction, IFormButtonActionHttpRequest } from '../config/types/formManager';
import {
  evaluateCondition,
  evaluateFormValueExpression,
  type IFormAttachmentFolderUrlContext,
  type IFormRuleRuntimeContext,
} from './formRuleEngine';
import { isDynamicToken } from '../dynamicTokens';

export type IFormButtonFieldOverlay = {
  show: Set<string>;
  hide: Set<string>;
  showOnStepId?: Record<string, string>;
};

export interface IReduceChainedActionsOpts {
  lookupOptionSnapshots?: IFormRuleRuntimeContext['lookupOptionSnapshots'];
}

function cloneStringSet(s: Set<string>): Set<string> {
  const n = new Set<string>();
  s.forEach((x) => {
    n.add(x);
  });
  return n;
}

function formatJoinedFieldValue(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'object' && v !== null && 'Title' in (v as object)) {
    return String((v as Record<string, unknown>).Title ?? '');
  }
  return String(v);
}

function getDeepValue(root: unknown, path: string): unknown {
  const parts = path
    .split('.')
    .map((p) => p.trim())
    .filter((p) => p.length > 0);
  let cur: unknown = root;
  for (let i = 0; i < parts.length; i++) {
    if (cur === null || cur === undefined) return undefined;
    const p = parts[i];
    if (typeof cur === 'object' && !Array.isArray(cur) && p in (cur as object)) {
      cur = (cur as Record<string, unknown>)[p];
    } else if (Array.isArray(cur) && /^\d+$/.test(p)) {
      cur = cur[Number(p)];
    } else {
      return undefined;
    }
  }
  return cur;
}

export function interpolateButtonMustacheTemplates(
  template: string,
  formValues: Record<string, unknown>,
  httpOutputs: Record<string, unknown>
): string {
  return template.replace(/\{\{([^}]+)\}\}/g, (_m, raw: string) => {
    const full = String(raw).trim();
    if (!full) return '';
    const dot = full.indexOf('.');
    if (dot > 0) {
      const head = full.slice(0, dot).trim();
      const rest = full.slice(dot + 1).trim();
      if (head && rest && Object.prototype.hasOwnProperty.call(httpOutputs, head)) {
        const got = getDeepValue(httpOutputs[head], rest);
        if (got === null || got === undefined) return '';
        if (typeof got === 'object') return JSON.stringify(got);
        return String(got);
      }
    }
    const fieldKey = full.split('/')[0]?.trim() ?? full;
    return formatJoinedFieldValue(formValues[fieldKey]);
  });
}

function buildInterpolatedUrl(
  urlTpl: string,
  query: Array<{ key: string; value: string }> | undefined,
  formValues: Record<string, unknown>,
  httpOutputs: Record<string, unknown>
): string {
  const base = interpolateButtonMustacheTemplates(urlTpl.trim(), formValues, httpOutputs);
  if (!query || !query.length) return base;
  let u: URL;
  try {
    u = new URL(base, typeof window !== 'undefined' ? window.location.href : 'https://local.invalid');
  } catch {
    throw new Error(`URL inválido após interpolação: ${base.slice(0, 120)}`);
  }
  for (let i = 0; i < query.length; i++) {
    const q = query[i];
    const k = interpolateButtonMustacheTemplates(String(q.key ?? '').trim(), formValues, httpOutputs).trim();
    if (!k) continue;
    const v = interpolateButtonMustacheTemplates(String(q.value ?? ''), formValues, httpOutputs);
    u.searchParams.append(k, v);
  }
  return u.toString();
}

function buildHttpRequestFromAction(
  action: IFormButtonActionHttpRequest,
  formValues: Record<string, unknown>,
  httpOutputs: Record<string, unknown>
): { url: string; init: RequestInit } {
  const method = action.method;
  const url = buildInterpolatedUrl(action.url, action.query, formValues, httpOutputs);
  const headersInit = new Headers();
  const hdrs = action.headers ?? [];
  for (let i = 0; i < hdrs.length; i++) {
    const h = hdrs[i];
    const hk = interpolateButtonMustacheTemplates(String(h.key ?? '').trim(), formValues, httpOutputs).trim();
    if (!hk) continue;
    const hv = interpolateButtonMustacheTemplates(String(h.value ?? ''), formValues, httpOutputs);
    headersInit.append(hk, hv);
  }
  const init: RequestInit = { method, headers: headersInit };
  const bodyRaw = action.body;
  const trimmedBody = typeof bodyRaw === 'string' ? bodyRaw.trim() : '';
  if (trimmedBody.length > 0 && method !== 'GET' && method !== 'DELETE') {
    init.body = interpolateButtonMustacheTemplates(bodyRaw ?? '', formValues, httpOutputs);
  }
  return { url, init };
}

async function fetchHttpParsedBody(url: string, init: RequestInit, stepId: string): Promise<unknown> {
  let res: Response;
  try {
    res = await fetch(url, init);
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    throw new Error(`Pedido HTTP (${stepId}) falhou na rede: ${msg}`);
  }
  const text = await res.text();
  let data: unknown = text;
  if (text.trim().length > 0) {
    try {
      data = JSON.parse(text) as unknown;
    } catch {
      data = text;
    }
  } else {
    data = null;
  }
  if (!res.ok) {
    const bit = typeof text === 'string' && text.length > 0 ? text.slice(0, 400) : res.statusText;
    throw new Error(`HTTP ${res.status} (${stepId}): ${bit}`);
  }
  return data;
}

function applySyncButtonAction(
  a: Exclude<TFormButtonAction, IFormButtonActionHttpRequest>,
  next: Record<string, unknown>,
  mergedOverlay: IFormButtonFieldOverlay,
  httpOutputs: Record<string, unknown>,
  dynamicContext: IDynamicContext,
  attachmentFolderUrl: IFormAttachmentFolderUrlContext | undefined,
  userGroupTitles: string[],
  conditionOpts?: IReduceChainedActionsOpts
): void {
  if (a.kind === 'setFieldValue') {
    const tplRaw = String(a.valueTemplate ?? '');
    const trimmed = tplRaw.trim();
    let useExpr = trimmed.startsWith('str:') || trimmed.startsWith('attfolder:');
    if (!useExpr) useExpr = isDynamicToken(trimmed);
    if (!useExpr && dynamicContext && trimmed.indexOf('[') !== -1) useExpr = true;
    const raw = useExpr
      ? evaluateFormValueExpression(tplRaw, next, dynamicContext, attachmentFolderUrl)
      : tplRaw;
    next[a.field] = raw;
  } else if (a.kind === 'joinFields') {
    const tpl = (a.valueTemplate ?? '').trim();
    if (tpl.length > 0) {
      const rawTpl = a.valueTemplate ?? '';
      const interpolated = interpolateButtonMustacheTemplates(rawTpl, next, httpOutputs);
      next[a.targetField] = interpolated;
    } else {
      const parts = a.sourceFields.map((f) => formatJoinedFieldValue(next[f]));
      next[a.targetField] = parts.join(a.separator);
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

export function reduceCustomButtonActions(
  actions: TFormButtonAction[],
  startValues: Record<string, unknown>,
  dynamicContext: IDynamicContext,
  baseOverlay: IFormButtonFieldOverlay,
  attachmentFolderUrl: IFormAttachmentFolderUrlContext | undefined,
  userGroupTitles: string[],
  conditionOpts?: IReduceChainedActionsOpts,
  httpOutputs: Record<string, unknown> = {}
): { mergedValues: Record<string, unknown>; mergedOverlay: IFormButtonFieldOverlay } {
  const next = { ...startValues };
  const mergedOverlay: IFormButtonFieldOverlay = {
    show: cloneStringSet(baseOverlay.show),
    hide: cloneStringSet(baseOverlay.hide),
    ...(baseOverlay.showOnStepId && Object.keys(baseOverlay.showOnStepId).length > 0
      ? { showOnStepId: { ...baseOverlay.showOnStepId } }
      : {}),
  };
  for (let i = 0; i < actions.length; i++) {
    const a = actions[i];
    if (a.kind === 'httpRequest') continue;
    if (a.when && !evaluateCondition(a.when, next, dynamicContext, userGroupTitles, conditionOpts)) {
      continue;
    }
    applySyncButtonAction(a, next, mergedOverlay, httpOutputs, dynamicContext, attachmentFolderUrl, userGroupTitles, conditionOpts);
  }
  return { mergedValues: next, mergedOverlay };
}

export interface IChainedActionsTimelineSink {
  onStepStart: () => void;
  onStepOk: () => void;
  onStepErr: () => void;
}

export async function reduceCustomButtonActionsAsync(
  actions: TFormButtonAction[],
  startValues: Record<string, unknown>,
  dynamicContext: IDynamicContext,
  baseOverlay: IFormButtonFieldOverlay,
  attachmentFolderUrl: IFormAttachmentFolderUrlContext | undefined,
  userGroupTitles: string[],
  conditionOpts: IReduceChainedActionsOpts | undefined,
  timeline: IChainedActionsTimelineSink | undefined
): Promise<{
  mergedValues: Record<string, unknown>;
  mergedOverlay: IFormButtonFieldOverlay;
  httpOutputs: Record<string, unknown>;
  error?: string;
}> {
  let httpOutputs: Record<string, unknown> = {};
  const next = { ...startValues };
  const mergedOverlay: IFormButtonFieldOverlay = {
    show: cloneStringSet(baseOverlay.show),
    hide: cloneStringSet(baseOverlay.hide),
    ...(baseOverlay.showOnStepId && Object.keys(baseOverlay.showOnStepId).length > 0
      ? { showOnStepId: { ...baseOverlay.showOnStepId } }
      : {}),
  };

  for (let i = 0; i < actions.length; i++) {
    const a = actions[i];
    const whenOk = !a.when || evaluateCondition(a.when, next, dynamicContext, userGroupTitles, conditionOpts);
    if (!whenOk) continue;

    try {
      if (timeline) timeline.onStepStart();
      if (a.kind === 'httpRequest') {
        const { url, init } = buildHttpRequestFromAction(a, next, httpOutputs);
        const parsed = await fetchHttpParsedBody(url, init, a.stepId);
        httpOutputs = { ...httpOutputs, [a.stepId]: parsed };
      } else {
        applySyncButtonAction(a, next, mergedOverlay, httpOutputs, dynamicContext, attachmentFolderUrl, userGroupTitles, conditionOpts);
      }
      if (timeline) timeline.onStepOk();
    } catch (e) {
      if (timeline) timeline.onStepErr();
      const msg = e instanceof Error ? e.message : String(e);
      return { mergedValues: next, mergedOverlay, httpOutputs, error: msg };
    }
  }
  return { mergedValues: next, mergedOverlay, httpOutputs };
}
