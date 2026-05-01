function isExcludedFromNativeToolbar(el: Element, excludeInside: HTMLElement): boolean {
  return excludeInside.contains(el);
}

function isLikelyVisible(el: HTMLElement): boolean {
  const rect = el.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) return false;
  const st = window.getComputedStyle(el);
  if (st.visibility === 'hidden' || st.display === 'none') return false;
  if (el.hasAttribute('disabled')) return false;
  if (el.getAttribute('aria-disabled') === 'true') return false;
  return true;
}

function clickIfEligible(el: Element | null, excludeInside: HTMLElement): boolean {
  if (!el || !(el instanceof HTMLElement)) return false;
  if (isExcludedFromNativeToolbar(el, excludeInside)) return false;
  if (!isLikelyVisible(el)) return false;
  el.click();
  return true;
}

function queryAutomationButtons(doc: Document, automationId: string): HTMLElement[] {
  const sel = `[data-automation-id="${automationId}"]`;
  const nodes = Array.from(doc.querySelectorAll(sel));
  const out: HTMLElement[] = [];
  for (let i = 0; i < nodes.length; i++) {
    const node = nodes[i];
    if (node instanceof HTMLButtonElement) out.push(node);
    else if (node instanceof HTMLElement) {
      const inner = node.querySelector('button');
      if (inner) out.push(inner);
      else out.push(node);
    }
  }
  return out;
}

const EDIT_AUTOMATION_IDS = ['editPageCommandButton', 'editPageButton', 'EditPage', 'pageEditButton'];

const SAVE_AUTOMATION_IDS = [
  'SiteHeaderSubmit',
  'publishOrDiscardButton',
  'savePublishButton',
  'submitButton',
  'PrimaryPublishButton',
];

export function tryClickSharePointPageEditButton(excludeInside: HTMLElement): boolean {
  const doc = excludeInside.ownerDocument ?? document;
  for (let i = 0; i < EDIT_AUTOMATION_IDS.length; i++) {
    const candidates = queryAutomationButtons(doc, EDIT_AUTOMATION_IDS[i]);
    for (let j = 0; j < candidates.length; j++) {
      if (clickIfEligible(candidates[j], excludeInside)) return true;
    }
  }
  const buttons = doc.querySelectorAll('button, [role="button"]');
  for (let k = 0; k < buttons.length; k++) {
    const el = buttons[k];
    if (!(el instanceof HTMLElement)) continue;
    if (isExcludedFromNativeToolbar(el, excludeInside)) continue;
    const aria = (el.getAttribute('aria-label') || '').toLowerCase();
    const title = (el.getAttribute('title') || '').toLowerCase();
    const name = (el.getAttribute('name') || '').toLowerCase();
    const combined = `${aria} ${title} ${name}`;
    if (
      combined.includes('editar página') ||
      combined.includes('editar pagina') ||
      combined.includes('edit page') ||
      combined.includes('modifier la page')
    ) {
      if (clickIfEligible(el, excludeInside)) return true;
    }
  }
  return false;
}

function forEachButtonInSiteHeader(doc: Document, fn: (el: HTMLElement) => boolean): boolean {
  const roots = doc.querySelectorAll(
    '[data-automation-id="SiteHeader"], [data-automation-id="SuiteNavWrapper"], #SuiteNavPlaceholder'
  );
  if (roots.length === 0) return false;
  for (let r = 0; r < roots.length; r++) {
    const root = roots[r];
    const buttons = root.querySelectorAll('button, [role="button"]');
    for (let i = 0; i < buttons.length; i++) {
      const el = buttons[i];
      if (el instanceof HTMLElement && fn(el)) return true;
    }
  }
  return false;
}

export function tryClickSharePointPageSaveOrPublishButton(excludeInside: HTMLElement): boolean {
  const doc = excludeInside.ownerDocument ?? document;
  for (let i = 0; i < SAVE_AUTOMATION_IDS.length; i++) {
    const candidates = queryAutomationButtons(doc, SAVE_AUTOMATION_IDS[i]);
    for (let j = 0; j < candidates.length; j++) {
      if (clickIfEligible(candidates[j], excludeInside)) return true;
    }
  }
  const headerHit = forEachButtonInSiteHeader(doc, (el) => {
    if (isExcludedFromNativeToolbar(el, excludeInside)) return false;
    const aria = (el.getAttribute('aria-label') || '').toLowerCase();
    const title = (el.getAttribute('title') || '').toLowerCase();
    const text = (el.textContent || '').trim().toLowerCase();
    const name = (el.getAttribute('name') || '').toLowerCase();
    const combined = `${aria} ${title} ${name} ${text}`;
    if (
      combined.includes('republish') ||
      combined.includes('republicar') ||
      combined.includes('publish') ||
      combined.includes('publicar') ||
      combined.includes('save and') ||
      combined.includes('salvar e') ||
      (combined.includes('salvar') &&
        (combined.includes('fechar') ||
          combined.includes('rascunho') ||
          combined.includes('alterações') ||
          combined.includes('changes') ||
          combined.includes('página') ||
          combined.includes('pagina') ||
          combined.includes('page')))
    ) {
      return clickIfEligible(el, excludeInside);
    }
    return false;
  });
  if (headerHit) return true;
  const buttons = doc.querySelectorAll('button, [role="button"]');
  for (let k = 0; k < buttons.length; k++) {
    const el = buttons[k];
    if (!(el instanceof HTMLElement)) continue;
    if (isExcludedFromNativeToolbar(el, excludeInside)) continue;
    const aria = (el.getAttribute('aria-label') || '').toLowerCase();
    const title = (el.getAttribute('title') || '').toLowerCase();
    const text = (el.textContent || '').trim().toLowerCase();
    const combined = `${aria} ${title} ${text}`;
    if (
      combined.includes('republish') ||
      combined.includes('republicar') ||
      combined.includes('publish') ||
      combined.includes('publicar')
    ) {
      if (clickIfEligible(el, excludeInside)) return true;
    }
  }
  return false;
}

export function runNativePagePersistAfterPropertyWrite(
  excludeInside: HTMLElement,
  isReadMode: boolean,
  afterEditDelayMs: number
): void {
  if (isReadMode) {
    const opened = tryClickSharePointPageEditButton(excludeInside);
    window.setTimeout(() => {
      tryClickSharePointPageSaveOrPublishButton(excludeInside);
    }, opened ? afterEditDelayMs : 0);
  } else {
    tryClickSharePointPageSaveOrPublishButton(excludeInside);
  }
}

export const DINAMIC_SX_OPEN_SLIDER_EVENT = 'dinamic-sx-open-slider';
export const DINAMIC_SX_CLOSE_SLIDER_EVENT = 'dinamic-sx-close-slider';

const SWITCH_INPUT_CLICK_DELAY_MS = 1500;

function scheduleClickSwitchInputsByIdPrefix(): void {
  window.setTimeout(() => {
    document.querySelectorAll('input[id^="switch-"]').forEach((input) => {
      if (input instanceof HTMLElement) {
        const id = input.id || '(sem id)';
        console.log('[DinamicSX] click switch:', id);
        input.click();
      }
    });
  }, SWITCH_INPUT_CLICK_DELAY_MS);
}

const nativeEditSaveBridgeHosts = new Set<HTMLElement>();
let nativeEditSaveBridgeHandler: ((event: MouseEvent) => void) | undefined;

function hasAncestorAutomationId(el: Element, ids: readonly string[]): boolean {
  let n: Element | null = el;
  for (let i = 0; i < 16 && n; i++) {
    const aid = n.getAttribute('data-automation-id');
    if (aid && ids.includes(aid)) return true;
    n = n.parentElement;
  }
  return false;
}

function isUnderSharePointPageChrome(el: Element): boolean {
  return (
    el.closest(
      '[data-automation-id="SiteHeader"], [data-automation-id="SuiteNavWrapper"], #SuiteNavPlaceholder, [data-automation-id="CommandBar"], .ms-CommandBar-primaryCommand'
    ) !== null
  );
}

function matchesNativeEditPageControl(ctrl: HTMLElement): boolean {
  const aid = ctrl.getAttribute('data-automation-id');
  if (aid && EDIT_AUTOMATION_IDS.includes(aid)) return true;
  const aria = (ctrl.getAttribute('aria-label') || '').toLowerCase();
  const title = (ctrl.getAttribute('title') || '').toLowerCase();
  const name = (ctrl.getAttribute('name') || '').toLowerCase();
  const combinedMeta = `${aria} ${title} ${name}`;
  if (
    combinedMeta.includes('editar página') ||
    combinedMeta.includes('editar pagina') ||
    combinedMeta.includes('edit page') ||
    combinedMeta.includes('modifier la page')
  ) {
    return true;
  }
  const tl = (ctrl.textContent || '').trim().toLowerCase();
  if (tl === 'editar' || tl === 'edit') {
    return isUnderSharePointPageChrome(ctrl);
  }
  return false;
}

function isNativeEditPageToolbarClick(target: Element): boolean {
  if (hasAncestorAutomationId(target, EDIT_AUTOMATION_IDS)) return true;
  const ctrl = target.closest('button, [role="button"], a');
  if (ctrl instanceof HTMLElement) return matchesNativeEditPageControl(ctrl);
  return false;
}

function matchesNativeSavePublishControl(ctrl: HTMLElement): boolean {
  const aid = ctrl.getAttribute('data-automation-id');
  if (aid && SAVE_AUTOMATION_IDS.includes(aid)) return true;
  const aria = (ctrl.getAttribute('aria-label') || '').toLowerCase();
  const title = (ctrl.getAttribute('title') || '').toLowerCase();
  const text = (ctrl.textContent || '').trim().toLowerCase();
  const name = (ctrl.getAttribute('name') || '').toLowerCase();
  const combined = `${aria} ${title} ${name} ${text}`;
  if (
    combined.includes('republish') ||
    combined.includes('republicar') ||
    combined.includes('publish') ||
    combined.includes('publicar') ||
    combined.includes('save and') ||
    combined.includes('salvar e')
  ) {
    return true;
  }
  if (
    combined.includes('salvar') &&
    (combined.includes('fechar') ||
      combined.includes('rascunho') ||
      combined.includes('alterações') ||
      combined.includes('changes') ||
      combined.includes('página') ||
      combined.includes('pagina') ||
      combined.includes('page'))
  ) {
    return true;
  }
  const rawTrim = (ctrl.textContent || '').trim();
  if (rawTrim === 'Salvar' || rawTrim === 'Save') {
    return isUnderSharePointPageChrome(ctrl);
  }
  return false;
}

function isNativeSavePublishToolbarClick(target: Element): boolean {
  if (hasAncestorAutomationId(target, SAVE_AUTOMATION_IDS)) return true;
  const ctrl = target.closest('button, [role="button"], a');
  if (ctrl instanceof HTMLElement) return matchesNativeSavePublishControl(ctrl);
  return false;
}

function onNativeToolbarBridgeClick(event: MouseEvent): void {
  const raw = event.target;
  const el = raw instanceof Element ? raw : raw instanceof Node ? raw.parentElement : null;
  if (!el) return;
  let insideHost = false;
  nativeEditSaveBridgeHosts.forEach((host) => {
    if (host.contains(el)) insideHost = true;
  });
  if (insideHost) return;

  if (isNativeEditPageToolbarClick(el)) {
    console.log('Botão nativo Editar clicado');
    window.dispatchEvent(new CustomEvent(DINAMIC_SX_OPEN_SLIDER_EVENT));
    scheduleClickSwitchInputsByIdPrefix();
  } else if (isNativeSavePublishToolbarClick(el)) {
    console.log('Botão nativo Salvar/Publicar clicado');
    window.dispatchEvent(new CustomEvent(DINAMIC_SX_CLOSE_SLIDER_EVENT));
  }
}

export function registerNativeEditSaveToolbarBridge(hostElement: HTMLElement): () => void {
  nativeEditSaveBridgeHosts.add(hostElement);
  if (nativeEditSaveBridgeHosts.size === 1) {
    nativeEditSaveBridgeHandler = onNativeToolbarBridgeClick;
    document.addEventListener('click', nativeEditSaveBridgeHandler, true);
  }
  return (): void => {
    nativeEditSaveBridgeHosts.delete(hostElement);
    if (nativeEditSaveBridgeHosts.size === 0 && nativeEditSaveBridgeHandler !== undefined) {
      document.removeEventListener('click', nativeEditSaveBridgeHandler, true);
      nativeEditSaveBridgeHandler = undefined;
    }
  };
}
