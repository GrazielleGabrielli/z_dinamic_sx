import type { IFormStepConfig, TFormManagerFormMode } from '../config/types/formManager';

export const ALL_FORM_MANAGER_MODES: readonly TFormManagerFormMode[] = ['create', 'edit', 'view'];

export function stepVisibleInFormMode(step: IFormStepConfig, mode: TFormManagerFormMode): boolean {
  const sel = step.showInFormModes;
  if (!sel || sel.length === 0) return true;
  return sel.indexOf(mode) !== -1;
}

export function toggleStepShowInFormMode(
  current: TFormManagerFormMode[] | undefined,
  mode: TFormManagerFormMode,
  checked: boolean
): TFormManagerFormMode[] | undefined {
  let next: TFormManagerFormMode[] =
    current && current.length > 0 ? current.slice() : [...ALL_FORM_MANAGER_MODES];
  const idx = next.indexOf(mode);
  if (checked && idx === -1) next.push(mode);
  if (!checked && idx !== -1) next.splice(idx, 1);
  if (next.length === 0) next = ['create'];
  const all =
    next.length === ALL_FORM_MANAGER_MODES.length &&
    ALL_FORM_MANAGER_MODES.every((m) => next.indexOf(m) !== -1);
  if (all) return undefined;
  return next.slice().sort((a, b) => ALL_FORM_MANAGER_MODES.indexOf(a) - ALL_FORM_MANAGER_MODES.indexOf(b));
}
