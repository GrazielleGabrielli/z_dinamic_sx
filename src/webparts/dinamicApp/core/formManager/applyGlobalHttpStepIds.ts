import type { IFormCustomButtonConfig, TFormButtonAction } from '../config/types/formManager';

export function applyGlobalHttpStepIds(buttons: IFormCustomButtonConfig[]): IFormCustomButtonConfig[] {
  let n = 0;
  return buttons.map((btn) => ({
    ...btn,
    actions: btn.actions.map((a: TFormButtonAction) => {
      if (a.kind !== 'httpRequest') return a;
      n += 1;
      const stepId = `http${n}`;
      return a.stepId === stepId ? a : { ...a, stepId };
    }),
  }));
}
