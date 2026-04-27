import * as React from 'react';
import { useState, useCallback, useEffect, useMemo } from 'react';
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  ProgressIndicator,
  Separator,
} from '@fluentui/react';
import { IDynamicViewConfig, TViewMode } from '../../core/config/types';
import {
  getDefaultConfig,
  getDefaultConfigForMode,
  getDefaultFormManagerConfig,
} from '../../core/config/utils';
import { applyWizardPartial, cloneJson } from '../../core/config/configMemory';
import { IWizardFormState, configToWizardState } from './types';
import { Step1DataSource } from './steps/Step1DataSource';
import { Step2Mode } from './steps/Step2Mode';
import { Step3Dashboard } from './steps/Step3Dashboard';
import { Step3FormStepLayout } from './steps/Step3FormStepLayout';
import { Step4Pagination } from './steps/Step4Pagination';
import { Step5ViewModes } from './steps/Step5ViewModes';

interface IConfigWizardProps {
  siteUrl: string;
  onComplete: (config: IDynamicViewConfig) => void;
  initialValues?: IDynamicViewConfig;
  onCancel?: () => void;
  forcedMode?: TViewMode;
}

const LIST_STEP_LABELS = ['Fonte de dados', 'Modo', 'Dashboard', 'Paginação', 'Modos de visualização'];
const FORM_MANAGER_STEP_LABELS = ['Fonte de dados', 'Modo', 'Layout das etapas'];
const LIST_STEP_LABELS_LOCKED = ['Fonte de dados', 'Dashboard', 'Paginação', 'Modos de visualização'];
const FORM_MANAGER_STEP_LABELS_LOCKED = ['Fonte de dados', 'Layout das etapas'];

function applyForcedMode(base: IDynamicViewConfig, forcedMode: TViewMode): IDynamicViewConfig {
  const merged: IDynamicViewConfig = {
    ...base,
    mode: forcedMode,
  };
  if (forcedMode === 'formManager' && merged.formManager === undefined) {
    merged.formManager = getDefaultFormManagerConfig();
  }
  return merged;
}

function initialDraft(iv?: IDynamicViewConfig, forcedMode?: TViewMode): IDynamicViewConfig {
  const withMemory =
    iv !== undefined
      ? cloneJson({
          ...iv,
          configMemory: iv.configMemory ?? { bySource: {} },
        })
      : {
          ...(forcedMode !== undefined ? getDefaultConfigForMode(forcedMode) : getDefaultConfig()),
          configMemory: { bySource: {} },
        };
  if (forcedMode !== undefined) {
    return applyForcedMode(withMemory, forcedMode);
  }
  return withMemory;
}

function isFormManagerWizard(form: IWizardFormState, forcedMode?: TViewMode): boolean {
  if (forcedMode === 'formManager') return true;
  if (forcedMode === 'list' || forcedMode === 'projectManagement') return false;
  return form.mode === 'formManager';
}

function totalStepsForForm(form: IWizardFormState, forcedMode?: TViewMode): number {
  if (forcedMode === 'formManager') return 2;
  if (forcedMode === 'list' || forcedMode === 'projectManagement') return 4;
  return isFormManagerWizard(form, forcedMode) ? 3 : 5;
}

function stepLabelsForForm(form: IWizardFormState, forcedMode?: TViewMode): string[] {
  if (forcedMode === 'formManager') return FORM_MANAGER_STEP_LABELS_LOCKED;
  if (forcedMode === 'list' || forcedMode === 'projectManagement') return LIST_STEP_LABELS_LOCKED;
  return isFormManagerWizard(form, forcedMode) ? FORM_MANAGER_STEP_LABELS : LIST_STEP_LABELS;
}

function stepIndicesForForm(form: IWizardFormState, forcedMode?: TViewMode): number[] {
  const n = totalStepsForForm(form, forcedMode);
  return Array.from({ length: n }, (_, i) => i + 1);
}

function isStepValid(step: number, form: IWizardFormState, forcedMode?: TViewMode): boolean {
  if (forcedMode === 'formManager') {
    switch (step) {
      case 1:
        return form.title.trim().length > 0;
      case 2:
        return true;
      default:
        return false;
    }
  }
  if (forcedMode === 'list' || forcedMode === 'projectManagement') {
    switch (step) {
      case 1:
        return form.title.trim().length > 0;
      case 2:
        return !form.dashboardEnabled || (form.dashboardType === 'cards' ? form.cardsCount >= 1 : true);
      case 3:
        return form.paginationEnabled ? form.pageSize > 0 : true;
      case 4:
        return (form.viewModes?.length ?? 0) > 0;
      default:
        return false;
    }
  }
  switch (step) {
    case 1:
      return form.title.trim().length > 0;
    case 2:
      return form.mode === 'list' || form.mode === 'projectManagement' || form.mode === 'formManager';
    case 3:
      if (isFormManagerWizard(form, forcedMode)) return true;
      return !form.dashboardEnabled || (form.dashboardType === 'cards' ? form.cardsCount >= 1 : true);
    case 4:
      return form.paginationEnabled ? form.pageSize > 0 : true;
    case 5:
      return (form.viewModes?.length ?? 0) > 0;
    default:
      return false;
  }
}

export const ConfigWizard: React.FC<IConfigWizardProps> = ({
  siteUrl,
  onComplete,
  initialValues,
  onCancel,
  forcedMode,
}) => {
  const isEditMode = initialValues !== undefined;

  const [step, setStep] = useState(1);
  const [draft, setDraft] = useState<IDynamicViewConfig>(() => initialDraft(initialValues, forcedMode));

  const form = useMemo(() => configToWizardState(draft), [draft]);

  const updateForm = useCallback((partial: Partial<IWizardFormState>): void => {
    setDraft((prev) => applyWizardPartial(prev, partial));
  }, []);

  useEffect(() => {
    if (forcedMode !== undefined) {
      setDraft((d) => applyForcedMode(d, forcedMode));
    }
  }, [forcedMode]);

  useEffect(() => {
    if (forcedMode === 'formManager' && step > 2) setStep(2);
    else if ((forcedMode === 'list' || forcedMode === 'projectManagement') && step > 4) setStep(4);
    else if (forcedMode === undefined && form.mode === 'formManager' && step > 3) setStep(3);
  }, [forcedMode, form.mode, step]);

  const totalSteps = totalStepsForForm(form, forcedMode);
  const stepLabels = stepLabelsForForm(form, forcedMode);
  const valid = isStepValid(step, form, forcedMode);
  const canSaveEdit =
    form.title.trim().length > 0 &&
    (form.mode === 'list' || form.mode === 'projectManagement' || form.mode === 'formManager') &&
    (form.mode === 'formManager' || (form.viewModes?.length ?? 0) > 0);

  const buildCurrentConfig = useCallback((): IDynamicViewConfig => {
    const base = cloneJson(draft);
    return forcedMode !== undefined ? applyForcedMode(base, forcedMode) : base;
  }, [draft, forcedMode]);

  const handleNext = (): void => {
    if (!valid) return;
    const max = totalStepsForForm(form, forcedMode);
    if (step < max) {
      setStep((s) => s + 1);
    } else {
      onComplete(buildCurrentConfig());
    }
  };

  const handleSaveAnyStep = (): void => {
    if (!isEditMode || !canSaveEdit) return;
    onComplete(buildCurrentConfig());
  };

  const handleBack = (): void => {
    if (step > 1) setStep((s) => s - 1);
  };

  const renderStep = (): React.ReactElement => {
    if (forcedMode === 'formManager') {
      switch (step) {
        case 1:
          return (
            <Step1DataSource form={form} onChange={updateForm} currentWebServerRelativeUrl={siteUrl} />
          );
        case 2:
          return <Step3FormStepLayout form={form} onChange={updateForm} />;
        default:
          return <></>;
      }
    }
    if (forcedMode === 'list' || forcedMode === 'projectManagement') {
      switch (step) {
        case 1:
          return (
            <Step1DataSource form={form} onChange={updateForm} currentWebServerRelativeUrl={siteUrl} />
          );
        case 2:
          return <Step3Dashboard form={form} onChange={updateForm} />;
        case 3:
          return <Step4Pagination form={form} onChange={updateForm} />;
        case 4:
          return (
            <Step5ViewModes
              form={form}
              listTitle={form.title}
              listWebServerRelativeUrl={form.dataSourceWebServerRelativeUrl}
              pageWebServerRelativeUrl={siteUrl}
              onChange={updateForm}
            />
          );
        default:
          return <></>;
      }
    }
    if (isFormManagerWizard(form, forcedMode)) {
      switch (step) {
        case 1:
          return (
            <Step1DataSource form={form} onChange={updateForm} currentWebServerRelativeUrl={siteUrl} />
          );
        case 2:
          return <Step2Mode form={form} onChange={updateForm} />;
        case 3:
          return <Step3FormStepLayout form={form} onChange={updateForm} />;
        default:
          return <></>;
      }
    }
    switch (step) {
      case 1:
          return (
            <Step1DataSource form={form} onChange={updateForm} currentWebServerRelativeUrl={siteUrl} />
          );
        case 2:
        return <Step2Mode form={form} onChange={updateForm} />;
      case 3:
        return <Step3Dashboard form={form} onChange={updateForm} />;
      case 4:
        return <Step4Pagination form={form} onChange={updateForm} />;
      case 5:
        return (
          <Step5ViewModes
            form={form}
            listTitle={form.title}
            listWebServerRelativeUrl={form.dataSourceWebServerRelativeUrl}
            pageWebServerRelativeUrl={siteUrl}
            onChange={updateForm}
          />
        );
      default:
        return <></>;
    }
  };

  const isFormFlow = isFormManagerWizard(form, forcedMode);

  return (
    <div
      style={{
        minHeight: '100%',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: 24,
        background: '#faf9f8',
      }}
    >
      <div
        style={{
          width: '100%',
          maxWidth: isFormFlow ? 960 : 580,
          background: '#fff',
          borderRadius: 12,
          border: '1px solid #edebe9',
          boxShadow: '0 2px 12px rgba(0,0,0,0.08)',
          overflow: 'hidden',
        }}
      >
        <div style={{ padding: '24px 32px 0' }}>
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Stack tokens={{ childrenGap: 2 }}>
              <Text variant="xxLarge" styles={{ root: { fontWeight: 700, color: '#0078d4' } }}>
                FlexView
              </Text>
              <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                {isEditMode ? 'Editar configuração' : 'Configuração inicial'}
              </Text>
            </Stack>
            <Stack horizontalAlign="end" tokens={{ childrenGap: 2 }}>
              <Text variant="small" styles={{ root: { color: '#605e5c', fontWeight: 600 } }}>
                Passo {step} de {totalSteps}
              </Text>
              <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                {stepLabels[step - 1] ?? ''}
              </Text>
            </Stack>
          </Stack>
          <div style={{ marginTop: 16 }}>
            <ProgressIndicator
              percentComplete={step / totalSteps}
              styles={{ itemProgress: { padding: 0 } }}
            />
          </div>
          <Stack horizontal tokens={{ childrenGap: 0 }} styles={{ root: { marginTop: 12, marginBottom: 4 } }}>
            {stepIndicesForForm(form, forcedMode).map((s) => (
              <button
                key={s}
                type="button"
                onClick={() => setStep(s)}
                style={{
                  flex: 1,
                  padding: '8px 4px',
                  border: 'none',
                  borderBottom: step === s ? '2px solid #0078d4' : '2px solid transparent',
                  background: step === s ? 'rgba(0,120,212,0.08)' : 'transparent',
                  cursor: 'pointer',
                  fontSize: 12,
                  fontWeight: step === s ? 600 : 400,
                  color: step === s ? '#0078d4' : '#605e5c',
                }}
              >
                {s}. {stepLabelsForForm(form, forcedMode)[s - 1] ?? ''}
              </button>
            ))}
          </Stack>
        </div>

        <Separator />

        <div style={{ padding: '24px 32px' }}>{renderStep()}</div>

        <Separator />

        <div style={{ padding: '16px 32px', background: '#faf9f8' }}>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            {step > 1 && <DefaultButton text="Voltar" onClick={handleBack} />}
            {isEditMode && (
              <PrimaryButton
                text="Salvar configuração"
                onClick={handleSaveAnyStep}
                disabled={!canSaveEdit}
              />
            )}
            <PrimaryButton
              text={step === totalSteps ? 'Concluir configuração' : 'Próximo'}
              onClick={handleNext}
              disabled={!valid}
            />
            {onCancel !== undefined && (
              <DefaultButton text="Cancelar" onClick={onCancel} />
            )}
          </Stack>
        </div>
      </div>
    </div>
  );
};
