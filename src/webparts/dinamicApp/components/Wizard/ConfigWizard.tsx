import * as React from 'react';
import { useState, useCallback } from 'react';
import {
  Stack,
  Text,
  PrimaryButton,
  DefaultButton,
  ProgressIndicator,
  Separator,
} from '@fluentui/react';
import { IDynamicViewConfig } from '../../core/config/types';
import { buildConfig } from '../../core/config/builders';
import { getDefaultFormManagerConfig } from '../../core/config/utils';
import { IWizardFormState, WIZARD_INITIAL_STATE, configToWizardState } from './types';
import { Step1DataSource } from './steps/Step1DataSource';
import { Step2Mode } from './steps/Step2Mode';
import { Step3Dashboard } from './steps/Step3Dashboard';
import { Step4Pagination } from './steps/Step4Pagination';
import { Step5ViewModes } from './steps/Step5ViewModes';

interface IConfigWizardProps {
  siteUrl: string;
  onComplete: (config: IDynamicViewConfig) => void;
  initialValues?: IDynamicViewConfig;
  onCancel?: () => void;
}

const TOTAL_STEPS = 5;
const STEP_LABELS = ['Fonte de dados', 'Modo', 'Dashboard', 'Paginação', 'Modos de visualização'];

function isStepValid(step: number, form: IWizardFormState): boolean {
  switch (step) {
    case 1: return form.title.trim().length > 0;
    case 2: return form.mode === 'list' || form.mode === 'projectManagement' || form.mode === 'formManager';
    case 3: return !form.dashboardEnabled || (form.dashboardType === 'cards' ? form.cardsCount >= 1 : true);
    case 4: return form.paginationEnabled ? form.pageSize > 0 : true;
    case 5: return (form.viewModes?.length ?? 0) > 0;
    default: return false;
  }
}

export const ConfigWizard: React.FC<IConfigWizardProps> = ({
  siteUrl,
  onComplete,
  initialValues,
  onCancel,
}) => {
  const isEditMode = initialValues !== undefined;

  const [step, setStep] = useState(1);
  const [form, setForm] = useState<IWizardFormState>(() =>
    isEditMode ? configToWizardState(initialValues) : WIZARD_INITIAL_STATE
  );

  const updateForm = useCallback((partial: Partial<IWizardFormState>): void => {
    setForm((prev) => ({ ...prev, ...partial }));
  }, []);

  const valid = isStepValid(step, form);
  const canSaveEdit =
    form.title.trim().length > 0 &&
    (form.mode === 'list' || form.mode === 'projectManagement' || form.mode === 'formManager') &&
    (form.viewModes?.length ?? 0) > 0;

  const buildCurrentConfig = useCallback((): IDynamicViewConfig => {
    const existingCards = initialValues?.dashboard.cards ?? [];
    const existingChartSeries = initialValues?.dashboard.chartSeries ?? [];
    const existingListView = initialValues?.listView;
    return buildConfig({
      dataSource: { kind: form.kind, title: form.title },
      mode: form.mode,
      dashboard: {
        enabled: form.dashboardEnabled,
        dashboardType: form.dashboardType,
        cardsCount: form.cardsCount,
        cards: existingCards,
        chartType: form.chartType,
        chartSeries: existingChartSeries,
      },
      pagination: {
        enabled: form.paginationEnabled,
        pageSize: form.pageSize,
        pageSizeOptions: form.pageSizeOptions,
      },
      listView: {
        ...existingListView,
        viewModes: form.viewModes,
        activeViewModeId: form.activeViewModeId,
      },
      projectManagement: initialValues?.projectManagement,
      formManager:
        form.mode === 'formManager'
          ? (initialValues?.formManager ?? getDefaultFormManagerConfig())
          : initialValues?.formManager,
    });
  }, [form, initialValues]);

  const handleNext = (): void => {
    if (!valid) return;
    if (step < TOTAL_STEPS) {
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
    switch (step) {
      case 1: return <Step1DataSource form={form} onChange={updateForm} />;
      case 2: return <Step2Mode form={form} onChange={updateForm} />;
      case 3: return <Step3Dashboard form={form} onChange={updateForm} />;
      case 4: return <Step4Pagination form={form} onChange={updateForm} />;
      case 5: return <Step5ViewModes form={form} listTitle={form.title} onChange={updateForm} />;
      default: return <></>;
    }
  };

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
          maxWidth: 580,
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
                Passo {step} de {TOTAL_STEPS}
              </Text>
              <Text variant="small" styles={{ root: { color: '#a19f9d' } }}>
                {STEP_LABELS[step - 1]}
              </Text>
            </Stack>
          </Stack>
          <div style={{ marginTop: 16 }}>
            <ProgressIndicator
              percentComplete={step / TOTAL_STEPS}
              styles={{ itemProgress: { padding: 0 } }}
            />
          </div>
          <Stack horizontal tokens={{ childrenGap: 0 }} styles={{ root: { marginTop: 12, marginBottom: 4 } }}>
            {[1, 2, 3, 4, 5].map((s) => (
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
                {s}. {STEP_LABELS[s - 1]}
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
              text={step === TOTAL_STEPS ? 'Concluir configuração' : 'Próximo'}
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
