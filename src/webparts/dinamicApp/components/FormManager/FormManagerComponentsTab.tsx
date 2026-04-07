import * as React from 'react';
import { useState, useEffect } from 'react';
import {
  Stack,
  Text,
  Spinner,
  SpinnerSize,
  Shimmer,
  ShimmerElementType,
  ProgressIndicator,
} from '@fluentui/react';
import { Dropdown } from '@fluentui/react';
import type {
  TFormStepLayoutKind,
  TFormStepNavButtonsKind,
  TFormDataLoadingUiKind,
  TFormSubmitLoadingUiKind,
} from '../../core/config/types/formManager';
import { FormStepLayoutPicker, FormStepNavButtonsPicker } from './FormStepLayoutUi';
import {
  FormDataLoadingView,
  FORM_DATA_LOADING_DROPDOWN_OPTIONS,
  FORM_SUBMIT_LOADING_DROPDOWN_OPTIONS,
} from './FormLoadingUi';

const loadingCardStyles = (): { root: Record<string, string | number> } => ({
  root: {
    padding: 16,
    borderRadius: 4,
    border: '1px solid #edebe9',
    background: '#faf9f8',
  },
});

export function FormManagerComponentsLoadingLayouts(): JSX.Element {
  const [idx, setIdx] = useState(0);
  useEffect(() => {
    const id = window.setInterval(() => setIdx((i) => (i + 1) % 3), 2600);
    return () => window.clearInterval(id);
  }, []);

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        A carregar metadados dos campos da lista…
      </Text>
      {idx === 0 && (
        <Stack horizontalAlign="center" tokens={{ childrenGap: 16 }} styles={loadingCardStyles()}>
          <Spinner size={SpinnerSize.large} label="A sincronizar componentes" />
          <Text variant="small" styles={{ root: { color: '#605e5c', textAlign: 'center' } }}>
            Pré-visualização 1 de 3 — indicador centrado
          </Text>
        </Stack>
      )}
      {idx === 1 && (
        <Stack tokens={{ childrenGap: 12 }} styles={loadingCardStyles()}>
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            Pré-visualização 2 de 3 — esqueleto (shimmer)
          </Text>
          <Shimmer
            shimmerElements={[{ type: ShimmerElementType.line, height: 10, width: '55%' }]}
          />
          <Shimmer
            shimmerElements={[
              { type: ShimmerElementType.line, height: 32, width: '42%' },
              { type: ShimmerElementType.gap, width: 12, height: 32 },
              { type: ShimmerElementType.line, height: 32, width: '42%' },
            ]}
          />
          <Shimmer
            shimmerElements={[{ type: ShimmerElementType.line, height: 72, width: '100%' }]}
          />
        </Stack>
      )}
      {idx === 2 && (
        <Stack tokens={{ childrenGap: 12 }} styles={loadingCardStyles()}>
          <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
            Pré-visualização 3 de 3 — progresso e faixas
          </Text>
          <ProgressIndicator label="A carregar opções de interface" />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 6, width: '28%' }]} />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 6, width: '72%' }]} />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 6, width: '48%' }]} />
        </Stack>
      )}
      <Text variant="small" styles={{ root: { color: '#a19f9d', fontStyle: 'italic' } }}>
        As três pré-visualizações alternam automaticamente enquanto espera.
      </Text>
    </Stack>
  );
}

export interface IFormManagerComponentsTabContentProps {
  loading: boolean;
  stepLayout: TFormStepLayoutKind;
  onStepLayoutChange: (v: TFormStepLayoutKind) => void;
  stepNavButtons: TFormStepNavButtonsKind;
  onStepNavButtonsChange: (v: TFormStepNavButtonsKind) => void;
  formDataLoadingKind: TFormDataLoadingUiKind;
  onFormDataLoadingKindChange: (v: TFormDataLoadingUiKind) => void;
  defaultSubmitLoadingKind: TFormSubmitLoadingUiKind;
  onDefaultSubmitLoadingKindChange: (v: TFormSubmitLoadingUiKind) => void;
}

export function FormManagerComponentsTabContent(props: IFormManagerComponentsTabContentProps): JSX.Element {
  if (props.loading) {
    return <FormManagerComponentsLoadingLayouts />;
  }
  return (
    <Stack tokens={{ childrenGap: 12 }}>
      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Carregar formulário / dados</Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Vista formulário: ao carregar campos da lista ou ao abrir um item para editar. Pré-visualização abaixo.
      </Text>
      <Dropdown
        label="Estilo de loading (dados)"
        options={FORM_DATA_LOADING_DROPDOWN_OPTIONS}
        selectedKey={props.formDataLoadingKind}
        onChange={(_, o) =>
          o && props.onFormDataLoadingKindChange(String(o.key) as TFormDataLoadingUiKind)
        }
      />
      <Stack
        horizontalAlign="center"
        styles={{
          root: {
            border: '1px solid #edebe9',
            borderRadius: 4,
            padding: 8,
            background: '#ffffff',
            minHeight: 140,
          },
        }}
      >
        <FormDataLoadingView
          kind={props.formDataLoadingKind}
          message="Pré-visualização — carregar campos / item"
        />
      </Stack>
      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>Gravar (padrão)</Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Usado em Enviar, Rascunho e em botões personalizados que não definem override. Cada botão pode escolher outro estilo na aba Botões.
      </Text>
      <Dropdown
        label="Estilo de loading ao gravar (padrão)"
        options={FORM_SUBMIT_LOADING_DROPDOWN_OPTIONS}
        selectedKey={props.defaultSubmitLoadingKind}
        onChange={(_, o) =>
          o && props.onDefaultSubmitLoadingKindChange(String(o.key) as TFormSubmitLoadingUiKind)
        }
      />
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Com mais do que uma etapa (aba Estrutura), o utilizador vê o passador de etapas neste estilo. Os botões de
        navegação no rodapé são configurados abaixo (são independentes do layout visual em cima).
      </Text>
      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
        Layout das etapas no formulário
      </Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Quando existir mais de uma etapa, o utilizador vê a navegação neste estilo.
      </Text>
      <FormStepLayoutPicker value={props.stepLayout} onChange={props.onStepLayoutChange} />
      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
        Botões «Etapa anterior» / «Próxima etapa»
      </Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Estilo apenas dos botões de navegação no rodapé (não altera o passador de etapas em cima).
      </Text>
      <FormStepNavButtonsPicker value={props.stepNavButtons} onChange={props.onStepNavButtonsChange} />
    </Stack>
  );
}
