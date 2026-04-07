import * as React from 'react';
import {
  Stack,
  Spinner,
  SpinnerSize,
  Shimmer,
  ShimmerElementType,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  Text,
  type IDropdownOption,
} from '@fluentui/react';
import type {
  IFormCustomButtonConfig,
  IFormManagerConfig,
  TFormDataLoadingUiKind,
  TFormSubmitLoadingUiKind,
} from '../../core/config/types/formManager';

const cardBox = (): { root: Record<string, string | number> } => ({
  root: {
    padding: 20,
    borderRadius: 4,
    border: '1px solid #edebe9',
    background: '#faf9f8',
    maxWidth: 420,
    width: '100%',
  },
});

export const FORM_DATA_LOADING_DROPDOWN_OPTIONS: IDropdownOption[] = [
  { key: 'spinner', text: 'Spinner Fluent (padrão)' },
  { key: 'spinnerLarge', text: 'Spinner grande' },
  { key: 'shimmer', text: 'Blocos shimmer' },
  { key: 'progress', text: 'Barra de progresso indeterminada' },
  { key: 'cardShimmer', text: 'Cartão com avatar + linhas' },
];

export const FORM_SUBMIT_LOADING_DROPDOWN_OPTIONS: IDropdownOption[] = [
  { key: 'overlay', text: 'Sobreposição + spinner (padrão)' },
  { key: 'topProgress', text: 'Barra de progresso no topo' },
  { key: 'formShimmer', text: 'Shimmer sobre o formulário' },
  { key: 'belowButtons', text: 'Spinner por baixo dos botões' },
  { key: 'infoBar', text: 'Faixa informativa' },
];

export const FORM_SUBMIT_LOADING_INHERIT_KEY = '__inherit';

export function resolveFormDataLoadingKind(fm: IFormManagerConfig): TFormDataLoadingUiKind {
  return fm.formDataLoadingKind ?? 'spinner';
}

export function resolveSubmitLoadingKind(
  fm: IFormManagerConfig,
  btn?: IFormCustomButtonConfig
): TFormSubmitLoadingUiKind {
  if (btn?.submitLoadingKind) return btn.submitLoadingKind;
  return fm.defaultSubmitLoadingKind ?? 'overlay';
}

export function FormDataLoadingView(props: {
  kind: TFormDataLoadingUiKind;
  message: string;
}): JSX.Element {
  const { kind, message } = props;
  return (
    <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { padding: 32, minHeight: 160 } }}>
      {kind === 'spinner' && <Spinner label={message} />}
      {kind === 'spinnerLarge' && <Spinner size={SpinnerSize.large} label={message} />}
      {kind === 'shimmer' && (
        <Stack tokens={{ childrenGap: 12 }} styles={cardBox()}>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{message}</Text>
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 12, width: '70%' }]} />
          <Shimmer
            shimmerElements={[
              { type: ShimmerElementType.line, height: 28, width: '40%' },
              { type: ShimmerElementType.gap, width: 12, height: 28 },
              { type: ShimmerElementType.line, height: 28, width: '40%' },
            ]}
          />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 48, width: '100%' }]} />
        </Stack>
      )}
      {kind === 'progress' && (
        <Stack tokens={{ childrenGap: 16 }} styles={{ root: { width: '100%', maxWidth: 400 } }}>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{message}</Text>
          <ProgressIndicator label={message} />
        </Stack>
      )}
      {kind === 'cardShimmer' && (
        <Stack tokens={{ childrenGap: 10 }} styles={cardBox()}>
          <Shimmer width={120} shimmerElements={[{ type: ShimmerElementType.circle, height: 40 }]} />
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{message}</Text>
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 10, width: '90%' }]} />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 10, width: '60%' }]} />
        </Stack>
      )}
    </Stack>
  );
}

export function FormSubmitLoadingChrome(props: {
  kind: TFormSubmitLoadingUiKind;
  active: boolean;
  message: string;
}): JSX.Element | null {
  const { kind, active, message } = props;
  if (!active) return null;
  if (kind === 'overlay') {
    return (
      <div
        style={{
          position: 'absolute',
          inset: 0,
          background: 'rgba(255,255,255,0.72)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 20,
          borderRadius: 2,
        }}
      >
        <Spinner size={SpinnerSize.medium} label={message} />
      </div>
    );
  }
  if (kind === 'topProgress') {
    return (
      <Stack styles={{ root: { marginBottom: 8 } }}>
        <ProgressIndicator label={message} />
      </Stack>
    );
  }
  if (kind === 'formShimmer') {
    return (
      <div
        style={{
          position: 'absolute',
          inset: 0,
          zIndex: 15,
          padding: 12,
          background: 'rgba(250,249,248,0.92)',
          pointerEvents: 'none',
        }}
      >
        <Stack tokens={{ childrenGap: 10 }}>
          <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>{message}</Text>
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 14, width: '100%' }]} />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 14, width: '85%' }]} />
          <Shimmer shimmerElements={[{ type: ShimmerElementType.line, height: 14, width: '70%' }]} />
        </Stack>
      </div>
    );
  }
  if (kind === 'belowButtons') {
    return (
      <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 8 } }}>
        <Spinner size={SpinnerSize.small} />
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>{message}</Text>
      </Stack>
    );
  }
  if (kind === 'infoBar') {
    return (
      <MessageBar messageBarType={MessageBarType.info} styles={{ root: { marginBottom: 8 } }}>
        {message}
      </MessageBar>
    );
  }
  return null;
}
