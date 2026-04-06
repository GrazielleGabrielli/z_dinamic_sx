import * as React from 'react';
import { Stack, Text, PrimaryButton, DefaultButton, IconButton, Link } from '@fluentui/react';
import type { TFormStepLayoutKind, TFormStepNavButtonsKind } from '../../core/config/types/formManager';

export const FORM_STEP_LAYOUT_OPTIONS: {
  id: TFormStepLayoutKind;
  title: string;
  description: string;
}[] = [
  {
    id: 'rail',
    title: 'Trilho lateral',
    description: 'Números em coluna à esquerda, linha guia e títulos alinhados. Ideal para formulários longos.',
  },
  {
    id: 'segmented',
    title: 'Segmentos',
    description: 'Pílulas horizontais com contraste suave; estilo familiar em apps empresariais.',
  },
  {
    id: 'timeline',
    title: 'Linha do tempo',
    description: 'Pontos sobre uma barra de progresso; mostra claramente o quanto falta.',
  },
  {
    id: 'cards',
    title: 'Cartões',
    description: 'Cada etapa em cartão com sombra; hierarquia visual forte e legível.',
  },
];

const accent = '#0078d4';
const line = '#c8c6c4';
const muted = '#605e5c';
const done = '#107c10';

export interface IFormStepLayoutPickerProps {
  value: TFormStepLayoutKind;
  onChange: (id: TFormStepLayoutKind) => void;
}

export const FormStepLayoutMiniPreview: React.FC<{ kind: TFormStepLayoutKind }> = ({ kind }) => {
  const w = 72;
  const h = 40;
  const dot = (activeDot: boolean, doneDot: boolean): React.ReactNode => (
    <div
      style={{
        width: 8,
        height: 8,
        borderRadius: '50%',
        background: doneDot ? done : activeDot ? accent : '#edebe9',
        flexShrink: 0,
      }}
    />
  );
  if (kind === 'rail') {
    return (
      <div style={{ width: w, height: h, display: 'flex', gap: 4 }}>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 2 }}>
          {dot(true, false)}
          <div style={{ width: 2, flex: 1, background: line, minHeight: 6 }} />
          {dot(false, false)}
        </div>
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', justifyContent: 'space-between', padding: '2px 0' }}>
          <div style={{ height: 4, background: accent, borderRadius: 2, opacity: 0.35 }} />
          <div style={{ height: 4, background: '#edebe9', borderRadius: 2 }} />
        </div>
      </div>
    );
  }
  if (kind === 'segmented') {
    return (
      <div style={{ width: w, height: h, display: 'flex', alignItems: 'center', gap: 3 }}>
        <div style={{ flex: 1, height: 18, borderRadius: 9, background: accent, opacity: 0.9 }} />
        <div style={{ flex: 1, height: 18, borderRadius: 9, background: '#edebe9' }} />
      </div>
    );
  }
  if (kind === 'timeline') {
    return (
      <div style={{ width: w, height: h, display: 'flex', flexDirection: 'column', justifyContent: 'center', gap: 4 }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          {dot(true, false)}
          <div style={{ flex: 1, height: 3, margin: '0 4px', background: `linear-gradient(90deg, ${accent} 50%, ${line} 50%)`, borderRadius: 2 }} />
          {dot(false, false)}
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between' }}>
          <div style={{ width: 20, height: 3, background: '#edebe9', borderRadius: 2 }} />
          <div style={{ width: 20, height: 3, background: '#edebe9', borderRadius: 2 }} />
        </div>
      </div>
    );
  }
  return (
    <div style={{ width: w, height: h, display: 'flex', gap: 4, alignItems: 'stretch' }}>
      <div style={{ flex: 1, borderRadius: 4, border: `2px solid ${accent}`, background: '#f3f9ff', boxShadow: '0 1px 4px rgba(0,0,0,0.08)' }} />
      <div style={{ flex: 1, borderRadius: 4, border: '1px solid #edebe9', background: '#fff' }} />
    </div>
  );
};

export const FormStepLayoutPicker: React.FC<IFormStepLayoutPickerProps> = ({ value, onChange }) => (
  <Stack tokens={{ childrenGap: 12 }} wrap horizontal>
    {FORM_STEP_LAYOUT_OPTIONS.map((opt) => {
      const sel = value === opt.id;
      return (
        <button
          key={opt.id}
          type="button"
          onClick={() => onChange(opt.id)}
          style={{
            display: 'flex',
            flexDirection: 'row',
            alignItems: 'flex-start',
            gap: 12,
            width: '100%',
            maxWidth: 260,
            minWidth: 220,
            boxSizing: 'border-box',
            textAlign: 'left',
            font: 'inherit',
            padding: 14,
            borderRadius: 10,
            border: sel ? `2px solid ${accent}` : '2px solid #edebe9',
            background: sel ? '#f3f9ff' : '#fff',
            cursor: 'pointer',
            boxShadow: sel ? '0 4px 14px rgba(0,120,212,0.12)' : '0 1px 4px rgba(0,0,0,0.06)',
            transition: 'border 0.15s ease, box-shadow 0.15s ease',
          }}
        >
          <span style={{ flexShrink: 0, lineHeight: 0 }}>
            <FormStepLayoutMiniPreview kind={opt.id} />
          </span>
          <span
            style={{
              flex: 1,
              minWidth: 0,
              display: 'block',
            }}
          >
            <span
              style={{
                display: 'block',
                fontWeight: 600,
                color: '#323130',
                fontSize: 14,
                lineHeight: 1.3,
                marginBottom: 6,
              }}
            >
              {opt.title}
            </span>
            <span
              style={{
                display: 'block',
                color: muted,
                fontSize: 12,
                lineHeight: 1.45,
                whiteSpace: 'normal',
                wordBreak: 'break-word',
              }}
            >
              {opt.description}
            </span>
          </span>
        </button>
      );
    })}
  </Stack>
);

export interface IFormStepNavigationProps {
  steps: { id: string; title: string }[];
  activeIndex: number;
  onStepSelect: (index: number) => void;
  layout: TFormStepLayoutKind;
}

export const FormStepNavigation: React.FC<IFormStepNavigationProps> = ({
  steps,
  activeIndex,
  onStepSelect,
  layout,
}) => {
  if (steps.length <= 1) return null;

  const go = (i: number): void => {
    if (i >= 0 && i < steps.length) onStepSelect(i);
  };

  if (layout === 'rail') {
    return (
      <div
        style={{
          borderRadius: 10,
          border: '1px solid #edebe9',
          padding: '16px 20px',
          background: '#faf9f8',
          boxShadow: '0 1px 3px rgba(0,0,0,0.06)',
        }}
      >
        {steps.map((st, i) => {
          const active = i === activeIndex;
          const completed = i < activeIndex;
          return (
            <div key={st.id} style={{ display: 'flex', minHeight: i === steps.length - 1 ? 'auto' : 52 }}>
              <div
                style={{
                  width: 44,
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  flexShrink: 0,
                }}
              >
                <button
                  type="button"
                  onClick={() => go(i)}
                  style={{
                    width: 36,
                    height: 36,
                    borderRadius: '50%',
                    border: 'none',
                    padding: 0,
                    cursor: 'pointer',
                    background: completed ? done : active ? accent : '#edebe9',
                    color: '#fff',
                    fontSize: 15,
                    fontWeight: 700,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    boxShadow: active ? '0 2px 8px rgba(0,120,212,0.35)' : 'none',
                  }}
                  aria-current={active ? 'step' : undefined}
                >
                  {completed ? '✓' : i + 1}
                </button>
                {i < steps.length - 1 && (
                  <div
                    style={{
                      width: 3,
                      flex: 1,
                      minHeight: 16,
                      marginTop: 4,
                      borderRadius: 2,
                      background: completed ? done : line,
                      opacity: completed ? 0.45 : 0.55,
                    }}
                  />
                )}
              </div>
              <div style={{ paddingLeft: 14, paddingBottom: i === steps.length - 1 ? 0 : 8, flex: 1 }}>
                <Text
                  styles={{
                    root: {
                      fontWeight: active ? 700 : completed ? 500 : 400,
                      color: active ? accent : completed ? '#323130' : muted,
                      fontSize: 15,
                      cursor: 'pointer',
                    },
                  }}
                  onClick={() => go(i)}
                >
                  {st.title}
                </Text>
                {active && (
                  <Text variant="small" styles={{ root: { color: muted, marginTop: 4, display: 'block' } }}>
                    Etapa {activeIndex + 1} de {steps.length}
                  </Text>
                )}
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  if (layout === 'segmented') {
    return (
      <div
        style={{
          display: 'flex',
          flexWrap: 'wrap',
          gap: 8,
          padding: 6,
          borderRadius: 12,
          background: '#f3f2f1',
          border: '1px solid #edebe9',
        }}
      >
        {steps.map((st, i) => {
          const active = i === activeIndex;
          return (
            <button
              key={st.id}
              type="button"
              onClick={() => go(i)}
              style={{
                flex: '1 1 auto',
                minWidth: 100,
                padding: '10px 18px',
                borderRadius: 999,
                border: 'none',
                cursor: 'pointer',
                fontWeight: active ? 700 : 500,
                fontSize: 14,
                color: active ? '#fff' : '#323130',
                background: active ? accent : '#fff',
                boxShadow: active ? '0 4px 12px rgba(0,120,212,0.35)' : '0 1px 2px rgba(0,0,0,0.08)',
                transition: 'background 0.15s ease, color 0.15s ease',
              }}
            >
              {st.title}
            </button>
          );
        })}
      </div>
    );
  }

  if (layout === 'timeline') {
    return (
      <div style={{ padding: '8px 4px 4px' }}>
        <div style={{ position: 'relative', marginBottom: 28 }}>
          <div
            style={{
              position: 'absolute',
              left: 0,
              right: 0,
              top: 14,
              height: 4,
              borderRadius: 2,
              background: '#edebe9',
              zIndex: 0,
            }}
          />
          <div
            style={{
              position: 'absolute',
              left: 0,
              top: 14,
              height: 4,
              borderRadius: 2,
              background: `linear-gradient(90deg, ${accent}, ${accent})`,
              width: `${steps.length <= 1 ? 0 : (activeIndex / (steps.length - 1)) * 100}%`,
              zIndex: 1,
              transition: 'width 0.25s ease',
            }}
          />
          <div style={{ display: 'flex', justifyContent: 'space-between', position: 'relative', zIndex: 2 }}>
            {steps.map((st, i) => {
              const active = i === activeIndex;
              const done = i < activeIndex;
              return (
                <button
                  key={st.id}
                  type="button"
                  onClick={() => go(i)}
                  style={{
                    width: 28,
                    height: 28,
                    borderRadius: '50%',
                    border: active ? `3px solid ${accent}` : '3px solid #fff',
                    background: done ? accent : active ? '#fff' : '#edebe9',
                    cursor: 'pointer',
                    boxShadow: active ? '0 0 0 2px rgba(0,120,212,0.25)' : '0 1px 3px rgba(0,0,0,0.12)',
                    padding: 0,
                  }}
                  aria-label={st.title}
                />
              );
            })}
          </div>
        </div>
        <div style={{ display: 'flex', justifyContent: 'space-between', gap: 8 }}>
          {steps.map((st, i) => (
            <div key={st.id} style={{ flex: 1, textAlign: 'center', minWidth: 0 }}>
              <Text
                variant="small"
                styles={{
                  root: {
                    fontWeight: i === activeIndex ? 700 : 400,
                    color: i === activeIndex ? accent : muted,
                    cursor: 'pointer',
                    display: 'block',
                    lineHeight: 1.3,
                  },
                }}
                onClick={() => go(i)}
              >
                {st.title}
              </Text>
            </div>
          ))}
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: 'flex', flexWrap: 'wrap', gap: 12 }}>
      {steps.map((st, i) => {
        const active = i === activeIndex;
        return (
          <button
            key={st.id}
            type="button"
            onClick={() => go(i)}
            style={{
              flex: '1 1 140px',
              maxWidth: 220,
              padding: '14px 16px',
              borderRadius: 10,
              border: active ? `2px solid ${accent}` : '1px solid #edebe9',
              background: active ? '#f3f9ff' : '#fff',
              cursor: 'pointer',
              textAlign: 'left',
              boxShadow: active ? '0 6px 20px rgba(0,120,212,0.15)' : '0 2px 8px rgba(0,0,0,0.06)',
              transition: 'box-shadow 0.15s ease, border 0.15s ease',
            }}
          >
            <Text variant="small" styles={{ root: { color: muted, fontWeight: 600, marginBottom: 6, display: 'block' } }}>
              Etapa {i + 1}
            </Text>
            <Text styles={{ root: { fontWeight: active ? 700 : 500, color: active ? accent : '#323130', fontSize: 14 } }}>
              {st.title}
            </Text>
          </button>
        );
      })}
    </div>
  );
};

export const FORM_STEP_NAV_BUTTONS_OPTIONS: {
  id: TFormStepNavButtonsKind;
  title: string;
  description: string;
}[] = [
  {
    id: 'fluent',
    title: 'Fluent padrão',
    description: 'Secundário «Etapa anterior» e primário «Próxima etapa», como botões Fluent.',
  },
  {
    id: 'pills',
    title: 'Pílulas',
    description: 'Dois botões largos em formato de cápsula.',
  },
  {
    id: 'dots',
    title: 'Bolinhas e setas',
    description: 'Indicadores redondos da etapa atual e botões circulares com setas.',
  },
  {
    id: 'icons',
    title: 'Só ícones',
    description: 'Setas compactas; o significado fica nos tooltips.',
  },
  {
    id: 'links',
    title: 'Ligações de texto',
    description: 'Estilo de hiperligação, visual leve.',
  },
];

export const FormStepNavButtonsMiniPreview: React.FC<{ kind: TFormStepNavButtonsKind }> = ({ kind }) => {
  const w = 72;
  const pill = (filled: boolean): React.ReactNode => (
    <div
      style={{
        height: 14,
        borderRadius: 999,
        flex: 1,
        background: filled ? accent : '#edebe9',
        opacity: filled ? 1 : 0.85,
      }}
    />
  );
  if (kind === 'fluent') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', gap: 4, alignItems: 'center' }}>
        <div style={{ flex: 1, height: 14, borderRadius: 2, border: `1px solid ${line}`, background: '#fff' }} />
        <div style={{ flex: 1, height: 14, borderRadius: 2, background: accent }} />
      </div>
    );
  }
  if (kind === 'pills') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', gap: 5, alignItems: 'center' }}>
        {pill(false)}
        {pill(true)}
      </div>
    );
  }
  if (kind === 'dots') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', alignItems: 'center', gap: 4, justifyContent: 'center' }}>
        <div style={{ width: 12, height: 12, borderRadius: '50%', border: `1px solid ${line}` }} />
        <div style={{ display: 'flex', gap: 3 }}>
          {[0, 1, 2].map((i) => (
            <div
              key={i}
              style={{
                width: 6,
                height: 6,
                borderRadius: '50%',
                background: i === 1 ? accent : '#edebe9',
              }}
            />
          ))}
        </div>
        <div style={{ width: 12, height: 12, borderRadius: '50%', border: `1px solid ${line}` }} />
      </div>
    );
  }
  if (kind === 'icons') {
    return (
      <div style={{ width: w, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10 }}>
        <span style={{ color: muted, fontSize: 14 }}>‹</span>
        <span style={{ color: accent, fontSize: 14 }}>›</span>
      </div>
    );
  }
  return (
    <div style={{ width: w, height: 32, display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
      <div style={{ height: 2, width: 22, background: accent, borderRadius: 1 }} />
      <div style={{ height: 2, width: 22, background: accent, borderRadius: 1 }} />
    </div>
  );
};

export interface IFormStepNavButtonsPickerProps {
  value: TFormStepNavButtonsKind;
  onChange: (id: TFormStepNavButtonsKind) => void;
}

export const FormStepNavButtonsPicker: React.FC<IFormStepNavButtonsPickerProps> = ({ value, onChange }) => (
  <Stack tokens={{ childrenGap: 12 }} wrap horizontal>
    {FORM_STEP_NAV_BUTTONS_OPTIONS.map((opt) => {
      const sel = value === opt.id;
      return (
        <button
          key={opt.id}
          type="button"
          onClick={() => onChange(opt.id)}
          style={{
            display: 'flex',
            flexDirection: 'row',
            alignItems: 'flex-start',
            gap: 12,
            width: '100%',
            maxWidth: 260,
            minWidth: 220,
            boxSizing: 'border-box',
            textAlign: 'left',
            font: 'inherit',
            padding: 14,
            borderRadius: 10,
            border: sel ? `2px solid ${accent}` : '2px solid #edebe9',
            background: sel ? '#f3f9ff' : '#fff',
            cursor: 'pointer',
            boxShadow: sel ? '0 4px 14px rgba(0,120,212,0.12)' : '0 1px 4px rgba(0,0,0,0.06)',
            transition: 'border 0.15s ease, box-shadow 0.15s ease',
          }}
        >
          <span style={{ flexShrink: 0, lineHeight: 0 }}>
            <FormStepNavButtonsMiniPreview kind={opt.id} />
          </span>
          <span style={{ flex: 1, minWidth: 0, display: 'block' }}>
            <span
              style={{
                display: 'block',
                fontWeight: 600,
                color: '#323130',
                fontSize: 14,
                lineHeight: 1.3,
                marginBottom: 6,
              }}
            >
              {opt.title}
            </span>
            <span
              style={{
                display: 'block',
                color: muted,
                fontSize: 12,
                lineHeight: 1.45,
                whiteSpace: 'normal',
                wordBreak: 'break-word',
              }}
            >
              {opt.description}
            </span>
          </span>
        </button>
      );
    })}
  </Stack>
);

export interface IFormStepPrevNextNavProps {
  variant: TFormStepNavButtonsKind;
  stepIndex: number;
  stepCount: number;
  onPrev: () => void;
  onNext: () => void;
  disabled?: boolean;
}

export const FormStepPrevNextNav: React.FC<IFormStepPrevNextNavProps> = ({
  variant,
  stepIndex,
  stepCount,
  onPrev,
  onNext,
  disabled,
}) => {
  if (stepCount <= 1) return null;
  const canPrev = stepIndex > 0;
  const canNext = stepIndex < stepCount - 1;
  const prevLabel = 'Etapa anterior';
  const nextLabel = 'Próxima etapa';

  const pillBtn = (primary: boolean, label: string, onClick: () => void, can: boolean): React.ReactNode => (
    <button
      type="button"
      onClick={onClick}
      disabled={disabled || !can}
      style={{
        padding: '10px 20px',
        borderRadius: 999,
        border: 'none',
        cursor: disabled || !can ? 'default' : 'pointer',
        fontWeight: primary ? 700 : 600,
        fontSize: 14,
        color: primary ? '#fff' : '#323130',
        background: primary ? accent : '#fff',
        boxShadow: primary ? '0 4px 12px rgba(0,120,212,0.35)' : '0 1px 3px rgba(0,0,0,0.1)',
        opacity: can ? 1 : 0.45,
      }}
    >
      {label}
    </button>
  );

  if (variant === 'dots') {
    return (
      <div
        style={{
          flexBasis: '100%',
          width: '100%',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          gap: 12,
          flexWrap: 'wrap',
          paddingTop: 4,
        }}
      >
        <button
          type="button"
          aria-label={prevLabel}
          onClick={onPrev}
          disabled={disabled || !canPrev}
          style={{
            width: 36,
            height: 36,
            borderRadius: '50%',
            border: `1px solid ${line}`,
            background: '#fff',
            cursor: disabled || !canPrev ? 'default' : 'pointer',
            fontSize: 18,
            color: canPrev ? accent : muted,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: 0,
            opacity: canPrev ? 1 : 0.45,
          }}
        >
          ‹
        </button>
        <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap', justifyContent: 'center' }}>
          {(() => {
            const nodes: React.ReactNode[] = [];
            for (let i = 0; i < stepCount; i++) {
              nodes.push(
                <div
                  key={i}
                  style={{
                    width: i === stepIndex ? 10 : 8,
                    height: i === stepIndex ? 10 : 8,
                    borderRadius: '50%',
                    background: i === stepIndex ? accent : i < stepIndex ? done : '#edebe9',
                    boxShadow: i === stepIndex ? '0 0 0 2px rgba(0,120,212,0.25)' : undefined,
                  }}
                />
              );
            }
            return nodes;
          })()}
        </div>
        <button
          type="button"
          aria-label={nextLabel}
          onClick={onNext}
          disabled={disabled || !canNext}
          style={{
            width: 36,
            height: 36,
            borderRadius: '50%',
            border: `1px solid ${line}`,
            background: '#fff',
            cursor: disabled || !canNext ? 'default' : 'pointer',
            fontSize: 18,
            color: canNext ? accent : muted,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: 0,
            opacity: canNext ? 1 : 0.45,
          }}
        >
          ›
        </button>
      </div>
    );
  }

  if (variant === 'pills') {
    return (
      <>
        {canPrev && pillBtn(false, prevLabel, onPrev, canPrev)}
        {canNext && pillBtn(true, nextLabel, onNext, canNext)}
      </>
    );
  }

  if (variant === 'icons') {
    return (
      <>
        {canPrev && (
          <IconButton
            iconProps={{ iconName: 'ChevronLeft' }}
            title={prevLabel}
            ariaLabel={prevLabel}
            onClick={onPrev}
            disabled={disabled || !canPrev}
          />
        )}
        {canNext && (
          <IconButton
            iconProps={{ iconName: 'ChevronRight' }}
            title={nextLabel}
            ariaLabel={nextLabel}
            onClick={onNext}
            disabled={disabled || !canNext}
          />
        )}
      </>
    );
  }

  if (variant === 'links') {
    return (
      <>
        {canPrev && (
          <Link
            href="#"
            onClick={(e) => {
              e.preventDefault();
              if (!disabled && canPrev) onPrev();
            }}
            styles={{ root: { cursor: disabled || !canPrev ? 'default' : 'pointer', opacity: canPrev ? 1 : 0.45 } }}
          >
            ← {prevLabel}
          </Link>
        )}
        {canNext && (
          <Link
            href="#"
            onClick={(e) => {
              e.preventDefault();
              if (!disabled && canNext) onNext();
            }}
            styles={{ root: { cursor: disabled || !canNext ? 'default' : 'pointer', opacity: canNext ? 1 : 0.45 } }}
          >
            {nextLabel} →
          </Link>
        )}
      </>
    );
  }

  return (
    <>
      {canPrev && <DefaultButton text={prevLabel} onClick={onPrev} disabled={disabled || !canPrev} />}
      {canNext && <PrimaryButton text={nextLabel} onClick={onNext} disabled={disabled || !canNext} />}
    </>
  );
};
