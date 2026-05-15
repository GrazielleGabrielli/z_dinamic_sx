import * as React from 'react';

export type TTableCardsLayoutKind = 'table' | 'cards';

export interface ITableCardsLayoutToggleProps {
  value: TTableCardsLayoutKind;
  onChange: (v: TTableCardsLayoutKind) => void;
  ariaLabel?: string;
}

function IconListLines(): JSX.Element {
  return (
    <svg width={17} height={13} viewBox="0 0 17 13" aria-hidden focusable="false">
      <circle cx={2.2} cy={2.5} r={0.95} fill="currentColor" />
      <line x1={5} y1={2.5} x2={15} y2={2.5} stroke="currentColor" strokeWidth={1.15} strokeLinecap="round" />
      <circle cx={2.2} cy={10.5} r={0.95} fill="currentColor" />
      <line x1={5} y1={10.5} x2={12.5} y2={10.5} stroke="currentColor" strokeWidth={1.15} strokeLinecap="round" />
    </svg>
  );
}

function IconCardTile(): JSX.Element {
  return (
    <svg width={16} height={16} viewBox="0 0 16 16" aria-hidden focusable="false">
      <rect x={1.6} y={1.6} width={12.8} height={12.8} rx={2.6} ry={2.6} fill="none" stroke="currentColor" strokeWidth={1.1} />
      <line x1={3.3} y1={5.4} x2={12.7} y2={5.4} stroke="currentColor" strokeWidth={1.05} strokeLinecap="round" />
      <line x1={8} y1={5.35} x2={8} y2={13.6} stroke="currentColor" strokeWidth={1.05} strokeLinecap="round" />
    </svg>
  );
}

export const TableCardsLayoutToggle: React.FC<ITableCardsLayoutToggleProps> = ({
  value,
  onChange,
  ariaLabel = 'Alternar entre tabela e cartões',
}) => {
  const track: React.CSSProperties = {
    display: 'inline-flex',
    alignItems: 'center',
    background: '#e8e8e8',
    borderRadius: 11,
    padding: 3,
    gap: 2,
    boxSizing: 'border-box',
  };

  const btn = (active: boolean): React.CSSProperties => ({
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    width: 40,
    height: 34,
    border: 'none',
    borderRadius: 7,
    cursor: 'pointer',
    background: active ? '#ffffff' : 'transparent',
    boxShadow: active ? '0 1px 3px rgba(0, 0, 0, 0.1)' : 'none',
    color: '#323130',
    outline: 'none',
    transition: 'background 0.12s ease, box-shadow 0.12s ease',
  });

  return (
    <div style={track} role="group" aria-label={ariaLabel}>
      <button
        type="button"
        aria-pressed={value === 'table'}
        title="Tabela"
        style={btn(value === 'table')}
        onClick={() => onChange('table')}
      >
        <IconListLines />
      </button>
      <button
        type="button"
        aria-pressed={value === 'cards'}
        title="Cartões"
        style={btn(value === 'cards')}
        onClick={() => onChange('cards')}
      >
        <IconCardTile />
      </button>
    </div>
  );
};
