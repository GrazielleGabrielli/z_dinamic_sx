import * as React from 'react';
import { DINAMIC_SX_TOOLBAR_CLASS } from './listToolbarLayouts';

export type TTableCardsLayoutKind = 'table' | 'cards';

export interface ITableCardsLayoutToggleProps {
  value: TTableCardsLayoutKind;
  onChange: (v: TTableCardsLayoutKind) => void;
  ariaLabel?: string;
}

function IconListLines(): JSX.Element {
  return (
    <svg viewBox="0 0 17 13" aria-hidden focusable="false">
      <circle cx={2.2} cy={2.5} r={0.95} fill="currentColor" />
      <line x1={5} y1={2.5} x2={15} y2={2.5} stroke="currentColor" strokeWidth={1.15} strokeLinecap="round" />
      <circle cx={2.2} cy={10.5} r={0.95} fill="currentColor" />
      <line x1={5} y1={10.5} x2={12.5} y2={10.5} stroke="currentColor" strokeWidth={1.15} strokeLinecap="round" />
    </svg>
  );
}

function IconCardTile(): JSX.Element {
  return (
    <svg viewBox="0 0 16 16" aria-hidden focusable="false">
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
  const C = DINAMIC_SX_TOOLBAR_CLASS;

  return (
    <div className={C.layoutToggle} role="group" aria-label={ariaLabel}>
      <button
        type="button"
        className={C.layoutToggleBtn}
        aria-pressed={value === 'table'}
        title="Tabela"
        onClick={() => onChange('table')}
      >
        <IconListLines />
      </button>
      <button
        type="button"
        className={C.layoutToggleBtn}
        aria-pressed={value === 'cards'}
        title="Cartões"
        onClick={() => onChange('cards')}
      >
        <IconCardTile />
      </button>
    </div>
  );
};
