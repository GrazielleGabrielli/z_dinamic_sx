import * as React from 'react';
import { useRef, useCallback } from 'react';
import { Dropdown, IDropdownOption, DefaultButton, PrimaryButton, Icon } from '@fluentui/react';
import type { IListViewModeConfig, TViewModePicker } from '../../core/config/types';
import { defaultViewModeIcon, normalizeViewModePicker, pillBadgeVariant } from './viewModePickerLayouts';

export interface IViewModePickerBarProps {
  picker: TViewModePicker | undefined;
  modes: IListViewModeConfig[];
  selectedId: string;
  onSelect: (id: string) => void;
  options: IDropdownOption[];
  showLabel?: boolean;
}

export const ViewModePickerBar: React.FC<IViewModePickerBarProps> = ({
  picker: pickerRaw,
  modes,
  selectedId,
  onSelect,
  options,
  showLabel = false,
}) => {
  const picker = normalizeViewModePicker(pickerRaw);
  const trackRef = useRef<HTMLDivElement>(null);

  const scrollTrack = useCallback((dir: 1 | -1): void => {
    const el = trackRef.current;
    if (!el) return;
    el.scrollBy({ left: dir * 160, behavior: 'smooth' });
  }, []);

  if (picker === 'dropdown') {
    return (
      <div className="dinamicSxViewModeBar dinamicSxViewModeDropdown">
        <Dropdown
          label={showLabel ? 'Visualização' : undefined}
          ariaLabel="Modos de visualização"
          placeholder="Modo de visualização"
          options={options}
          selectedKey={selectedId}
          onChange={(_: React.FormEvent<HTMLDivElement>, opt?: IDropdownOption) => {
            if (opt) onSelect(String(opt.key));
          }}
          styles={{ root: { maxWidth: 240 } }}
        />
      </div>
    );
  }

  const renderBadge = (mode: IListViewModeConfig, selected: boolean): React.ReactNode => {
    if (mode.badgeCount === undefined) return null;
    const variant = pillBadgeVariant(mode, selected);
    return (
      <span
        className={
          variant
            ? `dinamicSxViewModePillBadge dinamicSxViewModePillBadge--${variant}`
            : 'dinamicSxViewModePillBadge'
        }
      >
        {mode.badgeCount}
      </span>
    );
  };

  if (picker === 'buttons') {
    return (
      <div className="dinamicSxViewModeBar" role="presentation">
        <div className="dinamicSxViewModeButtonsRow" role="tablist" aria-label="Modos de visualização">
          {modes.map((m) =>
            selectedId === m.id ? (
              <PrimaryButton
                key={m.id}
                className="dinamicSxViewModeTab"
                role="tab"
                aria-selected={true}
                text={m.label}
                onClick={() => onSelect(m.id)}
                styles={{ root: { minHeight: 32 } }}
              />
            ) : (
              <DefaultButton
                key={m.id}
                className="dinamicSxViewModeTab"
                role="tab"
                aria-selected={false}
                text={m.label}
                onClick={() => onSelect(m.id)}
                styles={{ root: { minHeight: 32 } }}
              />
            )
          )}
        </div>
      </div>
    );
  }

  if (picker === 'segmented') {
    return (
      <div className="dinamicSxViewModeBar" role="presentation">
        <div className="dinamicSxViewModeSegmentedTrack" role="tablist" aria-label="Modos de visualização">
          {modes.map((m) => {
            const selected = selectedId === m.id;
            return (
              <button
                key={m.id}
                type="button"
                className="dinamicSxViewModeSegmentedItem"
                role="tab"
                aria-selected={selected}
                onClick={() => onSelect(m.id)}
              >
                <Icon iconName={defaultViewModeIcon(m)} />
                <span>{m.label}</span>
                {m.badgeCount !== undefined ? (
                  <span className="dinamicSxViewModePillBadge">{m.badgeCount}</span>
                ) : null}
              </button>
            );
          })}
        </div>
      </div>
    );
  }

  if (picker === 'tabs') {
    return (
      <div className="dinamicSxViewModeBar" role="presentation">
        <div className="dinamicSxViewModeUnderlineWrap">
          <div
            ref={trackRef}
            className="dinamicSxViewModeUnderlineTrack"
            role="tablist"
            aria-label="Modos de visualização"
          >
            {modes.map((m) => {
              const selected = selectedId === m.id;
              return (
                <button
                  key={m.id}
                  type="button"
                  className="dinamicSxViewModeUnderlineItem"
                  role="tab"
                  aria-selected={selected}
                  onClick={() => onSelect(m.id)}
                >
                  {m.label}
                </button>
              );
            })}
          </div>
          <button
            type="button"
            className="dinamicSxViewModeUnderlineScrollBtn"
            aria-label="Ver mais modos"
            onClick={() => scrollTrack(1)}
          >
            <Icon iconName="ChevronRight" />
          </button>
        </div>
      </div>
    );
  }

  if (picker === 'pills') {
    return (
      <div className="dinamicSxViewModeBar" role="presentation">
        <div className="dinamicSxViewModePillsRow" role="tablist" aria-label="Modos de visualização">
          {modes.map((m) => {
            const selected = selectedId === m.id;
            return (
              <button
                key={m.id}
                type="button"
                className="dinamicSxViewModePill"
                role="tab"
                aria-selected={selected}
                onClick={() => onSelect(m.id)}
              >
                <span>{m.label}</span>
                {m.badgeCount !== undefined ? renderBadge(m, selected) : null}
              </button>
            );
          })}
        </div>
      </div>
    );
  }

  const withIcons = picker === 'iconsUnderline' || picker === 'iconsBadges';
  const withBadges = picker === 'iconsBadges';

  if (withIcons) {
    return (
      <div className="dinamicSxViewModeBar" role="presentation">
        <div className="dinamicSxViewModeIconRow" role="tablist" aria-label="Modos de visualização">
          {modes.map((m) => {
            const selected = selectedId === m.id;
            return (
              <button
                key={m.id}
                type="button"
                className="dinamicSxViewModeIconItem"
                role="tab"
                aria-selected={selected}
                onClick={() => onSelect(m.id)}
              >
                <Icon iconName={defaultViewModeIcon(m)} />
                <span>{m.label}</span>
                {withBadges && m.badgeCount !== undefined ? renderBadge(m, selected) : null}
              </button>
            );
          })}
        </div>
      </div>
    );
  }

  return null;
};
