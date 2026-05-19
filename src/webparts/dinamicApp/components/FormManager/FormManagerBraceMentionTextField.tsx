import * as React from 'react';
import { useState, useRef, useEffect, useLayoutEffect, useCallback, useMemo } from 'react';
import { createPortal } from 'react-dom';
import { TextField, Text, type ITextField } from '@fluentui/react';
import type { IDropdownOption } from '@fluentui/react';

export const FORM_BRACE_MENTION_PORTAL_ATTR = 'data-dinamic-brace-mention';

type TMentionItem = {
  key: string;
  insert: string;
  primary: string;
  secondary: string;
};

function overflowScrollAncestors(start: HTMLElement | undefined): HTMLElement[] {
  const seen = new Set<HTMLElement>();
  const out: HTMLElement[] = [];
  let n: HTMLElement | undefined = start?.parentElement ?? undefined;
  while (n) {
    const st = window.getComputedStyle(n);
    if (
      /(auto|scroll|overlay)/.test(st.overflowY) ||
      /(auto|scroll|overlay)/.test(st.overflowX) ||
      /(auto|scroll|overlay)/.test(st.overflow)
    ) {
      if (!seen.has(n)) {
        seen.add(n);
        out.push(n);
      }
    }
    n = n.parentElement ?? undefined;
  }
  const root = document.documentElement;
  if (!seen.has(root)) out.push(root);
  return out;
}

/** Menção ativa: acabou de escrever `{{` e ainda não fechou com `}}`. */
export function getActiveBraceMentionRange(
  value: string,
  caret: number
): { from: number; to: number; filter: string } | undefined {
  if (caret < 2) return undefined;
  const before = value.slice(0, caret);
  const openIdx = before.lastIndexOf('{{');
  if (openIdx === -1) return undefined;
  if (openIdx > 0) {
    const prev = before[openIdx - 1];
    if (
      prev !== ' ' &&
      prev !== '\n' &&
      prev !== '\t' &&
      prev !== '(' &&
      prev !== '[' &&
      prev !== ']' &&
      prev !== '{' &&
      prev !== '}' &&
      prev !== ';' &&
      prev !== ',' &&
      prev !== '/' &&
      prev !== ':' &&
      prev !== '?' &&
      prev !== '&' &&
      prev !== '=' &&
      prev !== '%' &&
      prev !== '-' &&
      prev !== '_' &&
      prev !== '+' &&
      prev !== '*' &&
      prev !== '.' &&
      prev !== '|' &&
      prev !== '#' &&
      prev !== "'" &&
      prev !== '"' &&
      prev !== '\\'
    ) {
      return undefined;
    }
  }
  const afterOpen = before.slice(openIdx + 2);
  if (afterOpen.includes('}}')) return undefined;
  return { from: openIdx, to: caret, filter: afterOpen };
}

export function buildFieldBraceMentionItems(filter: string, fieldOptions: IDropdownOption[]): TMentionItem[] {
  const f = filter.trim().toLowerCase();
  const match = (s: string): boolean => !f || s.toLowerCase().includes(f);
  const out: TMentionItem[] = [];
  for (let i = 0; i < fieldOptions.length; i++) {
    const opt = fieldOptions[i];
    const k = String(opt.key);
    if (!k) continue;
    const ins = `{{${k}}}`;
    const lab = String(opt.text ?? k);
    if (match(k) || match(lab) || match(ins)) {
      out.push({
        key: `f-${k}-${i}`,
        insert: ins,
        primary: lab,
        secondary: ins,
      });
    }
  }
  return out;
}

export interface IFormManagerBraceMentionTextFieldProps {
  label: string;
  description?: string;
  value: string;
  onChange: (next: string) => void;
  fieldOptions: IDropdownOption[];
  multiline?: boolean;
  rows?: number;
}

export function FormManagerBraceMentionTextField({
  label,
  description,
  value,
  onChange,
  fieldOptions,
  multiline = false,
  rows = 2,
}: IFormManagerBraceMentionTextFieldProps): JSX.Element {
  const [mentionOpen, setMentionOpen] = useState(false);
  const [mentionRange, setMentionRange] = useState<{ from: number; to: number; filter: string } | null>(null);
  const [mentionHighlight, setMentionHighlight] = useState(0);
  const [mentionListPos, setMentionListPos] = useState<{
    top: number;
    left: number;
    width: number;
  } | null>(null);
  const tfRef = useRef<ITextField | null>(null);
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const mentionPortalRef = useRef<HTMLDivElement | null>(null);
  const pendingCaretRef = useRef<number | undefined>(undefined);
  const mentionRangeRef = useRef<{ from: number; to: number; filter: string } | undefined>(undefined);

  mentionRangeRef.current = mentionRange ?? undefined;

  const measureMentionListPos = useCallback((): void => {
    const w = wrapRef.current;
    if (!w) return;
    const r = w.getBoundingClientRect();
    setMentionListPos({
      top: r.bottom + 4,
      left: r.left,
      width: r.width,
    });
  }, []);

  const mentionItems = useMemo(() => {
    if (!mentionOpen || !mentionRange) return [];
    return buildFieldBraceMentionItems(mentionRange.filter, fieldOptions);
  }, [mentionOpen, mentionRange, fieldOptions]);

  useLayoutEffect(() => {
    const p = pendingCaretRef.current;
    if (p === undefined || !tfRef.current) return;
    pendingCaretRef.current = undefined;
    const tf = tfRef.current;
    tf.focus();
    requestAnimationFrame(() => {
      try {
        tf.setSelectionRange(p, p);
      } catch {
        //
      }
    });
  }, [value]);

  useLayoutEffect(() => {
    const show = mentionOpen && mentionItems.length > 0;
    if (!show) {
      setMentionListPos(null);
      return;
    }
    measureMentionListPos();
  }, [mentionOpen, mentionItems.length, value, measureMentionListPos]);

  useEffect(() => {
    const show = mentionOpen && mentionItems.length > 0;
    if (!show) return;
    measureMentionListPos();
    const roots = overflowScrollAncestors(wrapRef.current ?? undefined);
    const upd = (): void => {
      measureMentionListPos();
    };
    roots.forEach((el) => el.addEventListener('scroll', upd, true));
    window.addEventListener('resize', upd);
    return () => {
      roots.forEach((el) => el.removeEventListener('scroll', upd, true));
      window.removeEventListener('resize', upd);
    };
  }, [mentionOpen, mentionItems.length, measureMentionListPos]);

  useEffect(() => {
    const onDocDown = (e: MouseEvent): void => {
      const t = e.target as Node;
      if (wrapRef.current?.contains(t) || mentionPortalRef.current?.contains(t)) return;
      setMentionOpen(false);
    };
    document.addEventListener('mousedown', onDocDown);
    return () => document.removeEventListener('mousedown', onDocDown);
  }, []);

  useEffect(() => {
    if (mentionOpen && mentionItems.length === 0) setMentionOpen(false);
  }, [mentionOpen, mentionItems.length]);

  useEffect(() => {
    setMentionHighlight(0);
  }, [mentionRange?.filter]);

  const applyMentionInsert = useCallback(
    (insertText: string): void => {
      const r = mentionRangeRef.current;
      if (!r) return;
      const cur = value;
      const next = cur.slice(0, r.from) + insertText + cur.slice(r.to);
      pendingCaretRef.current = r.from + insertText.length;
      setMentionOpen(false);
      setMentionRange(null);
      onChange(next);
    },
    [value, onChange]
  );

  const handleChange = useCallback(
    (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, v: string | undefined): void => {
      const raw = v ?? '';
      const el = ev.target as HTMLTextAreaElement;
      const caret =
        typeof el.selectionStart === 'number' ? el.selectionStart : raw.length;
      const range = getActiveBraceMentionRange(raw, caret);
      if (range) {
        const items = buildFieldBraceMentionItems(range.filter, fieldOptions);
        if (items.length > 0) {
          setMentionRange(range);
          setMentionOpen(true);
          setMentionHighlight(0);
        } else {
          setMentionOpen(false);
          setMentionRange(null);
        }
      } else {
        setMentionOpen(false);
        setMentionRange(null);
      }
      onChange(raw);
    },
    [onChange, fieldOptions]
  );

  const handleKeyDown = useCallback(
    (ev: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>): void => {
      if (!mentionOpen || mentionItems.length === 0) return;
      if (ev.key === 'ArrowDown') {
        ev.preventDefault();
        setMentionHighlight((h) => Math.min(mentionItems.length - 1, h + 1));
      } else if (ev.key === 'ArrowUp') {
        ev.preventDefault();
        setMentionHighlight((h) => Math.max(0, h - 1));
      } else if (ev.key === 'Enter' && !ev.shiftKey && multiline) {
        ev.preventDefault();
        const it = mentionItems[mentionHighlight];
        if (it) applyMentionInsert(it.insert);
      } else if (ev.key === 'Tab' && !ev.shiftKey) {
        ev.preventDefault();
        const it = mentionItems[mentionHighlight];
        if (it) applyMentionInsert(it.insert);
      } else if (ev.key === 'Enter' && !multiline) {
        ev.preventDefault();
        const it = mentionItems[mentionHighlight];
        if (it) applyMentionInsert(it.insert);
      } else if (ev.key === 'Escape') {
        ev.preventDefault();
        setMentionOpen(false);
        setMentionRange(null);
      }
    },
    [mentionOpen, mentionItems, mentionHighlight, applyMentionInsert, multiline]
  );

  return (
    <div ref={wrapRef} style={{ position: 'relative', width: '100%' }}>
      <TextField
        label={label}
        description={description}
        multiline={multiline}
        rows={rows}
        value={value}
        componentRef={tfRef}
        onChange={handleChange}
        onKeyDown={handleKeyDown}
      />
      {mentionOpen && mentionItems.length > 0 && mentionListPos
        ? createPortal(
            <div
              ref={mentionPortalRef}
              {...{ [FORM_BRACE_MENTION_PORTAL_ATTR]: '' }}
              role="listbox"
              aria-label="Campos (placeholders {{}})"
              style={{
                position: 'fixed',
                left: mentionListPos.left,
                top: mentionListPos.top,
                width: mentionListPos.width,
                maxWidth:
                  typeof window !== 'undefined'
                    ? Math.max(0, window.innerWidth - mentionListPos.left - 8)
                    : mentionListPos.width,
                zIndex: 10000000,
                minWidth: 280,
                maxHeight: 280,
                overflowY: 'auto',
                border: '1px solid #edebe9',
                borderRadius: 4,
                boxShadow: '0 4px 12px rgba(0,0,0,0.12)',
                background: '#ffffff',
                boxSizing: 'border-box',
              }}
              onMouseDown={(e) => e.preventDefault()}
            >
              {mentionItems.map((it, idx) => (
                <div
                  key={it.key}
                  role="option"
                  aria-selected={idx === mentionHighlight}
                  style={{
                    padding: '8px 10px',
                    cursor: 'pointer',
                    background: idx === mentionHighlight ? '#edebe9' : 'transparent',
                    borderBottom:
                      idx < mentionItems.length - 1 ? '1px solid #f3f2f1' : undefined,
                  }}
                  onMouseEnter={() => setMentionHighlight(idx)}
                  onMouseDown={(e) => {
                    e.preventDefault();
                    applyMentionInsert(it.insert);
                  }}
                >
                  <Text variant="small" styles={{ root: { fontWeight: 600, display: 'block' } }}>
                    {it.primary}
                  </Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c', fontSize: 11 } }}>
                    {it.secondary}
                  </Text>
                </div>
              ))}
            </div>,
            document.body
          )
        : null}
    </div>
  );
}
