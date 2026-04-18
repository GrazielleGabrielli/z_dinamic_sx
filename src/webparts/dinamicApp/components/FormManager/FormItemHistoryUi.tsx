import * as React from 'react';
import { useEffect, useMemo, useState } from 'react';
import {
  Stack,
  Text,
  Spinner,
  MessageBar,
  MessageBarType,
  Panel,
  PanelType,
  Modal,
  DefaultButton,
  useTheme,
  type ITheme,
} from '@fluentui/react';
import type {
  IFormManagerActionLogConfig,
  TFormHistoryPresentationKind,
  TFormHistoryLayoutKind,
} from '../../core/config/types/formManager';
import { hexToRgbaString } from '../../core/formManager/formCustomButtonTheme';
import { ItemsService, FieldsService } from '../../../../services';

export interface IFormItemHistoryUiProps {
  actionLog: IFormManagerActionLogConfig | undefined;
  sourceItemId: number;
  presentationKind: TFormHistoryPresentationKind;
  layoutKind?: TFormHistoryLayoutKind;
  isOpen: boolean;
  onDismiss: () => void;
  title: string;
  subtitle?: string;
  /** Cor de realce (passador); omitido = primária do tema Fluent. */
  accentColor?: string;
}

interface IHistoryUiColors {
  accent: string;
  bodyText: string;
  bodySubtext: string;
  mutedHint: string;
  border: string;
  borderStrong: string;
  cardBg: string;
  listRowBg: string;
  timelineLine: string;
}

function historyColorsFromTheme(theme: ITheme, accentOverride?: string): IHistoryUiColors {
  const p = theme.palette;
  const s = theme.semanticColors;
  return {
    accent: accentOverride ?? p.themePrimary,
    bodyText: s.bodyText ?? p.neutralPrimary,
    bodySubtext: p.neutralSecondary,
    mutedHint: p.neutralTertiaryAlt ?? p.neutralTertiary ?? p.neutralSecondary,
    border: p.neutralLight,
    borderStrong: p.neutralQuaternaryAlt ?? p.neutralLight,
    cardBg: p.white,
    listRowBg: p.neutralLighterAlt ?? p.neutralLighter,
    timelineLine: p.neutralQuaternaryAlt ?? '#e1dfdd',
  };
}

interface IAuditEntry {
  key: string;
  actionLabel: string;
  createdStr: string;
  who: string;
  html: string;
}

function actionLabelFromItemTitle(title: string): string {
  const t = title.trim();
  const sep = ' · ';
  let idx = t.lastIndexOf(sep);
  while (idx > 0) {
    const tail = t.slice(idx + sep.length).trim();
    if (
      /\d{1,2}\/\d{1,2}\/\d{2,4}/.test(tail) ||
      /\d{4}-\d{2}-\d{2}/.test(tail) ||
      /\d{1,2}:\d{2}(:\d{2})?/.test(tail)
    ) {
      return t.slice(0, idx).trim() || t;
    }
    idx = t.lastIndexOf(sep, idx - 1);
  }
  return t;
}

function authorDisplay(row: Record<string, unknown>): string {
  const a = row.Author;
  if (a && typeof a === 'object' && a !== null && 'Title' in a) {
    return String((a as { Title?: string }).Title ?? '').trim();
  }
  if (typeof a === 'string') return a.trim();
  return '';
}

function formatCreatedValue(created: unknown): string {
  if (created == null || created === '') return '—';
  const d =
    created instanceof Date
      ? created
      : typeof created === 'string' || typeof created === 'number'
        ? new Date(created)
        : null;
  if (!d || Number.isNaN(d.getTime())) return '—';
  return d.toLocaleString(undefined, { dateStyle: 'short', timeStyle: 'short' });
}

function renderHtmlBlock(html: string, compact: boolean, colors: IHistoryUiColors): React.ReactNode {
  return html ? (
    <div
      className="form-audit-log-html"
      style={{ fontSize: compact ? 12 : 14, color: colors.bodyText, lineHeight: 1.45 }}
      dangerouslySetInnerHTML={{ __html: html }}
    />
  ) : (
    <Text
      variant="small"
      styles={{
        root: { color: colors.mutedHint, fontStyle: 'italic', fontSize: compact ? 11 : 12 },
      }}
    >
      (sem texto no campo de ação)
    </Text>
  );
}

function entryHeadline(e: IAuditEntry, colors: IHistoryUiColors): React.ReactNode {
  return (
    <>
      <span style={{ fontWeight: 600, color: colors.bodyText }}>{e.actionLabel}</span>
      <span style={{ color: colors.bodySubtext, fontWeight: 400 }}> · {e.createdStr}</span>
    </>
  );
}

function entryAuthorLine(e: IAuditEntry, fontSize: number, colors: IHistoryUiColors): React.ReactNode {
  return (
    <Text variant="small" styles={{ root: { color: colors.bodySubtext, fontSize, marginTop: 2 } }}>
      {e.who || '—'}
    </Text>
  );
}

function renderAuditEntries(
  entries: IAuditEntry[],
  layoutKind: TFormHistoryLayoutKind,
  colors: IHistoryUiColors
): React.ReactNode {
  if (layoutKind === 'compact') {
    return (
      <Stack tokens={{ childrenGap: 0 }}>
        {entries.map((e, i) => (
          <div
            key={e.key}
            style={{
              padding: '8px 0',
              borderBottom: i < entries.length - 1 ? `1px solid ${colors.border}` : undefined,
            }}
          >
            <Text variant="small" styles={{ root: { fontSize: 11 } }}>
              {entryHeadline(e, colors)}
            </Text>
            {entryAuthorLine(e, 11, colors)}
            <div style={{ marginTop: e.html ? 6 : 4 }}>{renderHtmlBlock(e.html, true, colors)}</div>
          </div>
        ))}
      </Stack>
    );
  }

  if (layoutKind === 'timeline') {
    return (
      <div style={{ position: 'relative', paddingLeft: 22 }}>
        <div
          style={{
            position: 'absolute',
            left: 5,
            top: 8,
            bottom: 8,
            width: 2,
            background: colors.timelineLine,
          }}
        />
        <Stack tokens={{ childrenGap: 14 }}>
          {entries.map((e) => (
            <div key={e.key} style={{ position: 'relative' }}>
              <div
                style={{
                  position: 'absolute',
                  left: -19,
                  top: 2,
                  width: 12,
                  height: 12,
                  borderRadius: '50%',
                  background: colors.accent,
                  border: `2px solid ${colors.cardBg}`,
                  boxShadow: `0 0 0 1px ${colors.borderStrong}`,
                }}
              />
              <Text variant="small" styles={{ root: { color: colors.bodyText } }}>
                {entryHeadline(e, colors)}
              </Text>
              {entryAuthorLine(e, 12, colors)}
              <div style={{ marginTop: 6 }}>{renderHtmlBlock(e.html, false, colors)}</div>
            </div>
          ))}
        </Stack>
      </div>
    );
  }

  if (layoutKind === 'cards') {
    const cardShadow = `0 2px 8px ${hexToRgbaString(colors.bodyText, 0.08)}`;
    return (
      <Stack tokens={{ childrenGap: 12 }}>
        {entries.map((e) => (
          <div
            key={e.key}
            style={{
              padding: 16,
              borderRadius: 8,
              background: colors.cardBg,
              boxShadow: cardShadow,
              border: `1px solid ${colors.border}`,
            }}
          >
            <Text variant="small" styles={{ root: { color: colors.bodyText } }}>
              {entryHeadline(e, colors)}
            </Text>
            {entryAuthorLine(e, 12, colors)}
            <div style={{ marginTop: 10 }}>{renderHtmlBlock(e.html, false, colors)}</div>
          </div>
        ))}
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 8 }}>
      {entries.map((e) => (
        <Stack
          key={e.key}
          tokens={{ childrenGap: 6 }}
          styles={{
            root: {
              padding: '10px 12px',
              borderRadius: 4,
              border: `1px solid ${colors.border}`,
              background: colors.listRowBg,
            },
          }}
        >
          <Text variant="small" styles={{ root: { color: colors.bodyText } }}>
            {entryHeadline(e, colors)}
          </Text>
          {entryAuthorLine(e, 12, colors)}
          {renderHtmlBlock(e.html, false, colors)}
        </Stack>
      ))}
    </Stack>
  );
}

export const FormItemHistoryUi: React.FC<IFormItemHistoryUiProps> = ({
  actionLog,
  sourceItemId,
  presentationKind,
  layoutKind = 'list',
  isOpen,
  onDismiss,
  title,
  subtitle,
  accentColor,
}) => {
  const theme = useTheme();
  const colors = useMemo(
    () => historyColorsFromTheme(theme, accentColor),
    [theme, accentColor]
  );
  const itemsService = useMemo(() => new ItemsService(), []);
  const fieldsService = useMemo(() => new FieldsService(), []);
  const [loading, setLoading] = useState(false);
  const [err, setErr] = useState<string | undefined>(undefined);
  const [rows, setRows] = useState<Record<string, unknown>[]>([]);
  const [resolvedActionField, setResolvedActionField] = useState<string>('');

  useEffect(() => {
    if (!isOpen || !sourceItemId || sourceItemId < 1) return;
    const logList = actionLog?.listTitle?.trim();
    const actionField = actionLog?.actionFieldInternalName?.trim();
    const linkField = actionLog?.sourceListLookupFieldInternalName?.trim();
    if (!logList || !actionField || !linkField) {
      setErr(
        'Indique na configuração do gestor (aba «Lista de logs») a lista de registo, o campo multilinhas e o lookup de vínculo à lista principal. Ative o histórico na aba «Componentes».'
      );
      setRows([]);
      setResolvedActionField('');
      return;
    }
    setErr(undefined);
    setLoading(true);
    setResolvedActionField(actionField);
    const filter = `${linkField}Id eq ${sourceItemId}`;
    void (async (): Promise<void> => {
      try {
        const meta = await fieldsService.getVisibleFields(logList);
        const names = new Set(meta.map((f) => f.InternalName));
        const select: string[] = ['Id', 'Title', 'Created', actionField];
        if (names.has('Author')) select.push('Author');
        const data = await itemsService.getItems<Record<string, unknown>>(logList, {
          filter,
          orderBy: { field: 'Created', ascending: false },
          top: 200,
          fieldMetadata: meta,
          select,
        });
        setRows(Array.isArray(data) ? data : []);
      } catch (e) {
        setErr(e instanceof Error ? e.message : String(e));
        setRows([]);
      } finally {
        setLoading(false);
      }
    })();
  }, [
    isOpen,
    sourceItemId,
    actionLog?.listTitle,
    actionLog?.actionFieldInternalName,
    actionLog?.sourceListLookupFieldInternalName,
    fieldsService,
    itemsService,
  ]);

  const entries: IAuditEntry[] = useMemo(() => {
    const out: IAuditEntry[] = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const id = r.Id;
      const key =
        typeof id === 'number' || typeof id === 'string' ? String(id) : `r-${i}`;
      const rawHtml = resolvedActionField ? r[resolvedActionField] : undefined;
      const html =
        typeof rawHtml === 'string'
          ? rawHtml
          : rawHtml !== undefined && rawHtml !== null
            ? String(rawHtml)
            : '';
      const createdStr = formatCreatedValue(r.Created);
      const who = authorDisplay(r);
      const lineTitle = typeof r.Title === 'string' ? r.Title : String(r.Title ?? '—');
      const actionLabel = actionLabelFromItemTitle(lineTitle);
      out.push({ key, actionLabel, createdStr, who, html });
    }
    return out;
  }, [rows, resolvedActionField]);

  const body = (
    <Stack tokens={{ childrenGap: 12 }}>
      {subtitle && (
        <Text variant="small" styles={{ root: { color: colors.bodySubtext } }}>
          {subtitle}
        </Text>
      )}
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      {loading && (
        <Spinner
          label="A carregar registos de auditoria…"
          styles={{ circle: { borderTopColor: colors.accent } }}
        />
      )}
      {!loading && !err && entries.length === 0 && (
        <Text variant="small" styles={{ root: { color: colors.bodySubtext } }}>
          Nenhum registo na lista de auditoria para este item (filtro pelo lookup configurado).
        </Text>
      )}
      {!loading && !err && entries.length > 0 && renderAuditEntries(entries, layoutKind, colors)}
    </Stack>
  );

  if (presentationKind === 'collapse') {
    if (!isOpen) return null;
    return (
      <Stack
        tokens={{ childrenGap: 12 }}
        styles={{
          root: {
            marginTop: 8,
            padding: 16,
            borderRadius: 4,
            border: `1px solid ${colors.border}`,
            background: colors.cardBg,
          },
        }}
      >
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600, color: colors.bodyText } }}>
            {title}
          </Text>
          <DefaultButton text="Fechar" onClick={onDismiss} />
        </Stack>
        {body}
      </Stack>
    );
  }

  if (presentationKind === 'modal') {
    return (
      <Modal isOpen={isOpen} onDismiss={onDismiss} isBlocking>
        <Stack
          tokens={{ childrenGap: 16 }}
          styles={{
            root: {
              margin: '48px auto',
              maxWidth: 560,
              background: colors.cardBg,
              padding: 24,
              borderRadius: 4,
              border: `1px solid ${colors.border}`,
              boxShadow: `0 6.4px 14.4px ${hexToRgbaString(colors.bodyText, 0.13)}`,
            },
          }}
        >
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="xLarge" styles={{ root: { fontWeight: 600, color: colors.bodyText } }}>
              {title}
            </Text>
            <DefaultButton text="Fechar" onClick={onDismiss} />
          </Stack>
          {body}
        </Stack>
      </Modal>
    );
  }

  return (
    <Panel
      isOpen={isOpen}
      type={PanelType.medium}
      headerText={title}
      onDismiss={onDismiss}
      isBlocking
      closeButtonAriaLabel="Fechar"
      styles={{
        main: { background: colors.cardBg },
        header: { borderBottom: `1px solid ${colors.border}` },
        headerText: { color: colors.bodyText },
        content: { paddingTop: 16 },
      }}
    >
      {body}
    </Panel>
  );
};
