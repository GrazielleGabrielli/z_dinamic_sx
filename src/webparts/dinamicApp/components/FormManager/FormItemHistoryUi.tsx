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
  IconButton,
  useTheme,
  type ITheme,
} from '@fluentui/react';
import type {
  IFormManagerActionLogConfig,
  IFormManagerItemVersioningConfig,
  TFormHistoryPresentationKind,
  TFormHistoryLayoutKind,
  TFormCustomButtonPaletteSlot,
} from '../../core/config/types/formManager';
import { FORM_BUILTIN_HISTORY_BUTTON_ID } from '../../core/config/types/formManager';
import {
  hexToRgbaString,
  resolveActionLogPaletteAccentHex,
} from '../../core/formManager/formCustomButtonTheme';
import {
  parseActionLogButtonIdFromStoredHtml,
  stripActionLogMarkerFromStoredHtml,
} from '../../core/formManager/formActionLog';
import { fieldInternalsForItemVersionODataSelect } from '../../core/formManager/itemVersionSnapshotFields';
import { ItemsService, FieldsService, type IFieldMetadata } from '../../../../services';
import { FormManagerCollapseSection } from './FormManagerComponentsTab';

export interface IFormItemHistoryUiProps {
  actionLog: IFormManagerActionLogConfig | undefined;
  itemVersioning: IFormManagerItemVersioningConfig | undefined;
  primaryListTitle: string;
  listWebServerRelativeUrl?: string;
  sourceItemId: number;
  presentationKind: TFormHistoryPresentationKind;
  layoutKind?: TFormHistoryLayoutKind;
  isOpen: boolean;
  onDismiss: () => void;
  title: string;
  subtitle?: string;
  /** Cor de realce (passador); omitido = primária do tema Fluent. */
  accentColor?: string;
  /** Por entrada: bolinha / barra usam a cor do log por botão quando configurada. */
  logEntryPaletteContext?: {
    slotByButtonId: Record<string, TFormCustomButtonPaletteSlot>;
    customButtons: readonly { id: string; label: string }[];
    historyButtonLabel: string;
  };
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
  entryAccentHex: string;
}

function resolveLogButtonIdForPalette(
  rawHtml: string,
  actionLabel: string,
  buttons: readonly { id: string; label: string }[],
  historyButtonLabel: string
): string | undefined {
  const fromHtml = parseActionLogButtonIdFromStoredHtml(rawHtml);
  if (fromHtml) return fromHtml;
  const a = actionLabel.trim();
  const h = historyButtonLabel.trim() || 'Histórico';
  if (a === h) return FORM_BUILTIN_HISTORY_BUTTON_ID;
  for (let i = 0; i < buttons.length; i++) {
    const b = buttons[i];
    const disp = (b.label || b.id).trim();
    if (a === disp || a === b.id.trim()) return b.id;
  }
  return undefined;
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

function stripLegacySourceMetaLine(html: string): string {
  if (!html) return html;
  return html
    .replace(
      /<p\b[^>]*>\s*<em>\s*Lista de origem\s*<\/em>\s*:\s*[\s\S]*?·\s*<em>\s*Item\s*<\/em>\s*:\s*[\s\S]*?·\s*<em>\s*Modo\s*<\/em>\s*:\s*[\s\S]*?<\/p>/gi,
      ''
    )
    .trim();
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
    <Text variant="small" styles={{ root: { color: colors.bodySubtext, fontSize, marginTop: 4, fontStyle: 'italic' } }}>
      <em>Autor</em>: {e.who || '—'}
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
              padding: '8px 0 8px 10px',
              borderLeft: `3px solid ${e.entryAccentHex}`,
              borderBottom: i < entries.length - 1 ? `1px solid ${colors.border}` : undefined,
            }}
          >
            <Text variant="small" styles={{ root: { fontSize: 11 } }}>
              {entryHeadline(e, colors)}
            </Text>
            <div style={{ marginTop: e.html ? 6 : 4 }}>{renderHtmlBlock(e.html, true, colors)}</div>
            {entryAuthorLine(e, 11, colors)}
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
                  background: e.entryAccentHex,
                  border: `2px solid ${colors.cardBg}`,
                  boxShadow: `0 0 0 1px ${colors.borderStrong}`,
                }}
              />
              <Text variant="small" styles={{ root: { color: colors.bodyText } }}>
                {entryHeadline(e, colors)}
              </Text>
              <div style={{ marginTop: 6 }}>{renderHtmlBlock(e.html, false, colors)}</div>
              {entryAuthorLine(e, 12, colors)}
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
              borderLeft: `4px solid ${e.entryAccentHex}`,
            }}
          >
            <Text variant="small" styles={{ root: { color: colors.bodyText } }}>
              {entryHeadline(e, colors)}
            </Text>
            <div style={{ marginTop: 10 }}>{renderHtmlBlock(e.html, false, colors)}</div>
            {entryAuthorLine(e, 12, colors)}
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
              borderLeft: `4px solid ${e.entryAccentHex}`,
              background: colors.listRowBg,
            },
          }}
        >
          <Text variant="small" styles={{ root: { color: colors.bodyText } }}>
            {entryHeadline(e, colors)}
          </Text>
          {renderHtmlBlock(e.html, false, colors)}
          {entryAuthorLine(e, 12, colors)}
        </Stack>
      ))}
    </Stack>
  );
}

const HISTORY_SECTION_IDS = {
  audit: 'histAuditLog',
  versions: 'histSpVersions',
} as const;

function formatVersionFieldValue(v: unknown): string {
  if (v == null || v === '') return '—';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toLocaleString(undefined, { dateStyle: 'short', timeStyle: 'short' });
  if (typeof v === 'object' && v !== null && 'Title' in (v as Record<string, unknown>)) {
    return String((v as { Title?: string }).Title ?? '').trim() || '—';
  }
  if (typeof v === 'object' && v !== null && 'Label' in (v as Record<string, unknown>)) {
    return String((v as { Label?: string }).Label ?? '').trim() || '—';
  }
  return '[objeto]';
}

function renderVersionSnapshotTable(
  snap: Record<string, unknown>,
  colors: IHistoryUiColors
): React.ReactNode {
  const keys = Object.keys(snap)
    .filter((k) => !k.startsWith('odata') && k !== '__metadata' && k !== 'ID')
    .sort((a, b) => a.localeCompare(b));
  return (
    <Stack tokens={{ childrenGap: 6 }}>
      {keys.map((k) => (
        <Stack key={k} horizontal tokens={{ childrenGap: 8 }} verticalAlign="start" wrap>
          <Text
            variant="small"
            styles={{
              root: {
                color: colors.bodySubtext,
                fontFamily: 'monospace',
                minWidth: 120,
                flexShrink: 0,
              },
            }}
          >
            {k}
          </Text>
          <Text variant="small" styles={{ root: { color: colors.bodyText, flex: 1 } }}>
            {formatVersionFieldValue(snap[k])}
          </Text>
        </Stack>
      ))}
    </Stack>
  );
}

export const FormItemHistoryUi: React.FC<IFormItemHistoryUiProps> = ({
  actionLog,
  itemVersioning,
  primaryListTitle,
  listWebServerRelativeUrl,
  sourceItemId,
  presentationKind,
  layoutKind = 'list',
  isOpen,
  onDismiss,
  title,
  subtitle,
  accentColor,
  logEntryPaletteContext,
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
  const [versionRows, setVersionRows] = useState<
    { versionLabel: string; versionId: number; created?: string; isCurrentVersion?: boolean }[]
  >([]);
  const [versionLoading, setVersionLoading] = useState(false);
  const [versionErr, setVersionErr] = useState<string | undefined>(undefined);
  const [expandedVersionId, setExpandedVersionId] = useState<number | null>(null);
  const [verSnap, setVerSnap] = useState<Record<string, unknown> | null>(null);
  const [verSnapLoading, setVerSnapLoading] = useState(false);
  const [verSnapErr, setVerSnapErr] = useState<string | undefined>(undefined);
  const [versionSnapshotPrimaryMeta, setVersionSnapshotPrimaryMeta] = useState<IFieldMetadata[]>([]);
  const [openHistSections, setOpenHistSections] = useState<Record<string, boolean>>({
    [HISTORY_SECTION_IDS.audit]: true,
    [HISTORY_SECTION_IDS.versions]: true,
  });
  const toggleHistSection = (id: string): void => {
    setOpenHistSections((prev) => ({ ...prev, [id]: !prev[id] }));
  };
  const histSectionIsOpen = (id: string): boolean => openHistSections[id] !== false;

  const showAuditBlock = useMemo(() => {
    const logList = actionLog?.listTitle?.trim();
    const actionField = actionLog?.actionFieldInternalName?.trim();
    const linkField = actionLog?.sourceListLookupFieldInternalName?.trim();
    return !!(logList && actionField && linkField);
  }, [actionLog?.listTitle, actionLog?.actionFieldInternalName, actionLog?.sourceListLookupFieldInternalName]);

  const showVersionsBlock = useMemo(
    () =>
      itemVersioning?.showInHistoryPanel === true &&
      !!(primaryListTitle ?? '').trim() &&
      sourceItemId >= 1,
    [itemVersioning?.showInHistoryPanel, primaryListTitle, sourceItemId]
  );

  useEffect(() => {
    if (!isOpen || !showVersionsBlock || !primaryListTitle.trim()) {
      setVersionSnapshotPrimaryMeta([]);
      return;
    }
    const t = primaryListTitle.trim();
    void fieldsService
      .getVisibleFields(t, listWebServerRelativeUrl)
      .then((m) => setVersionSnapshotPrimaryMeta(m))
      .catch(() => setVersionSnapshotPrimaryMeta([]));
  }, [isOpen, showVersionsBlock, primaryListTitle, listWebServerRelativeUrl, fieldsService]);

  const versionSnapshotFieldInternalsResolved = useMemo(() => {
    const cf = itemVersioning?.snapshotFieldsInternalNames;
    if (cf && cf.length) {
      return cf
        .map((x) => String(x).trim())
        .filter((x) => /^[A-Za-z0-9_]+$/.test(x))
        .slice(0, 48);
    }
    return fieldInternalsForItemVersionODataSelect(versionSnapshotPrimaryMeta);
  }, [itemVersioning?.snapshotFieldsInternalNames, versionSnapshotPrimaryMeta]);

  const versionSnapshotResolvedKey = useMemo(
    () => versionSnapshotFieldInternalsResolved.join('\u0001'),
    [versionSnapshotFieldInternalsResolved]
  );

  useEffect(() => {
    if (!isOpen || !sourceItemId || sourceItemId < 1) return;
    if (!showAuditBlock) {
      setErr(undefined);
      setRows([]);
      setResolvedActionField('');
      setLoading(false);
      return;
    }
    const logList = actionLog?.listTitle?.trim() as string;
    const actionField = actionLog?.actionFieldInternalName?.trim() as string;
    const linkField = actionLog?.sourceListLookupFieldInternalName?.trim() as string;
    setErr(undefined);
    setLoading(true);
    setResolvedActionField(actionField);
    const filter = `${linkField}Id eq ${sourceItemId}`;
    void (async (): Promise<void> => {
      try {
        const meta = await fieldsService.getVisibleFields(logList);
        const select: string[] = ['Id', 'Title', 'Created', actionField, 'Author/Id', 'Author/Title', 'Author/EMail'];
        const data = await itemsService.getItems<Record<string, unknown>>(logList, {
          filter,
          orderBy: { field: 'Created', ascending: false },
          top: 200,
          fieldMetadata: meta,
          select,
          expand: ['Author'],
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
    showAuditBlock,
    actionLog?.listTitle,
    actionLog?.actionFieldInternalName,
    actionLog?.sourceListLookupFieldInternalName,
    fieldsService,
    itemsService,
  ]);

  useEffect(() => {
    if (!isOpen || !sourceItemId || sourceItemId < 1 || !showVersionsBlock) {
      setVersionRows([]);
      setVersionErr(undefined);
      setVersionLoading(false);
      setExpandedVersionId(null);
      setVerSnap(null);
      setVerSnapErr(undefined);
      return;
    }
    const lt = primaryListTitle.trim();
    setVersionLoading(true);
    setVersionErr(undefined);
    void (async (): Promise<void> => {
      try {
        const data = await itemsService.getItemVersions(lt, sourceItemId, listWebServerRelativeUrl);
        setVersionRows(Array.isArray(data) ? data : []);
      } catch (e) {
        setVersionErr(e instanceof Error ? e.message : String(e));
        setVersionRows([]);
      } finally {
        setVersionLoading(false);
      }
    })();
  }, [
    isOpen,
    sourceItemId,
    showVersionsBlock,
    primaryListTitle,
    listWebServerRelativeUrl,
    itemsService,
  ]);

  useEffect(() => {
    if (!expandedVersionId || !showVersionsBlock || !primaryListTitle.trim()) {
      setVerSnap(null);
      setVerSnapErr(undefined);
      setVerSnapLoading(false);
      return;
    }
    const lt = primaryListTitle.trim();
    const fieldInternalNames = versionSnapshotFieldInternalsResolved;
    setVerSnap(null);
    setVerSnapLoading(true);
    setVerSnapErr(undefined);
    void (async (): Promise<void> => {
      try {
        const snap = await itemsService.getItemVersionSnapshot(lt, sourceItemId, expandedVersionId, {
          webServerRelativeUrl: listWebServerRelativeUrl,
          fieldInternalNames,
        });
        setVerSnap(snap);
      } catch (e) {
        setVerSnapErr(e instanceof Error ? e.message : String(e));
        setVerSnap(null);
      } finally {
        setVerSnapLoading(false);
      }
    })();
  }, [
    expandedVersionId,
    showVersionsBlock,
    primaryListTitle,
    sourceItemId,
    listWebServerRelativeUrl,
    versionSnapshotResolvedKey,
    itemsService,
    versionSnapshotFieldInternalsResolved,
  ]);

  const entries: IAuditEntry[] = useMemo(() => {
    const out: IAuditEntry[] = [];
    const defaultAccent = colors.accent;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const id = r.Id;
      const key =
        typeof id === 'number' || typeof id === 'string' ? String(id) : `r-${i}`;
      const rawHtml = resolvedActionField ? r[resolvedActionField] : undefined;
      const htmlRaw =
        typeof rawHtml === 'string'
          ? rawHtml
          : rawHtml !== undefined && rawHtml !== null
            ? String(rawHtml)
            : '';
      const createdStr = formatCreatedValue(r.Created);
      const who = authorDisplay(r);
      const lineTitle = typeof r.Title === 'string' ? r.Title : String(r.Title ?? '—');
      const actionLabel = actionLabelFromItemTitle(lineTitle);
      let entryAccentHex = defaultAccent;
      if (logEntryPaletteContext) {
        const btnId = resolveLogButtonIdForPalette(
          htmlRaw,
          actionLabel,
          logEntryPaletteContext.customButtons,
          logEntryPaletteContext.historyButtonLabel
        );
        const slot =
          (btnId ? logEntryPaletteContext.slotByButtonId[btnId] : undefined) ?? 'themePrimary';
        entryAccentHex = resolveActionLogPaletteAccentHex(theme, slot);
      }
      const html = stripLegacySourceMetaLine(stripActionLogMarkerFromStoredHtml(htmlRaw));
      out.push({ key, actionLabel, createdStr, who, html, entryAccentHex });
    }
    return out;
  }, [rows, resolvedActionField, logEntryPaletteContext, theme, colors.accent]);

  const panelOrphan = isOpen && !showAuditBlock && !showVersionsBlock;

  const body = (
    <Stack tokens={{ childrenGap: 12 }}>
      {subtitle && (
        <Text variant="small" styles={{ root: { color: colors.bodySubtext } }}>
          {subtitle}
        </Text>
      )}
      {panelOrphan && (
        <MessageBar messageBarType={MessageBarType.warning}>
          Não há fontes para este painel. Na aba «Auditoria e versões» configure a lista de logs completa ou ative o
          versionamento do item no painel.
        </MessageBar>
      )}
      {showAuditBlock && (
        <FormManagerCollapseSection
          title="Lista de logs"
          isOpen={histSectionIsOpen(HISTORY_SECTION_IDS.audit)}
          onToggle={() => toggleHistSection(HISTORY_SECTION_IDS.audit)}
        >
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
        </FormManagerCollapseSection>
      )}
      {showVersionsBlock && (
        <FormManagerCollapseSection
          title="Versionamento do item (SharePoint)"
          isOpen={histSectionIsOpen(HISTORY_SECTION_IDS.versions)}
          onToggle={() => toggleHistSection(HISTORY_SECTION_IDS.versions)}
        >
          {versionErr && <MessageBar messageBarType={MessageBarType.error}>{versionErr}</MessageBar>}
          {versionLoading && (
            <Spinner
              label="A carregar versões do item…"
              styles={{ circle: { borderTopColor: colors.accent } }}
            />
          )}
          {!versionLoading && !versionErr && versionRows.length === 0 && (
            <Text variant="small" styles={{ root: { color: colors.bodySubtext } }}>
              Não existem versões (ou o versionamento está desativado na lista principal).
            </Text>
          )}
          {!versionLoading &&
            !versionErr &&
            versionRows.map((v) => (
              <div
                key={v.versionId}
                style={{
                  border: `1px solid ${colors.border}`,
                  borderRadius: 4,
                  overflow: 'hidden',
                  marginBottom: 8,
                }}
              >
                <Stack
                  horizontal
                  verticalAlign="center"
                  tokens={{ childrenGap: 4 }}
                  styles={{
                    root: {
                      padding: '6px 8px',
                      cursor: 'pointer',
                      background: colors.listRowBg,
                    },
                  }}
                  onClick={() =>
                    setExpandedVersionId((cur) => (cur === v.versionId ? null : v.versionId))
                  }
                >
                  <IconButton
                    iconProps={{
                      iconName: expandedVersionId === v.versionId ? 'ChevronDown' : 'ChevronRight',
                    }}
                    title={expandedVersionId === v.versionId ? 'Recolher' : 'Expandir'}
                    aria-expanded={expandedVersionId === v.versionId}
                    onClick={(e) => {
                      e.stopPropagation();
                      setExpandedVersionId((cur) => (cur === v.versionId ? null : v.versionId));
                    }}
                  />
                  <Stack grow styles={{ root: { minWidth: 0 } }}>
                    <Text variant="small" styles={{ root: { color: colors.bodyText } }}>
                      <span style={{ fontWeight: 600 }}>Versão {v.versionLabel}</span>
                      <span style={{ color: colors.bodySubtext, fontWeight: 400 }}>
                        {' '}
                        · {formatCreatedValue(v.created)}
                        {v.isCurrentVersion ? ' · atual' : ''}
                      </span>
                    </Text>
                  </Stack>
                </Stack>
                {expandedVersionId === v.versionId && (
                  <div
                    style={{
                      padding: 12,
                      borderTop: `1px solid ${colors.border}`,
                      background: colors.cardBg,
                    }}
                  >
                    {verSnapLoading && (
                      <Spinner
                        label="A carregar campos desta versão…"
                        styles={{ circle: { borderTopColor: colors.accent } }}
                      />
                    )}
                    {verSnapErr && <MessageBar messageBarType={MessageBarType.error}>{verSnapErr}</MessageBar>}
                    {!verSnapLoading && !verSnapErr && verSnap && renderVersionSnapshotTable(verSnap, colors)}
                  </div>
                )}
              </div>
            ))}
        </FormManagerCollapseSection>
      )}
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
