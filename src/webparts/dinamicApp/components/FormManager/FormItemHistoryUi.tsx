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
} from '@fluentui/react';
import type {
  IFormManagerActionLogConfig,
  TFormHistoryPresentationKind,
  TFormHistoryLayoutKind,
} from '../../core/config/types/formManager';
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
}

interface IAuditEntry {
  key: string;
  lineTitle: string;
  createdStr: string;
  who: string;
  html: string;
}

function authorDisplay(row: Record<string, unknown>): string {
  const a = row.Author;
  if (a && typeof a === 'object' && a !== null && 'Title' in a) {
    return String((a as { Title?: string }).Title ?? '');
  }
  return '';
}

function renderHtmlBlock(html: string, compact: boolean): React.ReactNode {
  return html ? (
    <div
      className="form-audit-log-html"
      style={{ fontSize: compact ? 12 : 14, color: '#323130', lineHeight: 1.45 }}
      dangerouslySetInnerHTML={{ __html: html }}
    />
  ) : (
    <Text
      variant="small"
      styles={{ root: { color: '#a19f9d', fontStyle: 'italic', fontSize: compact ? 11 : 12 } }}
    >
      (sem texto no campo de ação)
    </Text>
  );
}

function renderAuditEntries(entries: IAuditEntry[], layoutKind: TFormHistoryLayoutKind): React.ReactNode {
  const metaLine = (e: IAuditEntry): string => `${e.createdStr}${e.who ? ` · ${e.who}` : ''}`;

  if (layoutKind === 'compact') {
    return (
      <Stack tokens={{ childrenGap: 0 }}>
        {entries.map((e, i) => (
          <div
            key={e.key}
            style={{
              padding: '8px 0',
              borderBottom: i < entries.length - 1 ? '1px solid #edebe9' : undefined,
            }}
          >
            <Text variant="small" styles={{ root: { fontSize: 11, color: '#323130' } }}>
              <span style={{ fontWeight: 600 }}>{e.lineTitle}</span>
              <span style={{ color: '#605e5c' }}> · {metaLine(e)}</span>
            </Text>
            <div style={{ marginTop: e.html ? 6 : 4 }}>{renderHtmlBlock(e.html, true)}</div>
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
            background: '#e1dfdd',
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
                  background: '#0078d4',
                  border: '2px solid #fff',
                  boxShadow: '0 0 0 1px #c8c6c4',
                }}
              />
              <Text styles={{ root: { fontWeight: 600 } }}>{e.lineTitle}</Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {metaLine(e)}
              </Text>
              <div style={{ marginTop: 6 }}>{renderHtmlBlock(e.html, false)}</div>
            </div>
          ))}
        </Stack>
      </div>
    );
  }

  if (layoutKind === 'cards') {
    return (
      <Stack tokens={{ childrenGap: 12 }}>
        {entries.map((e) => (
          <div
            key={e.key}
            style={{
              padding: 16,
              borderRadius: 8,
              background: '#fff',
              boxShadow: '0 2px 8px rgba(0,0,0,0.08)',
              border: '1px solid #edebe9',
            }}
          >
            <Text styles={{ root: { fontWeight: 600 } }}>{e.lineTitle}</Text>
            <Text variant="small" styles={{ root: { color: '#605e5c', marginTop: 4 } }}>
              {metaLine(e)}
            </Text>
            <div style={{ marginTop: 10 }}>{renderHtmlBlock(e.html, false)}</div>
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
              border: '1px solid #edebe9',
              background: '#faf9f8',
            },
          }}
        >
          <Text styles={{ root: { fontWeight: 600 } }}>{e.lineTitle}</Text>
          <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
            {metaLine(e)}
          </Text>
          {renderHtmlBlock(e.html, false)}
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
}) => {
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
        'Indique na configuração do gestor (aba Lista de logs) a lista de registo, o campo multilinhas e o lookup de vínculo à lista principal.'
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
      const created = r.Created;
      const createdStr =
        typeof created === 'string'
          ? new Date(created).toLocaleString(undefined, { dateStyle: 'short', timeStyle: 'short' })
          : '—';
      const who = authorDisplay(r);
      const lineTitle = typeof r.Title === 'string' ? r.Title : String(r.Title ?? '—');
      out.push({ key, lineTitle, createdStr, who, html });
    }
    return out;
  }, [rows, resolvedActionField]);

  const body = (
    <Stack tokens={{ childrenGap: 12 }}>
      {subtitle && (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {subtitle}
        </Text>
      )}
      {err && <MessageBar messageBarType={MessageBarType.error}>{err}</MessageBar>}
      {loading && <Spinner label="A carregar registos de auditoria…" />}
      {!loading && !err && entries.length === 0 && (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Nenhum registo na lista de auditoria para este item (filtro pelo lookup configurado).
        </Text>
      )}
      {!loading && !err && entries.length > 0 && renderAuditEntries(entries, layoutKind)}
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
            border: '1px solid #edebe9',
            background: '#ffffff',
          },
        }}
      >
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
          <Text variant="mediumPlus" styles={{ root: { fontWeight: 600 } }}>
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
              background: '#ffffff',
              padding: 24,
              borderRadius: 4,
              boxShadow: '0 6.4px 14.4px rgba(0,0,0,.13)',
            },
          }}
        >
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="xLarge" styles={{ root: { fontWeight: 600 } }}>
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
    >
      {body}
    </Panel>
  );
};
