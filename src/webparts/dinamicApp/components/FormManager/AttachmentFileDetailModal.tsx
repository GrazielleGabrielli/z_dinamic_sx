import * as React from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { Modal, Stack, Text, PrimaryButton, DefaultButton, Icon } from '@fluentui/react';
import { attachmentFileKindIconName } from './attachmentFileKindIcon';

export type IAttachmentFileDetailTarget =
  | { kind: 'server'; fileName: string; fileUrl?: string; fileRef?: string }
  | { kind: 'pending'; file: File };

function formatKb(n: number): string {
  return (n / 1024).toFixed(1);
}

function isProbablyImageServer(fileName: string, url?: string): boolean {
  const ext = (fileName.toLowerCase().split('.').pop() || '').trim();
  if (['png', 'jpg', 'jpeg', 'gif', 'webp', 'bmp', 'svg'].indexOf(ext) !== -1) return true;
  if (url && /\.(png|jpe?g|gif|webp|bmp|svg)(\?|#|$)/i.test(url)) return true;
  return false;
}

function isImageFile(f: File): boolean {
  return (f.type || '').toLowerCase().startsWith('image/');
}

function DetailRow(props: { label: string; value: string; monospace?: boolean }): JSX.Element {
  return (
    <Stack tokens={{ childrenGap: 2 }}>
      <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
        {props.label}
      </Text>
      <Text
        variant="small"
        styles={{
          root: {
            color: '#605e5c',
            wordBreak: 'break-all',
            ...(props.monospace ? { fontFamily: 'monospace, Consolas, monospace' } : {}),
          },
        }}
      >
        {props.value}
      </Text>
    </Stack>
  );
}

export interface IAttachmentFileDetailModalProps {
  isOpen: boolean;
  onDismiss: () => void;
  target: IAttachmentFileDetailTarget | null;
}

export const AttachmentFileDetailModal: React.FC<IAttachmentFileDetailModalProps> = ({
  isOpen,
  onDismiss,
  target,
}) => {
  const [objUrl, setObjUrl] = useState<string | undefined>(undefined);
  const [urlCopied, setUrlCopied] = useState(false);

  useEffect(() => {
    if (!isOpen) setUrlCopied(false);
  }, [isOpen]);

  useEffect(() => {
    setUrlCopied(false);
  }, [target]);

  useEffect(() => {
    if (!isOpen || !target || target.kind !== 'pending') {
      setObjUrl(undefined);
      return;
    }
    if (!isImageFile(target.file)) {
      setObjUrl(undefined);
      return;
    }
    const u = URL.createObjectURL(target.file);
    setObjUrl(u);
    return () => URL.revokeObjectURL(u);
  }, [isOpen, target]);

  const title = useMemo(() => {
    if (!target) return 'Anexo';
    return target.kind === 'server' ? target.fileName : target.file.name;
  }, [target]);

  const iconName = useMemo(() => {
    if (!target) return 'Page';
    return target.kind === 'server'
      ? attachmentFileKindIconName(target.fileName)
      : attachmentFileKindIconName(target.file.name);
  }, [target]);

  const openUrl =
    target?.kind === 'server' && target.fileUrl?.trim()
      ? (): void => {
          window.open(target.fileUrl, '_blank', 'noopener,noreferrer');
        }
      : undefined;

  const serverUrlTrimmed =
    target?.kind === 'server' && target.fileUrl?.trim() ? target.fileUrl.trim() : '';

  const copyServerUrl = useCallback(async (): Promise<void> => {
    if (!serverUrlTrimmed) return;
    try {
      await navigator.clipboard.writeText(serverUrlTrimmed);
      setUrlCopied(true);
      window.setTimeout(() => setUrlCopied(false), 2000);
    } catch {
      try {
        const ta = document.createElement('textarea');
        ta.value = serverUrlTrimmed;
        ta.setAttribute('readonly', '');
        ta.style.position = 'fixed';
        ta.style.left = '-9999px';
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        setUrlCopied(true);
        window.setTimeout(() => setUrlCopied(false), 2000);
      } catch {
        //
      }
    }
  }, [serverUrlTrimmed]);

  return (
    <Modal
      isOpen={isOpen && Boolean(target)}
      onDismiss={onDismiss}
      isBlocking
      styles={{ main: { maxWidth: 520, width: 'min(90vw, 520px)', margin: 'auto' } }}
    >
      {target ? (
        <div style={{ padding: 24 }}>
          <Stack
            horizontal
            verticalAlign="start"
            tokens={{ childrenGap: 12 }}
            styles={{ root: { marginBottom: 16 } }}
          >
            <Icon iconName={iconName} styles={{ root: { fontSize: 28, color: '#0078d4', flexShrink: 0 } }} />
            <Stack tokens={{ childrenGap: 4 }} styles={{ root: { minWidth: 0, flex: 1 } }}>
              <Text variant="xLarge" styles={{ root: { fontWeight: 600, wordBreak: 'break-word' } }}>
                {title}
              </Text>
              <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                {target.kind === 'pending'
                  ? 'Ainda não enviado — pré-visualização local'
                  : 'Ficheiro no servidor'}
              </Text>
            </Stack>
          </Stack>

          {target.kind === 'pending' ? (
            <>
              {objUrl ? (
                <img
                  src={objUrl}
                  alt=""
                  style={{
                    maxWidth: '100%',
                    maxHeight: 240,
                    objectFit: 'contain',
                    borderRadius: 6,
                    marginBottom: 16,
                    display: 'block',
                    background: '#f3f2f1',
                  }}
                />
              ) : null}
              <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginBottom: 20 } }}>
                <DetailRow
                  label="Tamanho"
                  value={`${formatKb(target.file.size)} KB (${target.file.size} bytes)`}
                />
                <DetailRow label="Tipo MIME" value={target.file.type?.trim() ? target.file.type : '—'} />
                <DetailRow
                  label="Última modificação (local)"
                  value={
                    target.file.lastModified
                      ? new Date(target.file.lastModified).toLocaleString()
                      : '—'
                  }
                />
              </Stack>
            </>
          ) : (
            <>
              {target.fileUrl && isProbablyImageServer(target.fileName, target.fileUrl) ? (
                <img
                  src={target.fileUrl}
                  alt=""
                  style={{
                    maxWidth: '100%',
                    maxHeight: 240,
                    objectFit: 'contain',
                    borderRadius: 6,
                    marginBottom: 16,
                    display: 'block',
                    background: '#f3f2f1',
                  }}
                />
              ) : null}
              <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginBottom: 20 } }}>
                {serverUrlTrimmed ? (
                  <Stack tokens={{ childrenGap: 6 }}>
                    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} wrap>
                      <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                        URL
                      </Text>
                      <DefaultButton
                        text={urlCopied ? 'Copiado' : 'Copiar'}
                        iconProps={{ iconName: 'Copy' }}
                        onClick={() => void copyServerUrl()}
                      />
                    </Stack>
                    <Text
                      variant="small"
                      styles={{
                        root: {
                          color: '#605e5c',
                          wordBreak: 'break-all',
                          fontFamily: 'monospace, Consolas, monospace',
                        },
                      }}
                    >
                      {serverUrlTrimmed}
                    </Text>
                  </Stack>
                ) : (
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Sem URL direta para este ficheiro.
                  </Text>
                )}
              </Stack>
            </>
          )}

          <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 8 }} wrap>
            {openUrl !== undefined ? (
              <PrimaryButton text="Abrir em nova janela" onClick={openUrl} />
            ) : null}
            <DefaultButton text="Fechar" onClick={onDismiss} />
          </Stack>
        </div>
      ) : null}
    </Modal>
  );
};
