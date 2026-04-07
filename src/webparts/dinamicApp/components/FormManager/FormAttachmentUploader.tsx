import * as React from 'react';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { Stack, Text, Label, IconButton, Icon, PrimaryButton, DefaultButton } from '@fluentui/react';
import type {
  TFormAttachmentUploadLayoutKind,
  TFormAttachmentFilePreviewKind,
} from '../../core/config/types/formManager';
import { attachmentFileKindIconName } from './attachmentFileKindIcon';

export interface IFormAttachmentUploaderProps {
  files: File[];
  onFilesChange: (files: File[]) => void;
  disabled: boolean;
  label: string;
  description?: string;
  errorMessage?: string;
  required?: boolean;
  requiredEmptyHighlight?: boolean;
  layout?: TFormAttachmentUploadLayoutKind;
  filePreview?: TFormAttachmentFilePreviewKind;
  /** Extensões sem ponto; vazio/omitido = aceitar qualquer ficheiro. */
  allowedFileExtensions?: string[];
}

const accent = '#0078d4';
const borderErr = '#a4262c';
const bgErr = '#fef6f6';

function fileExtensionLower(name: string): string {
  const n = name.trim();
  const d = n.lastIndexOf('.');
  if (d < 0 || d >= n.length - 1) return '';
  return n.slice(d + 1).toLowerCase();
}

function mergeFilesFiltered(
  prev: File[],
  added: FileList | null,
  allowed: string[] | undefined
): { next: File[]; rejected: number } {
  if (!added || added.length === 0) return { next: prev.slice(), rejected: 0 };
  if (!allowed || allowed.length === 0) {
    const next = prev.slice();
    for (let i = 0; i < added.length; i++) next.push(added[i]);
    return { next, rejected: 0 };
  }
  const allow = new Set(
    allowed.map((x) => String(x).trim().replace(/^\./, '').toLowerCase()).filter(Boolean)
  );
  const next = prev.slice();
  let rejected = 0;
  for (let i = 0; i < added.length; i++) {
    const f = added[i];
    const ext = fileExtensionLower(f.name);
    if (ext && allow.has(ext)) next.push(f);
    else rejected++;
  }
  return { next, rejected };
}

function formatKb(n: number): string {
  return (n / 1024).toFixed(1);
}

function isImageFile(f: File): boolean {
  return (f.type || '').toLowerCase().startsWith('image/');
}

function fileKindIconName(f: File): string {
  const t = (f.type || '').toLowerCase();
  if (t.startsWith('image/')) return 'FileImage';
  return attachmentFileKindIconName(f.name);
}

export const FormAttachmentUploader: React.FC<IFormAttachmentUploaderProps> = ({
  files,
  onFilesChange,
  disabled,
  label,
  description,
  errorMessage,
  required,
  requiredEmptyHighlight,
  layout = 'default',
  filePreview: filePreviewProp,
  allowedFileExtensions,
}) => {
  const preview: TFormAttachmentFilePreviewKind = filePreviewProp ?? 'nameAndSize';
  const useCompactChips =
    layout === 'compact' && (preview === 'nameOnly' || preview === 'nameAndSize');
  const inputRef = useRef<HTMLInputElement>(null);
  const [dragOver, setDragOver] = useState(false);
  const [pickRejectHint, setPickRejectHint] = useState('');

  const inputAccept = useMemo(() => {
    if (!allowedFileExtensions || allowedFileExtensions.length === 0) return undefined;
    return allowedFileExtensions
      .map((x) => {
        const e = String(x).trim().replace(/^\./, '').toLowerCase();
        return e ? `.${e}` : '';
      })
      .filter(Boolean)
      .join(',');
  }, [allowedFileExtensions]);

  const imageUrls = useMemo(() => {
    return files.map((f) => (isImageFile(f) ? URL.createObjectURL(f) : undefined));
  }, [files]);

  useEffect(() => {
    return () => {
      for (let i = 0; i < imageUrls.length; i++) {
        const u = imageUrls[i];
        if (u) URL.revokeObjectURL(u);
      }
    };
  }, [imageUrls]);

  const openPicker = useCallback(() => {
    inputRef.current?.click();
  }, []);

  const onInputChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>): void => {
      const fl = e.target.files;
      const { next, rejected } = mergeFilesFiltered(files, fl, allowedFileExtensions);
      if (rejected > 0) {
        setPickRejectHint(
          rejected === 1
            ? 'Um ficheiro foi ignorado (extensão não permitida).'
            : `${rejected} ficheiros ignorados (extensão não permitida).`
        );
      } else {
        setPickRejectHint('');
      }
      onFilesChange(next);
      e.target.value = '';
    },
    [files, onFilesChange, allowedFileExtensions]
  );

  const addFromDataTransfer = useCallback(
    (dt: DataTransfer | null): void => {
      if (!dt?.files?.length) return;
      const { next, rejected } = mergeFilesFiltered(files, dt.files, allowedFileExtensions);
      if (rejected > 0) {
        setPickRejectHint(
          rejected === 1
            ? 'Um ficheiro foi ignorado (extensão não permitida).'
            : `${rejected} ficheiros ignorados (extensão não permitida).`
        );
      } else {
        setPickRejectHint('');
      }
      onFilesChange(next);
    },
    [files, onFilesChange, allowedFileExtensions]
  );

  const removeAt = useCallback(
    (idx: number): void => {
      onFilesChange(files.filter((_, i) => i !== idx));
    },
    [files, onFilesChange]
  );

  const reqWrapStyle: React.CSSProperties | undefined =
    requiredEmptyHighlight && !disabled
      ? { padding: 8, borderRadius: 8, border: `1px solid ${borderErr}`, background: bgErr }
      : undefined;

  const hiddenInput = (
    <input
      ref={inputRef}
      type="file"
      multiple
      accept={inputAccept}
      onChange={onInputChange}
      style={{ display: 'none' }}
      aria-hidden
      tabIndex={-1}
    />
  );

  const showSize = preview !== 'nameOnly';

  const renderThumbBox = (
    f: File,
    idx: number,
    px: number,
    rounded: number
  ): React.ReactNode => {
    const url = imageUrls[idx];
    if (url) {
      return (
        <img
          src={url}
          alt=""
          style={{
            width: px,
            height: px,
            objectFit: 'cover',
            borderRadius: rounded,
            flexShrink: 0,
            display: 'block',
          }}
        />
      );
    }
    return (
      <div
        style={{
          width: px,
          height: px,
          borderRadius: rounded,
          background: '#edebe9',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          flexShrink: 0,
        }}
      >
        <Icon iconName={fileKindIconName(f)} styles={{ root: { fontSize: Math.max(14, px * 0.45), color: '#605e5c' } }} />
      </div>
    );
  };

  const renderLargePreview = (f: File, idx: number): React.ReactNode => {
    const url = imageUrls[idx];
    if (url) {
      return (
        <img
          src={url}
          alt=""
          style={{
            width: '100%',
            maxHeight: 140,
            objectFit: 'contain',
            borderRadius: 6,
            background: '#f3f2f1',
            display: 'block',
          }}
        />
      );
    }
    return (
      <div
        style={{
          minHeight: 100,
          borderRadius: 6,
          background: '#edebe9',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
        }}
      >
        <Icon iconName={fileKindIconName(f)} styles={{ root: { fontSize: 48, color: '#605e5c' } }} />
      </div>
    );
  };

  const renderFileRows = (compactChips: boolean): React.ReactNode => {
    if (files.length === 0) return null;
    const p: TFormAttachmentFilePreviewKind =
      compactChips && preview === 'thumbnailLarge' ? 'thumbnailAndName' : preview;

    const removeBtn = (f: File, idx: number, small?: boolean) =>
      !disabled && (
        <IconButton
          iconProps={{
            iconName: 'Cancel',
            styles: small ? { root: { height: 24, width: 24 } } : undefined,
          }}
          title="Remover"
          ariaLabel={`Remover ${f.name}`}
          onClick={() => removeAt(idx)}
        />
      );

    if (compactChips) {
      return (
        <Stack horizontal wrap tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 4 } }}>
          {files.map((f, idx) => {
            const sz = showSize ? (
              <span style={{ color: '#605e5c', marginLeft: 6 }}>{formatKb(f.size)} KB</span>
            ) : null;
            let inner: React.ReactNode;
            if (p === 'nameOnly' || p === 'nameAndSize') {
              inner = (
                <Text
                  variant="small"
                  styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }}
                >
                  {f.name}
                  {sz}
                </Text>
              );
            } else if (p === 'iconAndName') {
              inner = (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 6 }} styles={{ root: { minWidth: 0 } }}>
                  <Icon iconName={fileKindIconName(f)} styles={{ root: { fontSize: 16, color: '#605e5c', flexShrink: 0 } }} />
                  <Text
                    variant="small"
                    styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }}
                  >
                    {f.name}
                    {sz}
                  </Text>
                </Stack>
              );
            } else {
              inner = (
                <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }} styles={{ root: { minWidth: 0 } }}>
                  {renderThumbBox(f, idx, 24, 4)}
                  <Text
                    variant="small"
                    styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }}
                  >
                    {f.name}
                    {sz}
                  </Text>
                </Stack>
              );
            }
            return (
              <Stack
                key={`${f.name}_${idx}_${f.size}`}
                horizontal
                verticalAlign="center"
                tokens={{ childrenGap: 4 }}
                styles={{
                  root: {
                    padding: '6px 12px',
                    background: '#f3f2f1',
                    borderRadius: 16,
                    border: '1px solid #edebe9',
                    maxWidth: '100%',
                  },
                }}
              >
                {inner}
                {removeBtn(f, idx, true)}
              </Stack>
            );
          })}
        </Stack>
      );
    }

    if (p === 'thumbnailLarge') {
      return (
        <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginTop: 4 } }}>
          {files.map((f, idx) => (
            <Stack
              key={`${f.name}_${idx}_${f.size}`}
              tokens={{ childrenGap: 8 }}
              styles={{
                root: {
                  padding: 12,
                  background: '#faf9f8',
                  borderRadius: 8,
                  border: '1px solid #edebe9',
                },
              }}
            >
              {renderLargePreview(f, idx)}
              <Stack horizontal verticalAlign="center" horizontalAlign="space-between" tokens={{ childrenGap: 8 }}>
                <Stack tokens={{ childrenGap: 2 }} styles={{ root: { minWidth: 0, flex: 1 } }}>
                  <Text variant="small" styles={{ root: { fontWeight: 600, overflow: 'hidden', textOverflow: 'ellipsis' } }}>
                    {f.name}
                  </Text>
                  {showSize && (
                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                      {formatKb(f.size)} KB
                    </Text>
                  )}
                </Stack>
                {removeBtn(f, idx)}
              </Stack>
            </Stack>
          ))}
        </Stack>
      );
    }

    return (
      <Stack tokens={{ childrenGap: 4 }} styles={{ root: { marginTop: 4 } }}>
        {files.map((f, idx) => {
          const sz = showSize ? (
            <span style={{ color: '#605e5c', marginLeft: 8 }}>{formatKb(f.size)} KB</span>
          ) : null;
          let row: React.ReactNode;
          if (p === 'nameOnly' || p === 'nameAndSize') {
            row = (
              <Text variant="small" styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis' } }}>
                {f.name}
                {sz}
              </Text>
            );
          } else if (p === 'iconAndName') {
            row = (
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} styles={{ root: { minWidth: 0, flex: 1 } }}>
                <Icon iconName={fileKindIconName(f)} styles={{ root: { fontSize: 20, color: accent, flexShrink: 0 } }} />
                <Text variant="small" styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis' } }}>
                  {f.name}
                  {sz}
                </Text>
              </Stack>
            );
          } else {
            row = (
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} styles={{ root: { minWidth: 0, flex: 1 } }}>
                {renderThumbBox(f, idx, 44, 6)}
                <Text variant="small" styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis' } }}>
                  {f.name}
                  {sz}
                </Text>
              </Stack>
            );
          }
          return (
            <Stack
              key={`${f.name}_${idx}_${f.size}`}
              horizontal
              verticalAlign="center"
              horizontalAlign="space-between"
              tokens={{ childrenGap: 8 }}
              styles={{
                root: {
                  padding: '8px 12px',
                  background: '#faf9f8',
                  borderRadius: 6,
                  border: '1px solid #edebe9',
                },
              }}
            >
              {row}
              {removeBtn(f, idx)}
            </Stack>
          );
        })}
      </Stack>
    );
  };

  const dropHandlers = (active: boolean) =>
    !disabled && active
      ? {
          onDragOver: (e: React.DragEvent) => {
            e.preventDefault();
            e.stopPropagation();
            e.dataTransfer.dropEffect = 'copy';
            setDragOver(true);
          },
          onDragLeave: (e: React.DragEvent) => {
            e.preventDefault();
            setDragOver(false);
          },
          onDrop: (e: React.DragEvent) => {
            e.preventDefault();
            e.stopPropagation();
            setDragOver(false);
            addFromDataTransfer(e.dataTransfer);
          },
        }
      : {};

  const header = (
    <>
      <Label required={required}>{label}</Label>
      {description && (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {description}
        </Text>
      )}
    </>
  );

  const footerErr = (
    <>
      {errorMessage && (
        <Text variant="small" styles={{ root: { color: borderErr } }}>
          {errorMessage}
        </Text>
      )}
      {pickRejectHint && (
        <Text variant="small" styles={{ root: { color: borderErr } }}>
          {pickRejectHint}
        </Text>
      )}
    </>
  );

  if (layout === 'default') {
    return (
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
        {header}
        {!disabled && (
          <div style={reqWrapStyle}>
            <input
              type="file"
              multiple
              accept={inputAccept}
              onChange={onInputChange}
              style={{ maxWidth: '100%' }}
            />
          </div>
        )}
        {renderFileRows(false)}
        {footerErr}
      </Stack>
    );
  }

  if (layout === 'ribbon') {
    return (
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
        {header}
        {!disabled && (
          <div style={reqWrapStyle}>
            {hiddenInput}
            <div
              style={{
                borderTop: `4px solid ${accent}`,
                borderRadius: '0 0 8px 8px',
                boxShadow: '0 1.6px 3.6px rgba(0,0,0,0.1)',
                padding: 14,
                background: '#fff',
                border: '1px solid #edebe9',
                borderTopColor: accent,
              }}
            >
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} wrap>
                <PrimaryButton text="Escolher ficheiros" onClick={openPicker} />
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  ou largue ficheiros nesta área
                </Text>
              </Stack>
              <div
                {...dropHandlers(true)}
                style={{
                  marginTop: 10,
                  minHeight: 48,
                  borderRadius: 6,
                  border: `2px dashed ${dragOver ? accent : '#c8c6c4'}`,
                  background: dragOver ? 'rgba(0, 120, 212, 0.06)' : '#faf9f8',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                }}
              >
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Área de largar
                </Text>
              </div>
            </div>
          </div>
        )}
        {renderFileRows(false)}
        {footerErr}
      </Stack>
    );
  }

  if (layout === 'compact') {
    return (
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
        {header}
        {!disabled && (
          <div style={reqWrapStyle}>
            {hiddenInput}
            <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 10 }} wrap>
              <DefaultButton iconProps={{ iconName: 'Attach' }} text="Adicionar ficheiros" onClick={openPicker} />
              <div
                {...dropHandlers(true)}
                style={{
                  padding: '6px 12px',
                  borderRadius: 4,
                  border: `1px dashed ${dragOver ? accent : '#c8c6c4'}`,
                  background: dragOver ? 'rgba(0, 120, 212, 0.05)' : 'transparent',
                }}
              >
                <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                  Largar aqui
                </Text>
              </div>
            </Stack>
          </div>
        )}
        {renderFileRows(useCompactChips)}
        {footerErr}
      </Stack>
    );
  }

  if (layout === 'card') {
    return (
      <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
        {header}
        {!disabled && (
          <div style={reqWrapStyle}>
            {hiddenInput}
            <div
              style={{
                borderRadius: 12,
                boxShadow: '0 3px 14px rgba(0,0,0,0.08)',
                border: '1px solid #edebe9',
                padding: 20,
                background: '#fff',
              }}
            >
              <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 12 }} styles={{ root: { marginBottom: 12 } }}>
                <div
                  style={{
                    width: 44,
                    height: 44,
                    borderRadius: 10,
                    background: `linear-gradient(135deg, ${accent}22, ${accent}44)`,
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                  }}
                >
                  <Icon iconName="CloudUpload" styles={{ root: { fontSize: 22, color: accent } }} />
                </div>
                <Stack tokens={{ childrenGap: 2 }}>
                  <Text styles={{ root: { fontWeight: 600, fontSize: 15 } }}>Anexar ficheiros</Text>
                  <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                    Arraste para aqui ou escolha no dispositivo
                  </Text>
                </Stack>
              </Stack>
              <div
                role="button"
                tabIndex={0}
                onClick={openPicker}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    openPicker();
                  }
                }}
                {...dropHandlers(true)}
                style={{
                  cursor: disabled ? 'default' : 'pointer',
                  minHeight: 88,
                  borderRadius: 10,
                  border: `2px dashed ${dragOver ? accent : '#c8c6c4'}`,
                  background: dragOver ? 'rgba(0, 120, 212, 0.07)' : '#faf9f8',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  gap: 8,
                }}
              >
                <Icon iconName="Upload" styles={{ root: { fontSize: 28, color: dragOver ? accent : '#8a8886' } }} />
                <Text variant="small" styles={{ root: { color: '#323130', fontWeight: 500 } }}>
                  Clique ou largue ficheiros
                </Text>
              </div>
            </div>
          </div>
        )}
        {renderFileRows(false)}
        {footerErr}
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
      {header}
      {!disabled && (
        <div style={reqWrapStyle}>
          {hiddenInput}
          <div
            role="button"
            tabIndex={0}
            onClick={openPicker}
            onKeyDown={(e) => {
              if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                openPicker();
              }
            }}
            {...dropHandlers(true)}
            style={{
              cursor: 'pointer',
              minHeight: 112,
              borderRadius: 10,
              border: `2px dashed ${dragOver ? accent : requiredEmptyHighlight ? borderErr : '#c8c6c4'}`,
              background: dragOver ? 'rgba(0, 120, 212, 0.08)' : requiredEmptyHighlight ? bgErr : '#faf9f8',
              display: 'flex',
              flexDirection: 'column',
              alignItems: 'center',
              justifyContent: 'center',
              gap: 10,
              padding: 16,
              transition: 'border-color 0.15s ease, background 0.15s ease',
            }}
          >
            <Icon iconName="CloudUpload" styles={{ root: { fontSize: 36, color: accent } }} />
            <Text styles={{ root: { fontWeight: 600, fontSize: 15, textAlign: 'center' } }}>
              Largue ficheiros aqui
            </Text>
            <Text variant="small" styles={{ root: { color: '#605e5c', textAlign: 'center' } }}>
              ou clique para escolher — vários ficheiros permitidos
            </Text>
          </div>
        </div>
      )}
      {renderFileRows(false)}
      {footerErr}
    </Stack>
  );
};
