import * as React from 'react';
import { useCallback } from 'react';
import { Stack, Text, Label, IconButton } from '@fluentui/react';

export interface IFormAttachmentUploaderProps {
  files: File[];
  onFilesChange: (files: File[]) => void;
  disabled: boolean;
  label: string;
  description?: string;
  errorMessage?: string;
  required?: boolean;
}

function mergeFiles(prev: File[], added: FileList | null): File[] {
  if (!added || added.length === 0) return prev.slice();
  const next = prev.slice();
  for (let i = 0; i < added.length; i++) next.push(added[i]);
  return next;
}

export const FormAttachmentUploader: React.FC<IFormAttachmentUploaderProps> = ({
  files,
  onFilesChange,
  disabled,
  label,
  description,
  errorMessage,
  required,
}) => {
  const onInputChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>): void => {
      const fl = e.target.files;
      onFilesChange(mergeFiles(files, fl));
      e.target.value = '';
    },
    [files, onFilesChange]
  );

  const removeAt = useCallback(
    (idx: number): void => {
      onFilesChange(files.filter((_, i) => i !== idx));
    },
    [files, onFilesChange]
  );

  return (
    <Stack tokens={{ childrenGap: 8 }} styles={{ root: { marginBottom: 12 } }}>
      <Label required={required}>{label}</Label>
      {description && (
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          {description}
        </Text>
      )}
      {!disabled && (
        <input type="file" multiple onChange={onInputChange} style={{ maxWidth: '100%' }} />
      )}
      {files.length > 0 && (
        <Stack tokens={{ childrenGap: 4 }}>
          {files.map((f, idx) => (
            <Stack
              key={`${f.name}_${idx}_${f.size}`}
              horizontal
              verticalAlign="center"
              horizontalAlign="space-between"
              tokens={{ childrenGap: 8 }}
              styles={{
                root: {
                  padding: '6px 10px',
                  background: '#faf9f8',
                  borderRadius: 4,
                  border: '1px solid #edebe9',
                },
              }}
            >
              <Text variant="small" styles={{ root: { overflow: 'hidden', textOverflow: 'ellipsis' } }}>
                {f.name}
                <span style={{ color: '#605e5c', marginLeft: 8 }}>{(f.size / 1024).toFixed(1)} KB</span>
              </Text>
              {!disabled && (
                <IconButton
                  iconProps={{ iconName: 'Cancel' }}
                  title="Remover"
                  ariaLabel={`Remover ${f.name}`}
                  onClick={() => removeAt(idx)}
                />
              )}
            </Stack>
          ))}
        </Stack>
      )}
      {errorMessage && (
        <Text variant="small" styles={{ root: { color: '#a4262c' } }}>
          {errorMessage}
        </Text>
      )}
    </Stack>
  );
};
