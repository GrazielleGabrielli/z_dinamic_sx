import * as React from 'react';
import { Stack, Text, TextField, Checkbox, Spinner } from '@fluentui/react';
import { normalizeLookupUserFieldPath } from '../../core/formManager/formButtonLookupUserVisibility';

export interface ILookupUserFieldPathsSectionProps {
  title?: string;
  description?: string;
  paths: string[] | undefined;
  onPathsChange: (paths: string[] | undefined) => void;
  options: { path: string; label: string }[];
  optionsLoading?: boolean;
  filter: string;
  onFilterChange: (value: string) => void;
  disabled?: boolean;
}

export const LookupUserFieldPathsSection: React.FC<ILookupUserFieldPathsSectionProps> = ({
  title = 'Campos de utilizador',
  description = 'Utilizador atual no campo (lista principal ou via lookup). Vazio = todos.',
  paths,
  onPathsChange,
  options,
  optionsLoading,
  filter,
  onFilterChange,
  disabled,
}) => {
  const cur = paths ?? [];
  const q = filter.trim().toLowerCase();
  const filteredOptions = q
    ? options.filter((o) => o.label.toLowerCase().includes(q) || o.path.toLowerCase().includes(q))
    : options;

  return (
    <Stack tokens={{ childrenGap: 6 }}>
      <Text variant="small" styles={{ root: { fontWeight: 600 } }}>
        {title}
      </Text>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        {description}
      </Text>
      <TextField
        placeholder="Filtrar campos por nome"
        value={filter}
        disabled={disabled}
        onChange={(_: unknown, v?: string) => onFilterChange(v ?? '')}
        styles={{ root: { maxWidth: 420 } }}
      />
      {optionsLoading ? <Spinner label="A carregar campos das listas ligadas…" /> : null}
      {!optionsLoading ? (
        <Stack
          tokens={{ childrenGap: 6 }}
          styles={{
            root: {
              maxHeight: 200,
              overflowY: 'auto',
              border: '1px solid #edebe9',
              borderRadius: 4,
              padding: 8,
            },
          }}
        >
          {cur
            .filter(
              (p) => !options.some((o) => normalizeLookupUserFieldPath(o.path) === normalizeLookupUserFieldPath(p))
            )
            .filter((p) => !q || p.toLowerCase().includes(q))
            .map((p, oi) => (
              <Checkbox
                key={`orphan-lu-${oi}-${p}`}
                label={`${p} (guardado; não encontrado)`}
                checked
                disabled={disabled}
                onChange={(_, c) => {
                  if (c) return;
                  const n = normalizeLookupUserFieldPath(p);
                  const next = cur.filter((x) => normalizeLookupUserFieldPath(x) !== n);
                  onPathsChange(next.length ? next : undefined);
                }}
              />
            ))}
          {filteredOptions.map((o) => {
            const n = normalizeLookupUserFieldPath(o.path);
            const checked = cur.some((x) => normalizeLookupUserFieldPath(x) === n);
            return (
              <Checkbox
                key={o.path}
                label={o.label}
                checked={checked}
                disabled={disabled}
                onChange={(_, c) => {
                  let next: string[];
                  if (c) {
                    next = checked ? cur : cur.concat([o.path]);
                  } else {
                    next = cur.filter((x) => normalizeLookupUserFieldPath(x) !== n);
                  }
                  onPathsChange(next.length ? next : undefined);
                }}
              />
            );
          })}
          {options.length > 0 && !filteredOptions.length && q ? (
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Nenhum campo corresponde ao filtro.
            </Text>
          ) : null}
          {!options.length && !cur.length ? (
            <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
              Nenhum campo user ou lookup→user no formulário.
            </Text>
          ) : null}
        </Stack>
      ) : null}
    </Stack>
  );
};
