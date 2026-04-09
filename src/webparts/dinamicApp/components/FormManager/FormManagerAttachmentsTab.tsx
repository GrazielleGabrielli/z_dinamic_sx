import * as React from 'react';
import { useState, useEffect, useMemo } from 'react';
import {
  Stack,
  Text,
  Dropdown,
  IDropdownOption,
  Spinner,
  MessageBar,
  MessageBarType,
  ChoiceGroup,
  IChoiceGroupOption,
  Checkbox,
} from '@fluentui/react';
import { ListsService, FieldsService } from '../../../../services';
import type { IListSummary, IFieldMetadata } from '../../../../services';
import type {
  TFormAttachmentUploadLayoutKind,
  TFormAttachmentFilePreviewKind,
  TFormAttachmentStorageKind,
  IAttachmentLibraryFolderTreeNode,
} from '../../core/config/types/formManager';
import { FormManagerFolderTreeEditor, type IFolderVisibilityEditorProps } from './FormManagerFolderTreeEditor';
import { FormAttachmentUploader } from './FormAttachmentUploader';
import {
  FORM_ATTACHMENT_LAYOUT_DROPDOWN_OPTIONS,
  FORM_ATTACHMENT_FILE_PREVIEW_DROPDOWN_OPTIONS,
  FORM_ATTACHMENT_EXTENSION_GROUPS,
  FormManagerCollapseSection,
} from './FormManagerComponentsTab';

const STORAGE_OPTIONS: IChoiceGroupOption[] = [
  { key: 'itemAttachments', text: 'Anexos ao item (lista principal)' },
  { key: 'documentLibrary', text: 'Biblioteca de documentos' },
];

function normListGuid(g: string | undefined): string {
  if (!g) return '';
  return g.replace(/[{}]/g, '').toLowerCase();
}

const SECTION_IDS = {
  storage: 'attStorage',
  folders: 'attFolders',
  ui: 'attUi',
  ext: 'attExt',
} as const;

export interface IFormManagerAttachmentsTabProps {
  loading: boolean;
  primaryListTitle: string;
  attachmentStorageKind: TFormAttachmentStorageKind;
  onAttachmentStorageKindChange: (v: TFormAttachmentStorageKind) => void;
  attachmentLibraryTitle: string;
  onAttachmentLibraryTitleChange: (v: string) => void;
  attachmentLibraryLookupField: string;
  onAttachmentLibraryLookupFieldChange: (v: string) => void;
  attachmentLibFolderTree: IAttachmentLibraryFolderTreeNode[];
  onAttachmentLibFolderTreeChange: (tree: IAttachmentLibraryFolderTreeNode[]) => void;
  attachmentUploadLayout: TFormAttachmentUploadLayoutKind;
  onAttachmentUploadLayoutChange: (v: TFormAttachmentUploadLayoutKind) => void;
  attachmentFilePreview: TFormAttachmentFilePreviewKind;
  onAttachmentFilePreviewChange: (v: TFormAttachmentFilePreviewKind) => void;
  attachmentAllowedExtensions: string[];
  onAttachmentExtensionToggle: (ext: string, selected: boolean) => void;
  attachmentFolderStepOptions: { id: string; title: string }[];
  attachmentFolderVisibilityEditor?: IFolderVisibilityEditorProps;
}

export function FormManagerAttachmentsTabContent(props: IFormManagerAttachmentsTabProps): JSX.Element {
  const {
    loading,
    primaryListTitle,
    attachmentStorageKind,
    onAttachmentStorageKindChange,
    attachmentLibraryTitle,
    onAttachmentLibraryTitleChange,
    attachmentLibraryLookupField,
    onAttachmentLibraryLookupFieldChange,
    attachmentLibFolderTree,
    onAttachmentLibFolderTreeChange,
    attachmentUploadLayout,
    onAttachmentUploadLayoutChange,
    attachmentFilePreview,
    onAttachmentFilePreviewChange,
    attachmentAllowedExtensions,
    onAttachmentExtensionToggle,
    attachmentFolderStepOptions,
    attachmentFolderVisibilityEditor,
  } = props;

  const [libs, setLibs] = useState<IListSummary[]>([]);
  const [libsLoading, setLibsLoading] = useState(false);
  const [libsErr, setLibsErr] = useState<string | undefined>(undefined);
  const [libFields, setLibFields] = useState<IFieldMetadata[]>([]);
  const [libFieldsLoading, setLibFieldsLoading] = useState(false);
  const [libFieldsErr, setLibFieldsErr] = useState<string | undefined>(undefined);
  const [primaryListId, setPrimaryListId] = useState<string | undefined>(undefined);
  const [primaryListLoading, setPrimaryListLoading] = useState(false);
  const [openSections, setOpenSections] = useState<Record<string, boolean>>({});
  const [attachDemoFiles, setAttachDemoFiles] = useState<File[]>([]);

  const toggleSection = (id: string): void => {
    setOpenSections((prev) => ({ ...prev, [id]: !prev[id] }));
  };
  const isOpen = (id: string): boolean => openSections[id] === true;

  useEffect(() => {
    setLibsErr(undefined);
    setLibsLoading(true);
    const svc = new ListsService();
    svc
      .getLists(false)
      .then((data) => {
        setLibs(data.filter((l) => l.IsLibrary));
        setLibsLoading(false);
      })
      .catch((e) => {
        setLibs([]);
        setLibsLoading(false);
        setLibsErr(e instanceof Error ? e.message : String(e));
      });
  }, []);

  useEffect(() => {
    const t = attachmentLibraryTitle.trim();
    if (!t) {
      setLibFields([]);
      setLibFieldsErr(undefined);
      return;
    }
    setLibFieldsErr(undefined);
    setLibFieldsLoading(true);
    const fs = new FieldsService();
    fs.getVisibleFields(t)
      .then((fields) => {
        setLibFields(fields);
        setLibFieldsLoading(false);
      })
      .catch((e) => {
        setLibFields([]);
        setLibFieldsLoading(false);
        setLibFieldsErr(e instanceof Error ? e.message : String(e));
      });
  }, [attachmentLibraryTitle]);

  useEffect(() => {
    const t = primaryListTitle.trim();
    if (!t) {
      setPrimaryListId(undefined);
      return;
    }
    setPrimaryListLoading(true);
    const ls = new ListsService();
    ls.getListByTitle(t)
      .then((m) => {
        setPrimaryListId(m.Id);
        setPrimaryListLoading(false);
      })
      .catch(() => {
        setPrimaryListId(undefined);
        setPrimaryListLoading(false);
      });
  }, [primaryListTitle]);

  const linkLookupFields = useMemo(() => {
    if (!primaryListId) return [];
    const target = normListGuid(primaryListId);
    return libFields.filter(
      (f) =>
        f.MappedType === 'lookup' &&
        !f.Hidden &&
        !f.ReadOnlyField &&
        !f.AllowMultipleValues &&
        normListGuid(f.LookupList) === target
    );
  }, [libFields, primaryListId]);

  const libOptions: IDropdownOption[] = useMemo(
    () => [{ key: '', text: '—' }, ...libs.map((l) => ({ key: l.Title, text: l.Title }))],
    [libs]
  );

  const linkFieldOptions: IDropdownOption[] = useMemo(() => {
    const base: IDropdownOption[] = [{ key: '', text: '—' }];
    const known = new Set(linkLookupFields.map((f) => f.InternalName));
    if (attachmentLibraryLookupField && !known.has(attachmentLibraryLookupField)) {
      base.push({
        key: attachmentLibraryLookupField,
        text: `${attachmentLibraryLookupField} (referência guardada)`,
      });
    }
    return base.concat(
      linkLookupFields.map((f) => ({
        key: f.InternalName,
        text: `${f.Title} (${f.InternalName})`,
      }))
    );
  }, [linkLookupFields, attachmentLibraryLookupField]);

  if (loading) {
    return (
      <Stack tokens={{ childrenGap: 12 }} styles={{ root: { marginTop: 12 } }}>
        <Spinner label="A carregar campos da lista…" />
      </Stack>
    );
  }

  return (
    <Stack tokens={{ childrenGap: 10 }} styles={{ root: { marginTop: 12 } }}>
      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
        Aqui define-se <span style={{ fontWeight: 600 }}>onde</span> os ficheiros são gravados (anexos do item na lista
        principal ou biblioteca com lookup) e o aspeto do controlo. <span style={{ fontWeight: 600 }}>Em que etapa</span>{' '}
        o upload aparece no formulário define-se na aba «Estrutura», ao colocar o campo virtual «Anexos ao item» na etapa
        certa.
      </Text>

      <FormManagerCollapseSection
        title="Destino do upload"
        isOpen={isOpen(SECTION_IDS.storage)}
        onToggle={() => toggleSection(SECTION_IDS.storage)}
      >
        <ChoiceGroup
          options={STORAGE_OPTIONS}
          selectedKey={attachmentStorageKind}
          onChange={(_, o) =>
            o && onAttachmentStorageKindChange(String(o.key) as TFormAttachmentStorageKind)
          }
        />
        {attachmentStorageKind === 'documentLibrary' && (
          <Stack tokens={{ childrenGap: 12 }} styles={{ root: { maxWidth: 520, marginTop: 8 } }}>
            {libsLoading && <Spinner label="A carregar bibliotecas…" />}
            {libsErr && <MessageBar messageBarType={MessageBarType.error}>{libsErr}</MessageBar>}
            <Dropdown
              label="Biblioteca de documentos"
              options={libOptions}
              selectedKey={attachmentLibraryTitle || ''}
              onChange={(_, o) => {
                const t = o ? String(o.key) : '';
                onAttachmentLibraryTitleChange(t);
              }}
              styles={{ root: { maxWidth: 480 } }}
              disabled={libsLoading}
            />
            {attachmentLibraryTitle.trim() ? (
              <>
                {libFieldsLoading && <Spinner label="A carregar campos da biblioteca…" />}
                {libFieldsErr && <MessageBar messageBarType={MessageBarType.error}>{libFieldsErr}</MessageBar>}
                {primaryListLoading && <Spinner label="A resolver a lista principal do formulário…" />}
                {!primaryListTitle.trim() && (
                  <MessageBar messageBarType={MessageBarType.info}>
                    Indique o título da lista principal no separador «Geral» (origem dos dados) para escolher o lookup de
                    vínculo ao item.
                  </MessageBar>
                )}
                {!!primaryListTitle.trim() && !primaryListLoading && !primaryListId && (
                  <MessageBar messageBarType={MessageBarType.warning}>
                    Não foi possível obter a lista principal «{primaryListTitle}». Confira o título no separador «Geral».
                  </MessageBar>
                )}
                <Dropdown
                  label="Lookup para a lista principal (vínculo ao item)"
                  options={linkFieldOptions}
                  selectedKey={attachmentLibraryLookupField || ''}
                  onChange={(_, o) => {
                    const k = o ? String(o.key) : '';
                    onAttachmentLibraryLookupFieldChange(k);
                  }}
                  styles={{ root: { maxWidth: 480 } }}
                  disabled={
                    libFieldsLoading || !!libFieldsErr || primaryListLoading || !primaryListId
                  }
                />
                {!libFieldsLoading &&
                  !libFieldsErr &&
                  primaryListId &&
                  linkLookupFields.length === 0 && (
                    <MessageBar messageBarType={MessageBarType.warning}>
                      Não há coluna de lookup nesta biblioteca que aponte para «{primaryListTitle}». Crie uma coluna
                      lookup para essa lista.
                    </MessageBar>
                  )}
                {attachmentLibraryLookupField.trim() && (
                  <FormManagerCollapseSection
                    title="Árvore de pastas na biblioteca"
                    isOpen={isOpen(SECTION_IDS.folders)}
                    onToggle={() => toggleSection(SECTION_IDS.folders)}
                  >
                    <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                      O nível 1 é sempre a pasta com o ID do item na lista principal. Abaixo só existe uma pasta raiz
                      por solicitação; a ramificação (irmãos e filhos) fica dentro dela, para agrupar por ID. Use texto
                      fixo ou placeholders <code style={{ fontSize: 12 }}>{'{{Title}}'}</code>,{' '}
                      <code style={{ fontSize: 12 }}>{'{{NomeInterno}}'}</code>, etc. Sem pastas configuradas = ficheiros
                      diretamente na pasta do ID.
                    </Text>
                    <Stack
                      horizontal
                      verticalAlign="center"
                      tokens={{ childrenGap: 8 }}
                      styles={{
                        root: {
                          padding: '8px 10px',
                          borderRadius: 4,
                          border: '1px solid #edebe9',
                          background: '#faf9f8',
                          maxWidth: 520,
                        },
                      }}
                    >
                      <Text variant="small" styles={{ root: { minWidth: 22, color: '#605e5c', fontWeight: 600 } }}>
                        1.
                      </Text>
                      <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                        ID do item (lista principal)
                      </Text>
                      <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
                        (automático — nome da pasta = id do item)
                      </Text>
                    </Stack>
                    <FormManagerFolderTreeEditor
                      nodes={attachmentLibFolderTree}
                      onChange={onAttachmentLibFolderTreeChange}
                      folderStepOptions={attachmentFolderStepOptions}
                      showFolderStepPicker={attachmentStorageKind === 'documentLibrary'}
                      folderVisibilityEditor={attachmentFolderVisibilityEditor}
                    />
                  </FormManagerCollapseSection>
                )}
              </>
            ) : null}
          </Stack>
        )}
      </FormManagerCollapseSection>

      <FormManagerCollapseSection
        title="Aspeto e pré-visualização do controlo"
        isOpen={isOpen(SECTION_IDS.ui)}
        onToggle={() => toggleSection(SECTION_IDS.ui)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Quando incluir «Anexos ao item» na Estrutura, o controlo de ficheiros usa estes estilos.
        </Text>
        <Dropdown
          label="Tipo de layout do input de anexos"
          options={FORM_ATTACHMENT_LAYOUT_DROPDOWN_OPTIONS}
          selectedKey={attachmentUploadLayout}
          onChange={(_, o) =>
            o && onAttachmentUploadLayoutChange(String(o.key) as TFormAttachmentUploadLayoutKind)
          }
        />
        <Dropdown
          label="Pré-visualização dos ficheiros selecionados"
          options={FORM_ATTACHMENT_FILE_PREVIEW_DROPDOWN_OPTIONS}
          selectedKey={attachmentFilePreview}
          onChange={(_, o) =>
            o && onAttachmentFilePreviewChange(String(o.key) as TFormAttachmentFilePreviewKind)
          }
        />
        <Stack
          styles={{
            root: {
              border: '1px solid #edebe9',
              borderRadius: 4,
              padding: 12,
              background: '#faf9f8',
            },
          }}
          tokens={{ childrenGap: 8 }}
        >
          <Text variant="small" styles={{ root: { fontWeight: 600, color: '#605e5c' } }}>
            Pré-visualização (pode adicionar ficheiros de teste)
          </Text>
          <FormAttachmentUploader
            files={attachDemoFiles}
            onFilesChange={setAttachDemoFiles}
            disabled={false}
            label="Anexos ao item"
            description="Texto de ajuda opcional, como no formulário."
            layout={attachmentUploadLayout}
            filePreview={attachmentFilePreview}
            allowedFileExtensions={
              attachmentAllowedExtensions.length > 0 ? attachmentAllowedExtensions : undefined
            }
          />
        </Stack>
      </FormManagerCollapseSection>

      <FormManagerCollapseSection
        title="Extensões permitidas"
        isOpen={isOpen(SECTION_IDS.ext)}
        onToggle={() => toggleSection(SECTION_IDS.ext)}
      >
        <Text variant="small" styles={{ root: { color: '#605e5c' } }}>
          Nenhuma selecionada = qualquer tipo de ficheiro. Com uma ou mais marcadas, só essas extensões são aceites no
          formulário e na validação ao gravar.
        </Text>
        <Stack tokens={{ childrenGap: 10 }}>
          {FORM_ATTACHMENT_EXTENSION_GROUPS.map((group) => (
            <Stack
              key={group.title}
              tokens={{ childrenGap: 8 }}
              styles={{
                root: {
                  padding: '10px 12px',
                  borderRadius: 6,
                  border: '1px solid #edebe9',
                  background: '#faf9f8',
                },
              }}
            >
              <Stack tokens={{ childrenGap: 2 }}>
                <Text variant="small" styles={{ root: { fontWeight: 600, color: '#323130' } }}>
                  {group.title}
                </Text>
                {group.hint && (
                  <Text variant="tiny" styles={{ root: { color: '#8a8886' } }}>
                    {group.hint}
                  </Text>
                )}
              </Stack>
              <Stack horizontal wrap tokens={{ childrenGap: 10 }} verticalAlign="center">
                {group.items.map((p) => {
                  const e = p.ext.toLowerCase();
                  const checked = attachmentAllowedExtensions.some((x) => x.toLowerCase() === e);
                  return (
                    <Checkbox
                      key={p.ext}
                      label={p.label}
                      checked={checked}
                      onChange={(_, c) => onAttachmentExtensionToggle(p.ext, !!c)}
                      styles={{ root: { minWidth: 0 } }}
                    />
                  );
                })}
              </Stack>
            </Stack>
          ))}
        </Stack>
      </FormManagerCollapseSection>
    </Stack>
  );
}
