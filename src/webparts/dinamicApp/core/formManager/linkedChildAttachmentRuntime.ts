import type {
  IAttachmentLibraryFolderTreeNode,
  IFormLinkedChildFormConfig,
  IFormManagerConfig,
} from '../config/types/formManager';
import { isFormAttachmentLibraryRuntime } from './formAttachmentLibrary';

export type TLinkedChildAttachmentResolved =
  | { kind: 'none' }
  | { kind: 'itemAttachments' }
  | {
      kind: 'documentLibrary';
      libraryTitle: string;
      lookupFieldInternalName: string;
      folderTree?: IAttachmentLibraryFolderTreeNode[];
      folderPathSegments?: string[];
    };

export function resolveLinkedChildAttachmentRuntime(
  cfg: IFormLinkedChildFormConfig,
  formManager: IFormManagerConfig
): TLinkedChildAttachmentResolved {
  const sk = cfg.childAttachmentStorageKind;
  if (!sk || sk === 'none') return { kind: 'none' };
  if (sk === 'itemAttachments') return { kind: 'itemAttachments' };
  if (sk === 'documentLibraryInheritMain') {
    if (!isFormAttachmentLibraryRuntime(formManager)) return { kind: 'none' };
    const lib = formManager.attachmentLibrary!;
    const lt = lib.libraryTitle?.trim() ?? '';
    const lkChild = (cfg.childAttachmentLibraryLookupToChildListField ?? '').trim();
    if (!lt || !lkChild) return { kind: 'none' };
    return {
      kind: 'documentLibrary',
      libraryTitle: lt,
      lookupFieldInternalName: lkChild,
      folderTree: lib.folderTree,
      folderPathSegments: lib.folderPathSegments,
    };
  }
  if (sk === 'documentLibraryCustom') {
    const c = cfg.childAttachmentLibrary;
    const lt = c?.libraryTitle?.trim() ?? '';
    const lk = c?.sourceListLookupFieldInternalName?.trim() ?? '';
    if (!lt || !lk) return { kind: 'none' };
    return {
      kind: 'documentLibrary',
      libraryTitle: lt,
      lookupFieldInternalName: lk,
      folderTree: c?.folderTree,
      folderPathSegments: c?.folderPathSegments,
    };
  }
  return { kind: 'none' };
}

export function buildMinimalFormManagerForLinkedLibraryUpload(
  resolved: Extract<TLinkedChildAttachmentResolved, { kind: 'documentLibrary' }>
): IFormManagerConfig {
  return {
    sections: [{ id: 'main', title: 'Geral', visible: true }],
    fields: [],
    rules: [],
    attachmentStorageKind: 'documentLibrary',
    attachmentLibrary: {
      libraryTitle: resolved.libraryTitle,
      sourceListLookupFieldInternalName: resolved.lookupFieldInternalName,
      ...(resolved.folderTree?.length ? { folderTree: resolved.folderTree } : {}),
      ...(resolved.folderPathSegments?.length ? { folderPathSegments: resolved.folderPathSegments } : {}),
    },
  };
}
