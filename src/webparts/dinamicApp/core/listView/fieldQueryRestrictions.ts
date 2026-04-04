import type { IFieldMetadata } from '../../../../services/shared/types';
import { getInternalFieldType } from '../table/utils/fieldTypeMapper';

export function isNoteFieldMeta(meta: IFieldMetadata | undefined): boolean {
  if (!meta) return false;
  return getInternalFieldType(meta.TypeAsString, meta.MappedType) === 'note';
}

export function isNoteFieldPath(fieldPath: string, fieldsMetadata: IFieldMetadata[] | undefined): boolean {
  if (!fieldPath.trim() || !fieldsMetadata?.length) return false;
  const root = fieldPath.split('/')[0];
  for (let i = 0; i < fieldsMetadata.length; i++) {
    if (fieldsMetadata[i].InternalName === root) {
      return isNoteFieldMeta(fieldsMetadata[i]);
    }
  }
  return false;
}
