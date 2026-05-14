import type { FieldMappedType, IFieldMetadata } from '../../../../services';

const ITEM_VERSION_SNAPSHOT_OMIT_TYPES = new Set<FieldMappedType>([
  'lookup',
  'lookupmulti',
  'user',
  'usermulti',
  'taxonomy',
  'taxonomymulti',
]);

export function fieldsForVersionSnapshotPicker(meta: IFieldMetadata[]): IFieldMetadata[] {
  return meta
    .filter(
      (f) =>
        !f.Hidden &&
        !ITEM_VERSION_SNAPSHOT_OMIT_TYPES.has(f.MappedType) &&
        /^[A-Za-z0-9_]+$/.test(f.InternalName.trim())
    )
    .sort((a, b) => a.Title.localeCompare(b.Title, undefined, { sensitivity: 'base' }));
}

export function fieldInternalsForItemVersionODataSelect(meta: IFieldMetadata[], max = 45): string[] {
  return fieldsForVersionSnapshotPicker(meta)
    .map((f) => f.InternalName.trim())
    .slice(0, max);
}
