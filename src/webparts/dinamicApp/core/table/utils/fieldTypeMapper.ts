import type { InternalFieldType } from '../types';
import type { FieldMappedType } from '../../../../../services/shared/types';

const SP_TO_INTERNAL: Record<string, InternalFieldType> = {
  Text: 'text',
  Note: 'note',
  Number: 'number',
  Currency: 'currency',
  DateTime: 'date',
  Boolean: 'boolean',
  Choice: 'choice',
  MultiChoice: 'multiChoice',
  Lookup: 'lookup',
  LookupMulti: 'lookupMulti',
  User: 'user',
  UserMulti: 'userMulti',
  URL: 'url',
  Calculated: 'calculated',
  TaxonomyFieldType: 'managedMetadata',
  TaxonomyFieldTypeMulti: 'managedMetadata',
  File: 'file',
  FileLeafRef: 'file',
  Attachments: 'file',
  Image: 'image',
};

const MAPPED_TO_INTERNAL: Record<FieldMappedType, InternalFieldType> = {
  text: 'text',
  multiline: 'note',
  number: 'number',
  currency: 'currency',
  boolean: 'boolean',
  datetime: 'date',
  choice: 'choice',
  multichoice: 'multiChoice',
  lookup: 'lookup',
  lookupmulti: 'lookupMulti',
  user: 'user',
  usermulti: 'userMulti',
  url: 'url',
  calculated: 'calculated',
  taxonomy: 'managedMetadata',
  taxonomymulti: 'managedMetadata',
  unknown: 'unknown',
};

export function mapSharePointTypeToInternal(typeAsString: string): InternalFieldType {
  return SP_TO_INTERNAL[typeAsString] ?? 'unknown';
}

export function mapMappedTypeToInternal(mappedType: FieldMappedType): InternalFieldType {
  return MAPPED_TO_INTERNAL[mappedType] ?? 'unknown';
}

export function getInternalFieldType(
  typeAsString?: string,
  mappedType?: FieldMappedType
): InternalFieldType {
  if (typeAsString) return mapSharePointTypeToInternal(typeAsString);
  if (mappedType) return mapMappedTypeToInternal(mappedType);
  return 'unknown';
}
