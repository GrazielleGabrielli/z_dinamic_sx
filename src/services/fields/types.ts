import { IFieldMetadata, FieldMappedType } from '../shared/types';

export { IFieldMetadata, FieldMappedType };

export interface IRawSPField {
  Id: string;
  Title: string;
  InternalName: string;
  TypeAsString: string;
  Required: boolean;
  ReadOnlyField: boolean;
  Hidden: boolean;
  Description: string;
  DefaultValue: string | null;
  Choices?: string[];
  LookupList?: string;
  LookupField?: string;
  AllowMultipleValues?: boolean;
  MaxLength?: number;
}
