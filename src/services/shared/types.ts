export interface IBaseServiceResponse<T> {
  data: T;
  error?: string;
}

export interface IPagedResult<T> {
  items: T[];
  hasNext: boolean;
  nextSkip: number;
  total?: number;
}

export interface IListMetadata {
  Id: string;
  Title: string;
  BaseTemplate: number;
  IsLibrary: boolean;
  ItemCount: number;
  Description: string;
  LastItemModifiedDate: string;
  EnableVersioning: boolean;
  Hidden: boolean;
}

export interface IFieldMetadata {
  Id: string;
  Title: string;
  InternalName: string;
  TypeAsString: string;
  MappedType: FieldMappedType;
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

export type FieldMappedType =
  | 'text'
  | 'multiline'
  | 'choice'
  | 'multichoice'
  | 'number'
  | 'currency'
  | 'boolean'
  | 'datetime'
  | 'lookup'
  | 'lookupmulti'
  | 'user'
  | 'usermulti'
  | 'url'
  | 'calculated'
  | 'taxonomy'
  | 'taxonomymulti'
  | 'unknown';

export interface IViewMetadata {
  Id: string;
  Title: string;
  DefaultView: boolean;
  Hidden: boolean;
  PersonalView: boolean;
  RowLimit: number;
  ViewQuery: string;
  ViewFields: string[];
}

export interface IUserSummary {
  Id: number;
  Title: string;
  Email: string;
  LoginName: string;
  IsSiteAdmin: boolean;
}

export interface IGroupSummary {
  Id: number;
  Title: string;
  Description: string;
  OwnerTitle: string;
  AllowMembersEditMembership: boolean;
}

export type ListSourceType = 'list' | 'library' | 'unknown';
