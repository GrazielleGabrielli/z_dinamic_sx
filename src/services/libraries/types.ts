import { IListMetadata } from '../shared/types';

export type ILibraryMetadata = IListMetadata;

export interface ILibrarySummary {
  Id: string;
  Title: string;
  ItemCount: number;
  Description: string;
}

export interface IDocumentField {
  InternalName: string;
  Title: string;
  TypeAsString: string;
  Required: boolean;
}
