import { IListMetadata, ListSourceType } from '../shared/types';

export { IListMetadata, ListSourceType };

export interface IListSummary {
  Id: string;
  Title: string;
  BaseTemplate: number;
  IsLibrary: boolean;
  Hidden: boolean;
  ItemCount: number;
}
