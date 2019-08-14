import { ISearchResult } from './ISearchResult';

export interface IGraphSearchApiState {
  results: Array<ISearchResult>;
  searchFor: string;
  resultType: string;
  includeFiles: boolean;
  includeMessages: boolean;
  includeEvents: boolean;
}