import { IRefiner } from "@pnp/sp/search";

export interface ISearchQueryOptions {
    queryText: string;
    queryTemplate?: string;
    sortBy?: string;
    sortOrder?: string;
    selectProperties?: string[];
    refinementFilters?: string[];
    refinerColumns?: string;
    rowLimit?: number;
    startRow?: number;
    enableQueryRules?: boolean;
    multiColumnSortBy?: any[];
    cacheEnabled?: boolean;
}

/** Represents a single normalized search result item shown in the UI. */
export interface ISearchResultItem {
    key: string;
    Title: string;
    Path: string;
    Snippet: string;
    Thumbnail?: string;
    Author?: string;
    Modified?: string;
    FileType?: string;
    Views?: string;
    SiteName?: string;
    SiteUrl?: string;
}

/** The shape returned by SearchService.search(): normalized items plus raw refiner buckets. */
export interface ISearchResult {
    items: ISearchResultItem[];
    refiners: IRefiner[]; //any[];
    totalRows: number;
}

export interface IRefinerConfig {
  displayName: string;
  managedProperty: string;
  isTaxonomy?: boolean;
}
export interface ITaxonomyEntry {
  token: string;
  label: string;
  count: string;
}
