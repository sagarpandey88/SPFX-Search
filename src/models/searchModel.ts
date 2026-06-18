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
    refiners: any[];
    totalRows: number;
}
