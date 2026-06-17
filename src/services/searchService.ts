
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { SearchProvider } from "../providers/spSearchProvider";
import { ActiveFilters } from "../models/searchModel";

/** Represents a single normalized search result item shown in the UI. */
export interface SearchResultItem {
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
export interface SearchResult {
    items: SearchResultItem[];
    refiners: any[];
}

/** All options that can be passed to SearchService.search(). */
export interface SearchOptions {
    queryText: string;
    queryTemplate?: string;
    sortBy?: string;
    sortOrder?: string;
    selectProperties?: string[];
    refinementFilters?: ActiveFilters;
    refinerColumns?: string;
    rowLimit?: number;
    startRow?: number;
    enableQueryRules?: boolean;
    multiColumnSortBy?: any[];
    cacheEnabled?: boolean;
}

/**
 * Thin service layer over SearchProvider.
 * Responsible for building sensible defaults and normalizing raw PnP results
 * into a consistent SearchResultItem shape consumed by the UI.
 */
export class SearchService {
    private context: BaseComponentContext;

    constructor(context: BaseComponentContext) {
        this.context = context;
    }

    /**
 * Converts the activeFilters map into SharePoint RefinementFilter strings.
 * Single selection: "Author:token"
 * Multi-selection:  "Author:or(token1,token2)" — SharePoint OR syntax.
 */
    private buildRefinementFilters = (filters: ActiveFilters): string[] =>
        Object.entries(filters)
            .filter(([, tokens]) => tokens.length > 0)
            .map(([name, tokens]) =>
                tokens.length === 1
                    ? `${name}:${tokens[0]}`
                    : `${name}:or(${tokens.join(",")})`
            );

    public async search(options: SearchOptions): Promise<SearchResult> {
        const searchProvider = new SearchProvider(this.context);

        // Managed properties to fetch from SharePoint search for each result row.
        const defaultSelect = [
            "Title",
            "Path",
            "HitHighlightedSummary",
            "Summary",
            "PreferredName",
            "Author",
            "Editor",
            "LastModifiedTime",
            "Modified",
            "FileType",
            "ViewsLifeTime",
            "ViewsRecent",
            "ServerRedirectedURL",
            "SPWebUrl",
            "SiteTitle",
            "SPSiteURL",
            "Thumbnail",
            "PictureURL",
            "PublishingImage"
        ];

        // Restrict results to documents (library items), list items, subsites and site collections.
        // Excludes people (urn:content-class:SPSPeople / STS_Person) and other non-content objects.
        const contentClassFilter =
            "(contentclass:STS_ListItem_DocumentLibrary OR " +
            "contentclass:STS_ListItem OR " +
            "contentclass:STS_Web OR " +
            "contentclass:STS_Site)";

        const opts = {
            ...options,
            refinementFilters: this.buildRefinementFilters(options.refinementFilters || {}),
            rowLimit: options.rowLimit ?? 5,
            selectProperties: options.selectProperties ?? defaultSelect,
            // Append the content-class scope to whatever query template the caller supplies,
            // or create one from scratch. {searchTerms} is the standard placeholder SharePoint
            // replaces with the user's query text at runtime.
            queryTemplate: options.queryTemplate
                ? `${options.queryTemplate} ${contentClassFilter}`
                : `{searchTerms} ${contentClassFilter}`,
            // Refiner columns requested from SharePoint; drives the Filter Results panel.
            // Author: built-in managed property.
            // RefinableString99: mapped from ows_MyDocumentCategory.
            refinerColumns: options.refinerColumns ?? "Author,RefinableString99",
        };

        const results = await searchProvider.searchQuery(opts);
        // Rows are PnP PrimarySearchResults — plain JS objects keyed by managed property name.
        const rows = results && results.rows ? results.rows : [];
console.log("Raw search results from provider:", results);
console.log("Raw search results from provider - refiners:", results.refiners);
        const items = rows.map((r: any) => {
            // Resolve each field by trying the most specific managed property first,
            // falling back to common aliases used across different content types.
            const title = r.Title || r.LinkTitle || r.TitleOWSTEXT || r.PreferredName || "(no title)";
            const path = r.Path || r.SPWebUrl || r.ServerRedirectedURL || "";
            const snippet = (r.HitHighlightedSummary || r.Summary || "") as string;
            const thumbnail = r.Thumbnail || r.ThumbnailUrl || r.PictureURL || r.PublishingImage || r.SiteLogo || undefined;
            const author = r.Author || r.Editor || r.PreferredName || undefined;
            const modifiedRaw = r.LastModifiedTime || r.Modified || undefined;
            // Format the date for display; fall back to undefined if unavailable.
            const modified = modifiedRaw ? (new Date(modifiedRaw)).toLocaleString() : undefined;
            const fileType = r.FileType || undefined;
            const views = r.ViewsLifeTime || r.ViewsRecent || undefined;
            // SiteTitle is the display name of the site; SPSiteURL / SPWebUrl give the URL.
            const siteName = r.SiteTitle || undefined;
            const siteUrl = r.SPSiteURL || r.SPWebUrl || undefined;

            return {
                key: r.Id || r.UniqueId || Math.random().toString(36).substring(2, 9),
                Title: title,
                Path: path,
                // Strip any residual HTML tags from HitHighlightedSummary.
                Snippet: snippet ? String(snippet).replace(/<[^>]*>/g, "") : "",
                Thumbnail: thumbnail,
                Author: author,
                Modified: modified,
                FileType: fileType,
                Views: views,
                SiteName: siteName,
                SiteUrl: siteUrl,
            };
        });

        return { items, refiners: results.refiners || [] };
    }


}




