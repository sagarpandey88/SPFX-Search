
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { SearchProvider } from "../providers/spSearchProvider";
import { ISearchQueryOptions, ISearchResult } from "../models/searchModel";
import { defaultSelectProperties } from "./searchUtils";

export class SearchService {
    private context: BaseComponentContext;

    constructor(context: BaseComponentContext) {
        this.context = context;
    }

    public async search(options: ISearchQueryOptions): Promise<ISearchResult> {
        const searchProvider = new SearchProvider(this.context);
        const contentClassFilter = "(contentclass:STS_ListItem OR " +
            "contentclass:STS_Web OR " +
            "contentclass:STS_Site)";
console.log("Refinement filters:", options.refinementFilters);
        const opts: ISearchQueryOptions = {
            ...options,
            refinementFilters: options.refinementFilters ?? [],
            rowLimit: options.rowLimit ?? 5,
            selectProperties: options.selectProperties ?? defaultSelectProperties,
            queryTemplate: options.queryTemplate
                ? `${options.queryTemplate} ${contentClassFilter}`
                : `{searchTerms} ${contentClassFilter}`,
            refinerColumns: options.refinerColumns ?? "",
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
            const rawSnippet = (r.HitHighlightedSummary || r.Summary || "") as string;
            // SharePoint wraps matched keyword spans with <c0>…</c0> and uses <ddd/> as an
            // ellipsis marker.  Convert to safe <mark> tags and strip everything else.
            const snippet = rawSnippet
                ? String(rawSnippet)
                    .replace(/<c0>/gi, "<mark>")
                    .replace(/<\/c0>/gi, "</mark>")
                    .replace(/<ddd\/>/gi, "\u2026")
                    .replace(/<(?!\/?mark)[^>]*>/g, "") // strip all tags except <mark>
                : "";
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
                Snippet: snippet,
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




