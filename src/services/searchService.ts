
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { SearchProvider } from "../providers/spSearchProvider";
import { ISearchQueryOptions, ISearchResult } from "../models/searchModel";
import { contentClassFilter, defaultSelectProperties } from "./searchUtils";

export class SearchService {
    private context: BaseComponentContext;
    constructor(context: BaseComponentContext) {
        this.context = context;
    }

    public async search(options: ISearchQueryOptions): Promise<ISearchResult> {
        const searchProvider = new SearchProvider(this.context);
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
        const rows = results && results.rows ? results.rows : [];
        console.log("Raw search results from search service:", results);

        const items = rows.map((r: any) => {
            
            const title = r.Title || r.LinkTitle || r.TitleOWSTEXT || r.PreferredName || "(no title)";
            const path = r.Path || r.SPWebUrl || r.ServerRedirectedURL || "";
            const rawSnippet = (r.HitHighlightedSummary || r.Summary || "") as string;
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
            const modified = modifiedRaw ? (new Date(modifiedRaw)).toLocaleString() : undefined;
            const fileType = r.FileType || undefined;
            const views = r.ViewsLifeTime || r.ViewsRecent || undefined;
            const siteName = r.SiteTitle || undefined;            // SiteTitle is the display name of the site; SPSiteURL / SPWebUrl give the URL.
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

        return { items, refiners: results.refiners || [], totalRows: results.totalRows ?? 0 };
    }


}




