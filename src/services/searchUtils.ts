import { ISearchQuery, SPFx, Web, spfi } from "@pnp/sp/presets/all";
import { IRefinerConfig, ITaxonomyEntry } from "../models/searchModel";

import "@pnp/sp/search";
import { SearchService } from "./searchService";


/** Number of results per page. Change this to adjust pagination globally. */
export const PAGE_SIZE = 5;
export const contentClassFilter = "(contentclass:STS_ListItem OR contentclass:STS_Web OR contentclass:STS_Site)";
export const REFINER_CONFIG: IRefinerConfig[] = [
  { displayName: "Author", managedProperty: "Author" },
  { displayName: "Category", managedProperty: "RefinableString99", isTaxonomy: false },
  { displayName: "Country", managedProperty: "RefinableString00", isTaxonomy: true },
];
export const defaultSelectProperties = [
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
  "PublishingImage",
  //"AutoPreviewUrl",
  //"ServerRedirectedEmbedURL"

];

export function getQueryParam(name: string): string | null {
  const params = new URLSearchParams(window.location.search);
  return params.get(name) || null;
}


export function buildRefinementFilters(filters: Record<string, string[]>): string[] {

  // Single selection: "Author:token"
  // Multi-selection:  "Author:or(token1,token2)" — SharePoint OR syntax.
  return Object.entries(filters)
    .filter(([, tokens]) => tokens.length > 0)
    .map(([name, tokens]) =>
      tokens.length === 1
        ? `${name}:${tokens[0]}`
        : `${name}:or(${tokens.join(",")})`
    );
}


export async function getHubSiteId(context: any): Promise<string | null> {
  try {
    const sp = spfi().using(SPFx(context));
    const { HubSiteId } = await sp.site.select("IsHubSite", "HubSiteId")();

    // HubSiteId is a zero GUID (00000000-0000-0000-0000-000000000000) when the
    // site has no hub association — treat that as "not a hub member".
    if (!HubSiteId || HubSiteId === "00000000-0000-0000-0000-000000000000") return null;

    return HubSiteId;
  } catch {
    return null;
  }
}

export function getProcessedEntries(entries: any[], isTaxonomy?: boolean): ITaxonomyEntry[] {
  if (!entries) return [];

  // 1. If it's a taxonomy field, filter out GTSet and GP0 noise rows
  const targets = isTaxonomy
    ? entries.filter(e => e.RefinementValue && e.RefinementValue.startsWith("L0|"))
    : entries;

  // 2. Map items into a standardized UI layout format
  return targets.map(entry => {
    let cleanLabel = entry.RefinementName;

    if (isTaxonomy) {
      const parts = entry.RefinementValue.split('|');
      cleanLabel = parts[parts.length - 1];
    }

    return {
      token: entry.RefinementToken,
      label: cleanLabel,
      count: entry.RefinementCount
    };
  });
}


export async function getSiteTitleById(siteId: string, context: any): Promise<string | null> {

  const results = await new SearchService(context).search({
    queryText: `SiteID:"${siteId}" AND contentclass:"STS_Site"`,
    selectProperties: ["Title","Path"],
    rowLimit: 1
  });
  return results.items[0]?.Title || null;

}

 