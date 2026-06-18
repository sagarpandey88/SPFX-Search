import { SPFx, spfi } from "@pnp/sp/presets/all";

/** Number of results per page. Change this to adjust pagination globally. */
export const PAGE_SIZE = 5;

/** Describes one refiner group: what to show in the UI and which managed property backs it. */
export interface RefinerConfig {
  displayName: string;
  managedProperty: string;
  isTaxonomy?: boolean;
}

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
            "PublishingImage"
        ];
/** Maps managed property names to human-readable refiner panel labels. */
export const REFINER_CONFIG: RefinerConfig[] = [
  { displayName: "Author", managedProperty: "Author" },
  { displayName: "Category", managedProperty: "RefinableString99", isTaxonomy: false },
  { displayName: "Country", managedProperty: "RefinableString00", isTaxonomy: true },
];

/**
 * Reads a single query-string parameter from the current page URL.
 * Returns null if the parameter is absent or empty.
 */
export function getQueryParam(name: string): string | null {
  const params = new URLSearchParams(window.location.search);
  return params.get(name) || null;
}

    /**
 * Converts the activeFilters map into SharePoint RefinementFilter strings.
 * Single selection: "Author:token"
 * Multi-selection:  "Author:or(token1,token2)" — SharePoint OR syntax.
 */
export function buildRefinementFilters(filters: Record<string, string[]>): string[] {
        return Object.entries(filters)
            .filter(([, tokens]) => tokens.length > 0)
            .map(([name, tokens]) =>
                tokens.length === 1
                    ? `${name}:${tokens[0]}`
                    : `${name}:or(${tokens.join(",")})`
            );
          }
            
/**
 * Fetches the hub site ID (GUID) for the current site collection via PnPjs:
 * equivalent to GET /_api/site?$select=IsHubSite,HubSiteId
 *
 * Returns the HubSiteId GUID string when the site is associated with (or is)
 * a hub site, or null if it is not part of any hub.
 */
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

/** Cleaned structure optimized for UI rendering */
export interface IProcessedEntry {
  token: string;
  label: string;
  count: string;
}

/**
 * Filters and transforms raw SharePoint search entries.
 * Isolates L0 records for taxonomy types and strips down raw pipe string structures.
 */
export function getProcessedEntries(entries: any[], isTaxonomy?: boolean): IProcessedEntry[] {
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
      cleanLabel = parts[parts.length - 1]; // Pulls clean text like "Swiss"
    }

    return {
      token: entry.RefinementToken,
      label: cleanLabel,
      count: entry.RefinementCount
    };
  });
}