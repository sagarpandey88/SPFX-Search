/** Describes one refiner group: what to show in the UI and which managed property backs it. */
export interface RefinerConfig {
  displayName: string;
  managedProperty: string;
  isTaxonomy?: boolean;
}

/** Maps managed property names to human-readable refiner panel labels. */
export const REFINER_CONFIG: RefinerConfig[] = [
  { displayName: "Author", managedProperty: "Author" },
  { displayName: "Category", managedProperty: "RefinableString99", isTaxonomy: true },
  { displayName: "Country", managedProperty: "RefinableString00", isTaxonomy: true },
];

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