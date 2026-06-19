import * as React from "react";
import { useState } from "react";
import { Checkbox, Icon } from "@fluentui/react";
import styles from "./SearchRefiners.module.scss";
import {
  getProcessedEntries,
  REFINER_CONFIG,
} from "../../../../services/searchUtils";

interface ISearchRefinersProps {
  refiners: any[];
  activeFilters: { [managedProp: string]: string[] };
  onChange: (managedProp: string, token: string, checked: boolean) => void;
}

const SearchRefiners: React.FC<ISearchRefinersProps> = ({
  refiners,
  activeFilters,
  onChange,
}) => {
  const [collapsed, setCollapsed] = useState<{ [key: string]: boolean }>({}); // Tracks which groups are collapsed; empty = all expanded by default.

  /** Toggle collapsed state for the group identified by its managed property name. */
  const toggle = (name: string) =>
    setCollapsed((prev) => ({ ...prev, [name]: !prev[name] }));

  // Only show groups whose managed property returned at least one refiner bucket.
  const visibleGroups = REFINER_CONFIG.filter((cfg) => {
    const refiner = refiners.find((r) => r.Name === cfg.managedProperty);
    return refiner && refiner.Entries && refiner.Entries.length > 0;
  });

  if (visibleGroups.length === 0) return null;

  return (
    <div className={styles.panel}>
      <div className={styles.panelHeader}>Filter Results</div>
      {visibleGroups.map((cfg) => {
        const refiner = refiners.find((r) => r.Name === cfg.managedProperty);
        const isCollapsed = !!collapsed[cfg.managedProperty];
        const selected = activeFilters[cfg.managedProperty] || [];
        const displayEntries = getProcessedEntries(
          refiner.Entries,
          cfg.isTaxonomy,
        ); //clean the refiners as taxonmoy refiners have extra GUID.

        return (
          <div key={cfg.managedProperty} className={styles.group}>
            {/* Collapsible group header button */}
            <button
              className={styles.groupHeader}
              onClick={() => toggle(cfg.managedProperty)}
              aria-expanded={!isCollapsed}
            >
              <span>{cfg.displayName}</span>
              <Icon
                iconName={isCollapsed ? "ChevronRight" : "ChevronDown"}
                className={styles.chevron}
              />
            </button>
            {!isCollapsed && (
              <div className={styles.entries}>
                {displayEntries.map((entry: any) => (
                  <div key={entry.token} className={styles.entry}>
                    {/* Each checkbox represents one refiner bucket (value + count). */}
                    <Checkbox
                      label={`${entry.label} (${entry.count})`}
                      checked={selected.includes(entry.token)}
                      onChange={(_, chk) =>
                        onChange(cfg.managedProperty, entry.token, !!chk)
                      }
                    />
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
};

export default SearchRefiners;
