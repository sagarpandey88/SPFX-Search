import * as React from "react";
import { useReducer, useCallback, useState, useEffect, useRef } from "react";
import styles from "./SearchPage.module.scss";
import { SearchService } from "../../../../services/searchService";
import {
  Stack,
  SearchBox,
  PrimaryButton,
  Spinner,
  Text,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import SearchResults from "../SearchResults/SearchResults";
import SearchRefiners from "../Refiners/SearchRefiners";

import { reducer, State } from "../../../../services/searchReducer";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import {
  REFINER_CONFIG,
  buildRefinementFilters,
} from "../../../../services/searchUtils";

export interface ISearchPageProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: BaseComponentContext;
  scopeFilters: string[];
  scopeLabel: string | null;
  initialQuery: string | null;
}

const initialState: State = {
  query: "",
  loading: false,
  results: [],
  refiners: [],
  error: null,
};
/**
 * Root search page component.
 */
const SearchPage: React.FC<ISearchPageProps> = (props) => {
  const { context, scopeFilters, scopeLabel, initialQuery } = props;
  const [state, dispatch] = useReducer(reducer, initialState);
  const [activeFilters, setActiveFilters] = useState<Record<string, string[]>>({});
  const refinerColumnsAsString = REFINER_CONFIG.map((cfg) => cfg.managedProperty).join(",");
  const scopeFiltersRef = useRef<string[]>(scopeFilters);

  const doSearch = useCallback(    async (query: string, filters: Record<string, string[]>, updateRefiners = true) => {
     
    //query validation
      if (!query || query.trim().length === 0) {
        dispatch({
          type: "searchError",
          payload: "Please enter a search query.",
        });
        return;
      }

      // Start the search: show loading spinner and clear previous errors.
      dispatch({ type: "searchStart" });

      try {
        // Convert the map of active user filters to SharePoint refinement filter strings,
        // then merge with the pre-built scope filter strings (hub or specific site).
        const mergedFilters: string[] = [
          ...scopeFiltersRef.current,
          ...buildRefinementFilters(filters),
        ];

        // Call the SearchService with the current query and active filters.
        const { items, refiners } = await new SearchService(context).search({
          queryText: query,
          rowLimit: 10,
          refinementFilters: mergedFilters,
          refinerColumns: refinerColumnsAsString,
        });

        // On success, update results in state and also update refiners if this is a fresh search (not just a refiner toggle).
        dispatch({
          type: "searchSuccess",
          payload: { items, refiners: updateRefiners ? refiners : undefined },
        });
      } catch (err: any) {
        dispatch({ type: "searchError", payload: err?.message ?? String(err) }); // On error, show an error message and clear results.
      }
    },
    [context],
  );

 
 // Auto-execute a search if a query was passed via the ?q= querystring (e.g. redirected from a child site).
  useEffect(() => {
    if (initialQuery) {
      dispatch({ type: "setQuery", payload: initialQuery });
      doSearch(initialQuery, {});
    }
  }, [doSearch]);

  /** Triggered by the Search button or Enter key. Clears active user-selected filters and runs a fresh search. */
  const onSearch = useCallback(() => {
    setActiveFilters({});
    doSearch(state.query, {}); // user refiner filters are cleared on a fresh search; scope filter is merged inside doSearch.
  }, [state.query, doSearch]);

  /**Called when a refiner checkbox is toggled.*/
  const onRefinerChange = useCallback(    (managedProp: string, token: string, checked: boolean) => {
     console.log(`Refiner change: ${managedProp} - ${token} - ${checked}`);
    setActiveFilters((prev) => {
        const current = prev[managedProp] || [];
        const updated = checked
          ? [...current, token]
          : current.filter((t) => t !== token);
        const newFilters = { ...prev, [managedProp]: updated };
        console.log("Updated active filters:", newFilters);
        doSearch(state.query, newFilters, false); // When a refiner is toggled, we want to apply the new filters but keep the same refiner options, so we set updateRefiners to false.
       //compute 
        return newFilters;
      });
    },
    [state.query, doSearch],
  );

  // Show refiners sidebar only when there are loaded results with non-empty refiner buckets.
  const hasResults = !state.loading && state.results.length > 0;
  const hasRefiners = state.refiners.some((r) => r.Entries?.length > 0);

  return (
    <section className={`${styles.searchPage}`}>
      <Stack tokens={{ childrenGap: 12 }}>
        {scopeLabel && (
          <Text variant="smallPlus" styles={{ root: { color: "#605e5c" } }}>
            {scopeLabel}
          </Text>
        )}
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="end">
          <SearchBox
            placeholder="Search"
            value={state.query}
            onChange={(_, newValue) =>
              dispatch({ type: "setQuery", payload: newValue || "" })
            }
            onSearch={() => onSearch()}
            styles={{ root: { width: 480 } }}
          />
          <PrimaryButton text="Search" onClick={onSearch} />
        </Stack>

        {state.loading && <Spinner label="Searching..." />}

        {state.error && (
          <MessageBar messageBarType={MessageBarType.error}>
            {state.error}
          </MessageBar>
        )}

        {!state.loading && !state.error && state.results.length === 0 && (
          <Text variant="small">No results to display.</Text>
        )}
        <div style={{ display: "flex", flexDirection: "row", width: "100%" }}>
          {hasRefiners && (
            <div className={styles.sidebar}>
              <SearchRefiners
                refiners={state.refiners}
                activeFilters={activeFilters}
                onChange={onRefinerChange}
              />
            </div>
          )}
          {hasResults && (
            <div className={styles.resultContainer}>
              <div className={styles.resultsArea}>
                <SearchResults items={state.results} />
              </div>
            </div>
          )}
        </div>
      </Stack>
    </section>
  );
};

export default SearchPage;
