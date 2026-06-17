import * as React from "react";
import { useReducer, useCallback, useState } from "react";
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
import { ActiveFilters } from "../../../../models/searchModel";
import { reducer, State } from "../../../../services/searchReducer";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { REFINER_CONFIG } from "../../../../services/searchUtils";

export interface ISearchPageProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: BaseComponentContext;
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
 * Manages query input, result state, active refiner filters,
 * and lays out the sidebar refiners panel alongside the results list.
 */
const SearchPage: React.FC<ISearchPageProps> = (props) => {
  const { context } = props;
  const [state, dispatch] = useReducer(reducer, initialState); //state management with reducer pattern due to multiple related values (query, loading, results, refiners, error) that change together in response to search actions.
  const [activeFilters, setActiveFilters] = useState<ActiveFilters>({}); //Tracks which refiner filters are currently applied by the user, structured as a map of managed property names to arrays of selected tokens. Updated when users toggle refiner checkboxes, and used as input to the search function to apply refinements without refreshing the refiner options list.
  const refinerColumnsAsString = REFINER_CONFIG.map((cfg) => cfg.managedProperty).join(","); //Comma-separated list of refiner managed properties to request from the search API, derived from the REFINER_CONFIG. This is passed to the SearchService to ensure we get back the refiner buckets for the properties we want to display in the UI. When a user toggles a refiner checkbox, we want to keep the same list of refiners in the search results so the options don't refresh and lose their checked state, which is why this is defined outside of the doSearch function and not regenerated on every search.
  /**
   * Core search executor.
   * @param updateRefiners - When true (fresh search), replaces the refiner options
   *   in state with the new buckets returned by SharePoint.  When false (refiner
   *   checkbox click), the existing refiner options are preserved so the user can
   *   continue selecting/deselecting without the panel re-rendering.
   */
  const doSearch = useCallback(
    async (
      query: string,
      filters: ActiveFilters,
      updateRefiners = true,
      refinerFields?: string[],
    ) => {
      //query validation: prevent empty searches and show an error message instead of calling the search API.
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
        // Call the SearchService with the current query and active filters. The service will handle building the query and normalizing results.
        const { items, refiners } = await new SearchService(context).search({
          queryText: query, //from text input
          rowLimit: 10,
          refinementFilters: filters,
          refinerColumns: refinerColumnsAsString //if refinerFields is provided (when toggling a refiner), use it to keep the same refiner options; otherwise (fresh search) use the full list of refiners from config.
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

  /** Triggered by the Search button or Enter key. Clears active filters and runs a fresh search. */
  const onSearch = useCallback(() => {
    setActiveFilters({});
    doSearch(state.query, {}); //refiner filters are cleared when doing a fresh search, so we pass an empty object.
  }, [state.query, doSearch]);

  /**
   * Called when a refiner checkbox is toggled.
   * Updates activeFilters and re-runs the search with refinement applied,
   * but does NOT refresh the refiner option list (updateRefiners = false).
   */
  const onRefinerChange = useCallback(
    (managedProp: string, token: string, checked: boolean) => {
      setActiveFilters((prev) => {
        const current = prev[managedProp] || [];
        const updated = checked
          ? [...current, token]
          : current.filter((t) => t !== token);
        const newFilters = { ...prev, [managedProp]: updated };
        doSearch(state.query, newFilters, false); // When a refiner is toggled, we want to apply the new filters but keep the same refiner options, so we set updateRefiners to false.
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
