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
import SearchPagination from "../Pagination/SearchPagination";

import { reducer, State } from "../../../../services/searchReducer";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import {
  REFINER_CONFIG,
  buildRefinementFilters,
  PAGE_SIZE,
} from "../../../../services/searchUtils";

export interface ISearchPageProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: BaseComponentContext;
  scopeFilters: string[];
  initialQuery: string | null;
  scopeLabel?: string | null;
  isSiteScoped?: boolean;
}

const initialState: State = {
  query: "",
  loading: false,
  results: [],
  refiners: [],
  error: null,
  currentPage: 1,
  totalRows: 0,
};
/**
 * Root search page component.
 */
const SearchPage: React.FC<ISearchPageProps> = (props) => {
  const { context, scopeFilters, scopeLabel, isSiteScoped, initialQuery } = props;
  const [state, dispatch] = useReducer(reducer, initialState);//Reducer state for managing search state
  const [activeFilters, setActiveFilters] = useState<Record<string, string[]>>({});//state for managing active filters selected by the user
  const refinerColumnsAsString = REFINER_CONFIG.map((cfg) => cfg.managedProperty).join(",");//string of refiner managed properties to pass to the search API
  const scopeFiltersRef = useRef<string[]>(scopeFilters);//ref to hold the scope filters (Hub or Child Site) so they don't change on re-renders

  const doSearch = useCallback(
    async (
      query: string,
      filters: Record<string, string[]>,  
      updateRefiners = true,
      page = 1,
    ) => {
      //query validation
      if (!query || query.trim().length === 0) {
        dispatch({ type: "searchError", payload: "Please enter a search query." });
        return;
      }

      dispatch({ type: "searchStart" });//start search and set loading state to true
      dispatch({ type: "setPage", payload: page });//set the current page in the state

      try {
        // Merge the scope filters (Hub or Child Site) with the user-selected refiner filters.
        const mergedFilters: string[] = [
          ...scopeFiltersRef.current,
          ...buildRefinementFilters(filters),
        ];

        // Call the search service to perform the search with the query, pagination, and filters.
        const { items, refiners, totalRows } = await new SearchService(context).search({
          queryText: query,
          rowLimit: PAGE_SIZE,
          startRow: (page - 1) * PAGE_SIZE,
          refinementFilters: mergedFilters,
          refinerColumns: refinerColumnsAsString,
        });

        // Update the state with the search results, refiners (if requested), and total rows.
        dispatch({
          type: "searchSuccess",
          payload: { items, refiners: updateRefiners ? refiners : undefined, totalRows },
        });
      } catch (err: any) {
        dispatch({ type: "searchError", payload: err?.message ?? String(err) });
      }
    },
    [context],
  );

 
 // Auto-execute a search if a query was passed via the ?q= querystring (e.g. redirected from a child site).
  useEffect(() => {
    if (initialQuery) {
      dispatch({ type: "setQuery", payload: initialQuery });
      doSearch(initialQuery, {}).catch((err) => {
        dispatch({ type: "searchError", payload: err?.message ?? String(err) });
      });
    }
  }, [doSearch]);

  /** Triggered by the Search button or Enter key. Clears active user-selected filters and resets to page 1. */
  const onSearch = useCallback(() => {
    setActiveFilters({});
    doSearch(state.query, {}, true, 1).catch((err) => {
      dispatch({ type: "searchError", payload: err?.message ?? String(err) });
    });
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
        doSearch(state.query, newFilters, false, 1).catch((err) => {
          dispatch({ type: "searchError", payload: err?.message ?? String(err) });
        }); // reset to page 1 when refiner changes
       //compute 
        return newFilters;
      });
    },
    [state.query, doSearch],
  );

  // Show refiners sidebar only when there are loaded results with non-empty refiner buckets.
  const hasResults = !state.loading && state.results.length > 0;
  const hasRefiners = state.refiners.some((r) => r.Entries?.length > 0);
  const totalPages = Math.ceil(state.totalRows / PAGE_SIZE);

  return (
    <section className={`${styles.searchPage}`}>
      <Stack tokens={{ childrenGap: 12 }}>
        {scopeLabel && (
          <Text variant="smallPlus" styles={{ root: { color: "#605e5c" } }}>
            {scopeLabel}
            {isSiteScoped && (
              <>
                {" \u2014 "}
                <a
                  href={(() => {
                    const keyword = (state.query || initialQuery || "").trim();
                    return window.location.pathname + (keyword ? `?q=${encodeURIComponent(keyword)}` : "");
                  })()}
                  className={styles.link}
                >
                  Click here to show results from all sites
                </a>
              </>
            )}
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
              <Text variant="small" styles={{ root: { color: "#605e5c", marginBottom: 4 } }}>
                {`${state.totalRows} result${state.totalRows !== 1 ? "s" : ""} found`}
              </Text>
              <div className={styles.resultsArea}>
                <SearchResults items={state.results} />
              </div>
              <SearchPagination
                currentPage={state.currentPage}
                totalPages={totalPages}
                onPageChange={(page) => doSearch(state.query, activeFilters, false, page)}
              />
            </div>
          )}
        </div>
      </Stack>
    </section>
  );
};

export default SearchPage;
