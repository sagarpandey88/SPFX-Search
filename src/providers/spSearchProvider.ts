/* eslint-disable */
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { SPFx, spfi, SPFI, IRefiner } from "@pnp/sp/presets/all";
import { ISearchQueryOptions } from "../models/searchModel";


export class SearchProvider {
  private context: BaseComponentContext;
  private sp: SPFI;

  /**
   * Constructor to initialize the SearchProvider with SPFx context.
   * @param context - The SPFx context of the web part or component.
   */
  constructor(context: BaseComponentContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(this.context));
  }

  public async searchQuery({
    queryText,
    queryTemplate,
    sortBy,
    sortOrder,
    selectProperties,
    refinementFilters,
    refinerColumns,
    rowLimit,
    startRow,
    enableQueryRules,
    multiColumnSortBy,
    cacheEnabled = false    
  }: ISearchQueryOptions): Promise<{ rows: any[]; refiners: IRefiner[] | undefined }> {
    try {
      let sp = this.sp;

    

      // Construct the search query payload
      const searchPayload: any = {
        Querytext: queryText,
        QueryTemplate: queryTemplate,
        SortList: multiColumnSortBy ? multiColumnSortBy :  sortBy ? [{ Property: sortBy, Direction: sortOrder?.toLowerCase() === "descending" ? 1 : 0 }] : undefined,
        SelectProperties: selectProperties,
        RefinementFilters: refinementFilters,
        RowLimit: rowLimit,
        StartRow: startRow,
        EnableQueryRules: enableQueryRules,
        Refiners:refinerColumns,
        EnableStemming: false,
        EnablePhonetic: false,
         EnableNicknames: false,
      
      };

      // Perform the search query using PnPjs
      const data = await sp.search(searchPayload);

      // Extract result rows and refiners
      const rows = data.PrimarySearchResults;      
      const refiners = data.RawSearchResults.PrimaryQueryResult?.RefinementResults ?
        data.RawSearchResults.PrimaryQueryResult?.RefinementResults.Refiners : [];

      return { rows, refiners };
    } catch (error) {
      console.error("Error performing search query:", error);
      throw error;
    }
  }

  /**
   * Performs a search query in SharePoint using query parameters as a string.
   * @param queryParams - The search query parameters as a string.
   * @param cacheEnabled - Optional parameter to enable caching.
   * @returns A Promise resolving to an object containing result rows and refiners.
   */
  public async searchQueryWithParams(queryParams: string, cacheEnabled: boolean = false): Promise<{ rows: any[]; refiners: IRefiner[] | undefined }> {
    try {
      let sp = this.sp;
 

      // Perform the search query using PnPjs with query parameters
      const data = await sp.search(queryParams);

      // Extract result rows and refiners
      const rows = data.PrimarySearchResults;
      const refiners = data.RawSearchResults.PrimaryQueryResult?.RefinementResults
        ? data.RawSearchResults.PrimaryQueryResult?.RefinementResults.Refiners : [];

      return { rows, refiners };
    } catch (error) {
      console.error("Error performing search query:", error);
      throw error;
    }
  }
}