import { ISearchResultItem } from "../models/searchModel";

export type State = {
  query: string;
  loading: boolean;
  results: ISearchResultItem[];
  refiners: any[];
  error?: string | null;
  currentPage: number;
  totalRows: number;
};

type Action =
  | { type: "setQuery"; payload: string }
  | { type: "searchStart" }
  | {
      type: "searchSuccess";
      payload: { items: ISearchResultItem[]; refiners: any[] | undefined; totalRows: number };
    }
  | { type: "searchError"; payload: string }
  | { type: "setPage"; payload: number };


export function reducer(state: State, action: Action): State {
  switch (action.type) {
    case "setQuery":
      return { ...state, query: action.payload };
    case "searchStart":
      return { ...state, loading: true, error: null };
    case "searchSuccess":
      return {
        ...state,
        loading: false,
        results: action.payload.items,
        refiners: action.payload.refiners ?? state.refiners,
        totalRows: action.payload.totalRows,
      };
    case "searchError":
      return {
        ...state,
        loading: false,
        error: action.payload,
        results: [],
        refiners: [],
        totalRows: 0,
        currentPage: 1,
      };
    case "setPage":
      return { ...state, currentPage: action.payload };
    default:
      return state;
  }
}