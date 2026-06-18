import { ISearchResultItem } from "../models/searchModel";



export type State = {
  query: string;
  loading: boolean;
  results: ISearchResultItem[];
  refiners: any[];
  error?: string | null;
};

type Action =
  | { type: "setQuery"; payload: string }
  | { type: "searchStart" }
  | {
      type: "searchSuccess";
      payload: { items: ISearchResultItem[]; refiners: any[] | undefined };
    }
  | { type: "searchError"; payload: string };



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
      };
    case "searchError":
      return {
        ...state,
        loading: false,
        error: action.payload,
        results: [],
        refiners: [],
      };
    default:
      return state;
  }
}