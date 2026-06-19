import * as React from "react";
import styles from "./SearchResults.module.scss";
import { ISearchResultItem } from "../../../../models/searchModel";

type ISearchResultsProps = {
  items: ISearchResultItem[];//any
};

const SearchResults: React.FC<ISearchResultsProps> = ({ items }) => {
  if (!items || items.length === 0) return null;

  return (
    <div className={styles.results}>
      {items.map((it) => (
        <div className={styles.item} key={it.key}>
          <div className={styles.thumbnail}>
            {/* Show thumbnail image when available, otherwise a neutral placeholder SVG. */}
            {it.Thumbnail ? (
              <img src={it.Thumbnail} alt={it.Title} />
            ) : (
              <svg width="48" height="36" viewBox="0 0 48 36" fill="none" xmlns="http://www.w3.org/2000/svg">
                <rect width="48" height="36" fill="#e1dfdd" />
              </svg>
            )}
          </div>
          <div className={styles.content}>
            {/* Title links to the item's SharePoint URL, opening in a new tab. */}
            <div className={styles.title}>
              <a href={it.Path || "#"} target="_blank" rel="noopener noreferrer">{it.Title}</a>
            </div>
            {it.Snippet && <div className={styles.snippet} dangerouslySetInnerHTML={{ __html: it.Snippet }} />}
            {/* Site name/URL line — shown below the snippet, above the meta line. */}
            {(it.SiteName || it.SiteUrl) && (
              <div className={styles.site}>
                In{" "}
                {it.SiteUrl
                  ? <a href={it.SiteUrl} target="_blank" rel="noopener noreferrer">{it.SiteName || it.SiteUrl}</a>
                  : <span>{it.SiteName}</span>}
              </div>
            )}
            {/* Meta line: only renders fields that have a value, separated by " · ". */}
            <div className={styles.meta}>
              {it.Author && <span>By {it.Author}</span>}
              {it.Modified && <span>{it.Author ? " · " : ""}Modified {it.Modified}</span>}
              {it.FileType && <span>{(it.Author || it.Modified) ? " · " : ""}{it.FileType}</span>}
              {it.Views && <span>{(it.Author || it.Modified || it.FileType) ? " · " : ""}{it.Views} views</span>}
            </div>
          </div>
        </div>
      ))}
    </div>
  );
};

export default SearchResults;
