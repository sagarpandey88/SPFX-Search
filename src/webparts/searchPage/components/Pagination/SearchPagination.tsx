import * as React from "react";
import { IconButton, Stack, Text } from "@fluentui/react";

export interface ISearchPaginationProps {
  currentPage: number;
  totalPages: number;
  onPageChange: (page: number) => void;
}

/**
 * Renders Previous / Next navigation buttons with a "Page X of Y" label.
 * Calls onPageChange with the new page number — the parent is responsible
 * for triggering the search API call.
 */
const SearchPagination: React.FC<ISearchPaginationProps> = ({
  currentPage,
  totalPages,
  onPageChange,
}) => {
  if (totalPages <= 1) return null;

  return (
    <Stack horizontal verticalAlign="center" tokens={{ childrenGap: 8 }}>
      <IconButton
        iconProps={{ iconName: "ChevronLeft" }}
        title="Previous page"
        ariaLabel="Previous page"
        disabled={currentPage <= 1}
        onClick={() => onPageChange(currentPage - 1)}
      />
      <Text variant="small">{`Page ${currentPage} of ${totalPages}`}</Text>
      <IconButton
        iconProps={{ iconName: "ChevronRight" }}
        title="Next page"
        ariaLabel="Next page"
        disabled={currentPage >= totalPages}
        onClick={() => onPageChange(currentPage + 1)}
      />
    </Stack>
  );
};

export default SearchPagination;
