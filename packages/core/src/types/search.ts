import type { SharePointItem } from "./models";
import type { ListChildrenSort } from "./rest";

/** Tùy chọn GET Search REST (security trim theo user hiện tại). */
export interface SearchAccessibleOptions {
  /** Số dòng một trang — mặc định 30. */
  rowLimit?: number;
  /** Offset trang sau (StartRow). */
  startRow?: number;
  /** Sort → sortlist managed property. */
  sort?: ListChildrenSort;
  /** InternalName cột option đang hiện (map sang selectproperties). */
  fieldInternalNames?: readonly string[];
  signal?: AbortSignal;
}

/** Một trang kết quả Search — phân trang StartRow, không @odata.nextLink. */
export interface SearchAccessiblePage {
  items: SharePointItem[];
  /** StartRow trang tiếp; undefined khi hết. */
  nextStartRow?: number;
  totalRows: number;
}

/** Cell trong PrimaryQueryResult.RelevantResults.Table.Rows (odata=nometadata). */
export interface RestSearchCell {
  Key?: string;
  Value?: string | null;
  ValueType?: string;
}

export interface RestSearchRow {
  Cells?: RestSearchCell[];
}

/** JSON Search REST GET query — chỉ phần cần map. */
export interface RestSearchQueryResponse {
  PrimaryQueryResult?: {
    RelevantResults?: {
      RowCount?: number;
      TotalRows?: number;
      TotalRowsIncludingDuplicates?: number;
      Table?: {
        Rows?: RestSearchRow[];
      };
    };
  };
}
