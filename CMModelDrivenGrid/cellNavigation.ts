export interface CellPosition {row: number; col: number}
export interface CellNavigationOptions {rows: number; columns: number; pageSize?: number; rtl?: boolean; edge?: boolean; visibleRows?: readonly number[]}

/** Pure navigation shared by mouse-independent range selection and virtual-row focus. */
export function navigateCell(current: CellPosition, key: string, options: CellNavigationOptions): CellPosition | undefined {
    if (options.rows < 1 || options.columns < 1) return undefined;
    let {row, col} = current;
    const lastColumn = options.columns - 1;
    const visible = options.visibleRows;
    let rowPosition = visible ? visible.indexOf(row) : row;
    const lastRow = (visible?.length ?? options.rows) - 1;
    if (lastRow < 0 || rowPosition < 0) return undefined;
    const page = Math.max(1, options.pageSize ?? 10);
    switch (key) {
        case 'ArrowLeft': col = options.edge ? (options.rtl ? lastColumn : 0) : col + (options.rtl ? 1 : -1); break;
        case 'ArrowRight': col = options.edge ? (options.rtl ? 0 : lastColumn) : col + (options.rtl ? -1 : 1); break;
        case 'ArrowUp': rowPosition = options.edge ? 0 : rowPosition - 1; break;
        case 'ArrowDown': rowPosition = options.edge ? lastRow : rowPosition + 1; break;
        case 'PageUp': rowPosition -= page; break;
        case 'PageDown': rowPosition += page; break;
        case 'Home': col = 0; if (options.edge) rowPosition = 0; break;
        case 'End': col = lastColumn; if (options.edge) rowPosition = lastRow; break;
        default: return undefined;
    }
    rowPosition = Math.max(0, Math.min(lastRow, rowPosition));
    row = visible?.[rowPosition] ?? rowPosition;
    return {row, col: Math.max(0, Math.min(lastColumn, col))};
}
