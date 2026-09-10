/** Excel-compatible tab-separated cells. Quote delimiters instead of changing the rectangle. */
export function serializeGridClipboard(rows: readonly (readonly string[])[]): string {
    return rows.map(row => row.map(value => /[\t\r\n"]/.test(value)
        ? `"${value.replace(/"/g, '""')}"` : value).join('\t')).join('\r\n');
}
