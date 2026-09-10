// Match the scoped grid CSS: selection cell 40px; data-cell padding 12px per side.
const SELECTION_WIDTH = 40;
const CELL_PADDING = 24;

/** Fill spare viewport space without shrinking preferred or manually resized widths. */
export function fillColumnWidths(preferred: readonly number[], viewportWidth: number, manuallySized: readonly boolean[] = []): number[] {
	const widths = preferred.map((value) => Math.max(80, Number.isFinite(value) ? value : 80));
	if (!widths.length || !Number.isFinite(viewportWidth) || viewportWidth <= 0) return widths;
	const total = widths.reduce((sum, value) => sum + value, 0);
	const extra = Math.floor(viewportWidth - SELECTION_WIDTH - CELL_PADDING * widths.length - total);
	if (extra <= 0) return widths;
	const flexible = widths.map((_value, index) => index).filter((index) => !manuallySized[index]);
	// If every column was resized, let the trailing column absorb remaining space.
	if (!flexible.length) flexible.push(widths.length - 1);
	const weight = flexible.reduce((sum, index) => sum + widths[index], 0);
	let allocated = 0;
	let cumulativeWeight = 0;
	flexible.forEach((index, position) => {
		cumulativeWeight += widths[index];
		const nextAllocated = position === flexible.length - 1 ? extra : Math.floor(extra * cumulativeWeight / weight);
		widths[index] += nextAllocated - allocated;
		allocated = nextAllocated;
	});
	return widths;
}
