import * as React from 'react';
import { HierarchyGrid, HierarchyGridProps } from './HierarchyGrid';

export interface ResponsiveGridProps extends HierarchyGridProps {
    /** Optional embedded-subgrid fallback; a definite Power Apps allocation wins. */
    fixedHeight?: number;
    /** Compact embedded grids; unconfigured/full-tab controls retain fill sizing. */
    fitToRows?: boolean;
}

export function resolveGridHeight(allocated?: number, fixed?: number, measured?: number): number {
    for (const value of [allocated, fixed, measured]) {
        if (typeof value === 'number' && Number.isFinite(value) && value > 0) return value;
    }
    return -1;
}

/** Fill the supplied tab/container without inspecting or modifying Dynamics DOM. */
export function measureGridContentHeight(host: HTMLElement): number | undefined {
    const layout = host.querySelector<HTMLElement>('.cm-grid-layout');
    const details = host.querySelector<HTMLElement>('.cm-grid-details');
    const grid = host.querySelector<HTMLElement>('.cm-model-driven-grid');
    if (!layout || !details || !grid) return undefined;
    const pixels = (value: string): number => Number.parseFloat(value) || 0;
    const chrome = Array.from(layout.children).reduce((sum, element) => {
        if (element.classList.contains('cm-grid-surface')) return sum;
        const style = window.getComputedStyle(element);
        return sum + element.getBoundingClientRect().height + pixels(style.marginTop) + pixels(style.marginBottom);
    }, 0);
    const scroll = host.querySelector<HTMLElement>('[data-is-scrollable="true"]');
    const scrollbar = scroll ? Math.max(0, scroll.offsetHeight - scroll.clientHeight) : 0;
    const style = window.getComputedStyle(grid);
    const content = Math.max(details.getBoundingClientRect().height, host.querySelector('.noRecords') ? 160 : 0);
    if (content <= 0) return undefined;
    return Math.ceil(content + chrome + scrollbar + pixels(style.borderTopWidth) + pixels(style.borderBottomWidth) + 1);
}

export function ResponsiveGrid({ fixedHeight, fitToRows, ...props }: ResponsiveGridProps): React.ReactElement {
    const host = React.useRef<HTMLDivElement>(null);
    const [measuredHeight, setMeasuredHeight] = React.useState<number>();
    const [compactHeight, setCompactHeight] = React.useState<number>();
    const requestedHeight = resolveGridHeight(props.height, fixedHeight);
    const automatic = requestedHeight < 0;

    React.useLayoutEffect(() => {
        if (!fitToRows) return;
        const element = host.current;
        if (!element) return;
        let frame: number | undefined;
        let disposed = false;
        const observed = new Set<Element>();
        const measure = (): void => {
            frame = undefined;
            if (disposed) return;
            const visible = window.innerHeight - element.getBoundingClientRect().top;
            const maximum = requestedHeight > 0 ? requestedHeight : Math.max(1, Math.floor(visible > 0 ? visible : window.innerHeight));
            const content = measureGridContentHeight(element);
            // Keep the scrolling viewport while more pages are available;
            // shrinking around its sentinel would fetch every page at once.
            const next = !props.itemsLoading && !props.hasNextPage && content ? Math.min(maximum, content) : maximum;
            setCompactHeight(current => current === next ? current : next);
        };
        const schedule = (): void => {
            if (!disposed && frame === undefined) frame = window.requestAnimationFrame(measure);
        };
        const observer = typeof ResizeObserver === 'undefined' ? undefined : new ResizeObserver(schedule);
        const observeContent = (): void => {
            const elements = new Set<Element>([element, ...Array.from(element.querySelectorAll('.cm-grid-details, .cm-grid-layout > :not(.cm-grid-surface)'))]);
            for (const old of observed) if (!elements.has(old)) { observer?.unobserve(old); observed.delete(old); }
            for (const next of elements) if (!observed.has(next)) { observer?.observe(next); observed.add(next); }
        };
        observeContent();
        const mutations = typeof MutationObserver === 'undefined' ? undefined : new MutationObserver(() => { observeContent(); schedule(); });
        mutations?.observe(element, { childList: true, subtree: true });
        window.addEventListener('resize', schedule);
        measure();
        return () => {
            disposed = true;
            observer?.disconnect(); mutations?.disconnect();
            window.removeEventListener('resize', schedule);
            if (frame !== undefined) window.cancelAnimationFrame(frame);
        };
    }, [fitToRows, requestedHeight, props.hasNextPage, props.itemsLoading]);

    React.useLayoutEffect(() => {
        if (!automatic || fitToRows) return;
        const element = host.current;
        if (!element) return;
        let disposed = false;
        let frame: number | undefined;
        const measure = (): void => {
            if (disposed) return;
            const available = element.clientHeight;
            const top = element.getBoundingClientRect().top;
            const visible = window.innerHeight - top;
            // Native full-tab subgrids can include command-bar space in their
            // host height. Keep our own content within the visible work area.
            const next = available > 0
                ? top >= 0 && visible > 0 ? Math.min(available, Math.max(1, Math.floor(visible))) : available
                : undefined;
            setMeasuredHeight(current => current === next ? current : next);
        };
        const schedule = (): void => {
            if (disposed || frame !== undefined) return;
            frame = window.requestAnimationFrame(() => { frame = undefined; measure(); });
        };
        measure();
        const observer = typeof ResizeObserver === 'undefined' ? undefined : new ResizeObserver(schedule);
        observer?.observe(element);
        window.addEventListener('resize', schedule);
        return () => {
            disposed = true;
            observer?.disconnect();
            window.removeEventListener('resize', schedule);
            if (frame !== undefined) window.cancelAnimationFrame(frame);
        };
    }, [automatic, fitToRows]);

    const height = fitToRows ? compactHeight ?? requestedHeight : automatic ? resolveGridHeight(undefined, undefined, measuredHeight) : requestedHeight;
    return <div ref={host} className="cm-grid-host" style={{
        height: fitToRows ? height > 0 ? height : 'auto' : automatic ? '100%' : requestedHeight, width: '100%', minHeight: 0, minWidth: 0,
    }}>
        <HierarchyGrid {...props} height={height}/>
    </div>;
}
