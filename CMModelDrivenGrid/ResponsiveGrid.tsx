import * as React from 'react';
import { HierarchyGrid, HierarchyGridProps } from './HierarchyGrid';

export interface ResponsiveGridProps extends HierarchyGridProps {
    /** Optional embedded-subgrid fallback; a definite Power Apps allocation wins. */
    fixedHeight?: number;
}

export function resolveGridHeight(allocated?: number, fixed?: number, measured?: number): number {
    for (const value of [allocated, fixed, measured]) {
        if (typeof value === 'number' && Number.isFinite(value) && value > 0) return value;
    }
    return -1;
}

/** Fill the supplied tab/container without inspecting or modifying Dynamics DOM. */
export function ResponsiveGrid({ fixedHeight, ...props }: ResponsiveGridProps): React.ReactElement {
    const host = React.useRef<HTMLDivElement>(null);
    const [measuredHeight, setMeasuredHeight] = React.useState<number>();
    const requestedHeight = resolveGridHeight(props.height, fixedHeight);
    const automatic = requestedHeight < 0;

    React.useLayoutEffect(() => {
        if (!automatic) return;
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
    }, [automatic]);

    const height = automatic ? resolveGridHeight(undefined, undefined, measuredHeight) : requestedHeight;
    return <div ref={host} className="cm-grid-host" style={{
        height: automatic ? '100%' : requestedHeight, width: '100%', minHeight: 0, minWidth: 0,
    }}>
        <HierarchyGrid {...props} height={height}/>
    </div>;
}
