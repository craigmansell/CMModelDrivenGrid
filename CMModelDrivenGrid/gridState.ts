import * as React from 'react';

/** Data/UI state survives virtual row unmounts, without retaining off-screen DOM. */
export type GridStateCache = Map<string, unknown>;

export function useGridState<T>(cache: GridStateCache | undefined, key: string, initial: () => T): [T, React.Dispatch<React.SetStateAction<T>>] {
    const [value, setValue] = React.useState<T>(() => {
        if (cache?.has(key)) return cache.get(key) as T;
        const value = initial();
        cache?.set(key, value);
        return value;
    });
    const current = React.useRef(value);
    const mounted = React.useRef(true);
    current.current = value;
    React.useEffect(() => { mounted.current = true; return () => {mounted.current = false;}; }, []);
    const update = React.useCallback((action: React.SetStateAction<T>) => {
        const next = typeof action === 'function' ? (action as (current: T) => T)(current.current) : action;
        current.current = next;
        cache?.set(key, next);
        if (mounted.current) setValue(next);
    }, [cache, key]);
    return [value, update];
}
