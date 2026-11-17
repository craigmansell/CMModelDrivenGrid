import * as React from 'react';
import {SetupComponent} from './blindtype-setup-model';

export const componentLabel = (component: SetupComponent): string => `${component.code} — ${component.name}`;
export function resolveComponent(choices: SetupComponent[], text: string): SetupComponent {
    const value = text.trim().toLocaleLowerCase('en-GB');
    const matches = choices.filter(component => [component.code, component.name, componentLabel(component)].some(label => label.toLocaleLowerCase('en-GB') === value));
    if (matches.length !== 1) throw new Error('Choose a component from the available matches.');
    return matches[0];
}
export function ComponentEntryRow({id, choices, initial, allowDefault, busy, onChange, onSave, onCancel}: {
    id: string; choices: SetupComponent[]; initial: {text: string; makeDefault: boolean}; allowDefault: boolean; busy: boolean;
    onChange: (value: {text: string; makeDefault: boolean}) => void; onSave: (text: string, makeDefault: boolean) => void; onCancel: () => void;
}): React.ReactElement {
    const [value, setValue] = React.useState(initial);
    const current = React.useRef(initial);
    const update = (patch: Partial<typeof value>): void => {const next = {...current.current, ...patch}; current.current = next; setValue(next); onChange(next);};
    return <form className="setup-component-entry" data-selection-disabled={true} onClick={event => event.stopPropagation()} onKeyDown={event => {event.stopPropagation(); if (event.key === 'Escape' && !busy) {event.preventDefault(); onCancel();}}} onSubmit={event => {event.preventDefault(); if (!busy) onSave(current.current.text, current.current.makeDefault);}}>
        <input aria-label="Component code or description" placeholder="Type a component code or description" list={`${id}-choices`} autoComplete="off" autoFocus disabled={busy} value={value.text} onChange={event => update({text: event.target.value})}/>
        <datalist id={`${id}-choices`}>{choices.map(component => <option key={component.id} value={componentLabel(component)}/>)}</datalist>
        {allowDefault && <label><input type="checkbox" checked={value.makeDefault} disabled={busy} onChange={event => update({makeDefault: event.target.checked})}/>Default</label>}
        <button type="submit" disabled={busy || !value.text.trim()}>{busy ? 'Saving…' : 'Save'}</button><button type="button" disabled={busy} onClick={onCancel}>Cancel</button>
    </form>;
}
