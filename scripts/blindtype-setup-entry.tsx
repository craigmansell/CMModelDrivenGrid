import * as React from 'react';
import * as ReactDOM from 'react-dom';
import {initializeIcons} from '@fluentui/react/lib/Icons';
import {BlindTypeSetup} from './blindtype-setup';
import {dataverseStore} from './blindtype-setup-store';

initializeIcons(undefined, {disableWarnings: true});
try {
    const parameters = new URLSearchParams(window.location.search);
    const id = (parameters.get('data') ?? parameters.get('id') ?? '').replace(/[{}]/g, '').toLowerCase();
    ReactDOM.render(<BlindTypeSetup store={dataverseStore(window.location.origin, id)}/>, document.getElementById('setup-root'));
} catch (error) {
    ReactDOM.render(<main><h1>Blind type setup</h1><p role="alert">{error instanceof Error ? error.message : 'The blind type could not be opened.'}</p></main>, document.getElementById('setup-root'));
}
