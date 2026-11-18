import type {SetupCommandState} from './blindtype-native-commands';
import {ownedSetupHub} from './blindtype-native-hub';

export const SETUP_COMMAND_CHANNEL = 'cm:blind-type-setup:commands:v1';
export interface SetupCommandEnvelope {
    protocol: 1;
    component: 'CMBlindTypeSetupGridV2';
    controlName: 'Subgrid_BlindTypeSetup';
    instanceId: string;
    recordHint?: string;
    loading: boolean;
    destroyed: boolean;
    payload: SetupCommandState;
}

// This channel belongs to our component and form adapter. It never reads the
// host application's DOM or private runtime and never leaves the current page.
export function publishSetupCommand(container: HTMLElement, envelope: SetupCommandEnvelope, refresh?: () => void): void {
    if (!container.isConnected) return;
    const view=container.ownerDocument.defaultView;
    if(view)ownedSetupHub(view).publish(container,envelope,refresh);
    const event = container.ownerDocument.createEvent('CustomEvent');
    event.initCustomEvent(SETUP_COMMAND_CHANNEL, true, false, envelope);
    container.dispatchEvent(event);
}
