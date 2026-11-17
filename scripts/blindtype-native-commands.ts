export type SetupCommand = 'new' | 'edit' | 'delete' | 'activate' | 'deactivate' | 'refresh';
export interface SetupCommandState {
    recordId?: string;
    bindRecord?: (recordId: string) => void;
    ready: boolean; selected: string[]; canNew: boolean; canEdit: boolean; canDelete: boolean; canActivate: boolean; canDeactivate: boolean; canRefresh: boolean;
    execute: (command: SetupCommand) => void;
}
export interface NativeSetupHost {
    publish: (state: SetupCommandState) => void;
    select: (sourceIds: string[], childSourceIds: string[]) => Promise<void>;
    refreshed: () => void;
}
