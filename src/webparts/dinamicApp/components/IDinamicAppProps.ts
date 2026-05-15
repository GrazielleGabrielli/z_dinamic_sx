import { DisplayMode } from '@microsoft/sp-core-library';
import { IDynamicViewConfig, TViewMode } from '../core/config/types';
import { TPersistStatus } from '../core/persist/types';

export interface IDinamicPropertyPaneCommandHandlers {
  openWizard: () => void;
  openPageComponentsPicker: () => void;
  openFormManager: () => void;
}

export interface IDinamicAppProps {
  configJson: string;
  siteUrl: string;
  instanceScopeId: string;
  onSaveConfig: (config: IDynamicViewConfig) => void;
  persistStatus: TPersistStatus;
  forcedMode?: TViewMode;
  displayMode: DisplayMode;
  onRegisterPropertyPaneCommands: (handlers: IDinamicPropertyPaneCommandHandlers | undefined) => void;
  onCanManageListConfigChange: (can: boolean) => void;
}
