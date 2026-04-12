import { IDynamicViewConfig, TViewMode } from '../core/config/types';
import { TPersistStatus } from '../core/persist/types';

export interface IDinamicAppProps {
  configJson: string;
  siteUrl: string;
  instanceScopeId: string;
  onSaveConfig: (config: IDynamicViewConfig) => void;
  persistStatus: TPersistStatus;
  forcedMode?: TViewMode;
}
