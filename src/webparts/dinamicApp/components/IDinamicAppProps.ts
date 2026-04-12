import { IDynamicViewConfig, TViewMode } from '../core/config/types';

export interface IDinamicAppProps {
  configJson: string;
  siteUrl: string;
  instanceScopeId: string;
  onSaveConfig: (config: IDynamicViewConfig) => void;
  openAiApiKey?: string;
  forcedMode?: TViewMode;
}
