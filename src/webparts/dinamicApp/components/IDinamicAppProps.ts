import { IDynamicViewConfig } from '../core/config/types';

export interface IDinamicAppProps {
  configJson: string;
  siteUrl: string;
  instanceScopeId: string;
  onSaveConfig: (config: IDynamicViewConfig) => void;
}
