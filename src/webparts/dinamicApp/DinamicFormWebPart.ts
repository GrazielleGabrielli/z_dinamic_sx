import { DinamicWebPartBase } from './DinamicWebPartBase';
import { TViewMode } from './core/config/types';

export default class DinamicFormWebPart extends DinamicWebPartBase {
  protected getForcedMode(): TViewMode | undefined {
    return 'formManager';
  }
}
