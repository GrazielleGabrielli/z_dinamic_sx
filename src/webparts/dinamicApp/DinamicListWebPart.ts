import { DinamicWebPartBase } from './DinamicWebPartBase';
import { TViewMode } from './core/config/types';

export default class DinamicListWebPart extends DinamicWebPartBase {
  protected getForcedMode(): TViewMode | undefined {
    return 'list';
  }
}
