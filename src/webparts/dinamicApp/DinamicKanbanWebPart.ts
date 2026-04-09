import { DinamicWebPartBase } from './DinamicWebPartBase';
import { TViewMode } from './core/config/types';

export default class DinamicKanbanWebPart extends DinamicWebPartBase {
  protected getForcedMode(): TViewMode | undefined {
    return 'projectManagement';
  }
}
