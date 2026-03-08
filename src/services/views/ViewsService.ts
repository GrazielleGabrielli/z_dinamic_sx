import { getSP } from '../core/sp';
import { IViewMetadata } from './types';

const VIEW_SELECT = 'Id,Title,DefaultView,Hidden,PersonalView,RowLimit,ViewQuery';

const listRef = (sp: ReturnType<typeof getSP>, titleOrId: string) => {
  const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
  return isGuid
    ? sp.web.lists.getById(titleOrId)
    : sp.web.lists.getByTitle(titleOrId);
};

export class ViewsService {
  private get sp() { return getSP(); }

  async getViews(listTitleOrId: string): Promise<IViewMetadata[]> {
    try {
      const views = await listRef(this.sp, listTitleOrId).views
        .select(VIEW_SELECT)();
      const result: IViewMetadata[] = [];

      for (const v of views) {
        const fields = await listRef(this.sp, listTitleOrId).views
          .getById(v['Id']).fields();
        result.push({
          ...v,
          ViewFields: fields.Items ?? [],
        } as IViewMetadata);
      }

      return result;
    } catch (e) {
      throw new Error(`ViewsService.getViews("${listTitleOrId}"): ${e}`);
    }
  }

  async getViewById(listTitleOrId: string, viewId: string): Promise<IViewMetadata> {
    try {
      const view = await listRef(this.sp, listTitleOrId).views
        .getById(viewId)
        .select(VIEW_SELECT)();
      const fields = await listRef(this.sp, listTitleOrId).views
        .getById(viewId).fields();
      return { ...view, ViewFields: fields.Items ?? [] } as IViewMetadata;
    } catch (e) {
      throw new Error(`ViewsService.getViewById("${listTitleOrId}", "${viewId}"): ${e}`);
    }
  }

  async getDefaultView(listTitleOrId: string): Promise<IViewMetadata> {
    try {
      const view = await listRef(this.sp, listTitleOrId).defaultView
        .select(VIEW_SELECT)();
      const fields = await listRef(this.sp, listTitleOrId).defaultView.fields();
      return { ...view, ViewFields: fields.Items ?? [] } as IViewMetadata;
    } catch (e) {
      throw new Error(`ViewsService.getDefaultView("${listTitleOrId}"): ${e}`);
    }
  }

  async getViewFields(listTitleOrId: string, viewId: string): Promise<string[]> {
    try {
      const fields = await listRef(this.sp, listTitleOrId).views
        .getById(viewId).fields();
      return fields.Items ?? [];
    } catch (e) {
      throw new Error(`ViewsService.getViewFields("${listTitleOrId}", "${viewId}"): ${e}`);
    }
  }
}
