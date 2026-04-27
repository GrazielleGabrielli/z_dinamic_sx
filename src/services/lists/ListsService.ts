import { getSPForWeb } from '../core/sp';
import { IListMetadata, IListSummary, ListSourceType } from './types';

const LIST_SELECT = 'Id,Title,BaseTemplate,ItemCount,Description,LastItemModifiedDate,EnableVersioning,Hidden';

// BaseTemplate 101 = Document Library
const LIBRARY_BASE_TEMPLATE = 101;

export class ListsService {
  private webSp(webServerRelativeUrl?: string) {
    return getSPForWeb(webServerRelativeUrl);
  }

  async getLists(includeHidden = false, webServerRelativeUrl?: string): Promise<IListSummary[]> {
    try {
      const sp = this.webSp(webServerRelativeUrl);
      let query = sp.web.lists
        .select('Id', 'Title', 'BaseTemplate', 'Hidden', 'ItemCount');

      if (!includeHidden) {
        query = query.filter('Hidden eq false');
      }

      const lists = await query() as unknown as Record<string, unknown>[];
      return lists.map((l) => ({
        ...l,
        IsLibrary: l['BaseTemplate'] === LIBRARY_BASE_TEMPLATE,
      })) as IListSummary[];
    } catch (e) {
      throw new Error(`ListsService.getLists: ${e}`);
    }
  }

  async getListById(listId: string, webServerRelativeUrl?: string): Promise<IListMetadata> {
    try {
      const sp = this.webSp(webServerRelativeUrl);
      const list = await sp.web.lists
        .getById(listId)
        .select(LIST_SELECT)();
      return { ...list, IsLibrary: list['BaseTemplate'] === LIBRARY_BASE_TEMPLATE } as IListMetadata;
    } catch (e) {
      throw new Error(`ListsService.getListById("${listId}"): ${e}`);
    }
  }

  async getListByTitle(title: string, webServerRelativeUrl?: string): Promise<IListMetadata> {
    try {
      const sp = this.webSp(webServerRelativeUrl);
      const list = await sp.web.lists
        .getByTitle(title)
        .select(LIST_SELECT)();
      return { ...list, IsLibrary: list['BaseTemplate'] === LIBRARY_BASE_TEMPLATE } as IListMetadata;
    } catch (e) {
      throw new Error(`ListsService.getListByTitle("${title}"): ${e}`);
    }
  }

  async detectListType(titleOrId: string, webServerRelativeUrl?: string): Promise<ListSourceType> {
    try {
      const sp = this.webSp(webServerRelativeUrl);
      const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
      const baseQuery = isGuid
        ? sp.web.lists.getById(titleOrId)
        : sp.web.lists.getByTitle(titleOrId);

      const list = await baseQuery.select('BaseTemplate')();

      if (list['BaseTemplate'] === LIBRARY_BASE_TEMPLATE) return 'library';
      if (list['BaseTemplate'] === 100) return 'list';
      return 'unknown';
    } catch (e) {
      return 'unknown';
    }
  }

  async getListMetadata(titleOrId: string, webServerRelativeUrl?: string): Promise<IListMetadata> {
    const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
    return isGuid
      ? this.getListById(titleOrId, webServerRelativeUrl)
      : this.getListByTitle(titleOrId, webServerRelativeUrl);
  }
}
