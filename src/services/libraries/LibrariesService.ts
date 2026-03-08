import { getSP } from '../core/sp';
import { ILibraryMetadata, ILibrarySummary, IDocumentField } from './types';

const LIBRARY_BASE_TEMPLATE = 101;
const LIB_SELECT = 'Id,Title,BaseTemplate,ItemCount,Description,LastItemModifiedDate,EnableVersioning,Hidden';

export class LibrariesService {
  private get sp() { return getSP(); }

  async getLibraries(includeHidden = false): Promise<ILibrarySummary[]> {
    try {
      let query = this.sp.web.lists
        .select('Id', 'Title', 'BaseTemplate', 'ItemCount', 'Description', 'Hidden')
        .filter(`BaseTemplate eq ${LIBRARY_BASE_TEMPLATE}`);

      if (!includeHidden) {
        query = this.sp.web.lists
          .select('Id', 'Title', 'BaseTemplate', 'ItemCount', 'Description', 'Hidden')
          .filter(`BaseTemplate eq ${LIBRARY_BASE_TEMPLATE} and Hidden eq false`);
      }

      const libs = await query();
      return libs as ILibrarySummary[];
    } catch (e) {
      throw new Error(`LibrariesService.getLibraries: ${e}`);
    }
  }

  async getLibraryByTitle(title: string): Promise<ILibraryMetadata> {
    try {
      const lib = await this.sp.web.lists
        .getByTitle(title)
        .select(LIB_SELECT)();

      if (lib['BaseTemplate'] !== LIBRARY_BASE_TEMPLATE) {
        throw new Error(`"${title}" não é uma biblioteca de documentos`);
      }

      return { ...lib, IsLibrary: true } as ILibraryMetadata;
    } catch (e) {
      throw new Error(`LibrariesService.getLibraryByTitle("${title}"): ${e}`);
    }
  }

  async getLibraryMetadata(titleOrId: string): Promise<ILibraryMetadata> {
    try {
      const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
      const lib = await (isGuid
        ? this.sp.web.lists.getById(titleOrId)
        : this.sp.web.lists.getByTitle(titleOrId)
      ).select(LIB_SELECT)();

      return { ...lib, IsLibrary: true } as ILibraryMetadata;
    } catch (e) {
      throw new Error(`LibrariesService.getLibraryMetadata("${titleOrId}"): ${e}`);
    }
  }

  /** Retorna apenas campos relevantes para documentos (não sistema) */
  async getDefaultDocumentFields(titleOrId: string): Promise<IDocumentField[]> {
    try {
      const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
      const fields = await (isGuid
        ? this.sp.web.lists.getById(titleOrId)
        : this.sp.web.lists.getByTitle(titleOrId)
      ).fields
        .filter('Hidden eq false and ReadOnlyField eq false')
        .select('InternalName', 'Title', 'TypeAsString', 'Required')();

      return fields as IDocumentField[];
    } catch (e) {
      throw new Error(`LibrariesService.getDefaultDocumentFields("${titleOrId}"): ${e}`);
    }
  }
}
