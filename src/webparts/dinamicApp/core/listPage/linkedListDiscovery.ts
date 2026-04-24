import { FieldsService, ListsService } from '../../../../services';
import type { IFieldMetadata } from '../../../../services';
import { listGuidFromLookupListField, normListGuid } from './sharePointLookupListGuid';

export interface IDiscoveredLinkedList {
  listTitle: string;
  listId: string;
  lookupFields: { internalName: string; title: string }[];
}

export async function discoverListsWithLookupToMain(
  mainListTitleOrId: string,
  listsService: ListsService,
  fieldsService: FieldsService
): Promise<IDiscoveredLinkedList[]> {
  const mainMeta = await listsService.getListMetadata(mainListTitleOrId.trim());
  const primaryIdNorm = normListGuid(mainMeta.Id);
  const summaries = await listsService.getLists(false);
  const out: IDiscoveredLinkedList[] = [];
  for (let i = 0; i < summaries.length; i++) {
    const sum = summaries[i];
    if (sum.IsLibrary) continue;
    const title = String(sum.Title ?? '').trim();
    if (!title) continue;
    if (normListGuid(sum.Id) === primaryIdNorm) continue;
    let fields: IFieldMetadata[];
    try {
      fields = await fieldsService.getFields(title);
    } catch {
      continue;
    }
    const lk = fields.filter(
      (m) =>
        (m.MappedType === 'lookup' || m.MappedType === 'lookupmulti') &&
        Boolean(m.LookupList) &&
        listGuidFromLookupListField(m.LookupList) === primaryIdNorm
    );
    if (lk.length === 0) continue;
    out.push({
      listTitle: title,
      listId: sum.Id,
      lookupFields: lk.map((m) => ({ internalName: m.InternalName, title: m.Title })),
    });
  }
  out.sort((a, b) => a.listTitle.localeCompare(b.listTitle, undefined, { sensitivity: 'base' }));
  return out;
}
