import { getSP } from '../../../services/core/sp';
import type {
  CreateValidationTemplateItemInput,
  ValidationTemplateItem
} from './validationTypes';

export const DEFAULT_VALIDATION_STATUS = 'Pendente';
export const DEFAULT_VALIDATION_TITLE = 'VALIDAR';
const VALIDATION_TEMPLATES_LIST_TITLE = 'ValidarTemplates';
const VALIDATION_TEMPLATE_SELECT_FIELDS = ['Id', 'ID', 'Title', 'TextoConsulta', 'Status', 'RespostaPBI', 'Modified', 'Created', 'GUID'];

const normalizeValidationTemplateItem = (item: Record<string, unknown>): ValidationTemplateItem => ({
  Id: Number(item.Id ?? item.ID ?? 0),
  ID: Number(item.ID ?? item.Id ?? 0),
  Title: String(item.Title ?? ''),
  TextoConsulta: String(item.TextoConsulta ?? ''),
  Status: String(item.Status ?? ''),
  RespostaPBI: String(item.RespostaPBI ?? ''),
  Modified: item.Modified ? String(item.Modified) : undefined,
  Created: item.Created ? String(item.Created) : undefined,
  GUID: item.GUID ? String(item.GUID) : undefined
});

export async function createValidationTemplateItem(
  data: CreateValidationTemplateItemInput
): Promise<ValidationTemplateItem> {
  const sp = getSP();

  if (!sp) {
    throw new Error('Contexto do SharePoint nao inicializado.');
  }

  try {
    const createdItemResult = await sp.web.lists.getByTitle(VALIDATION_TEMPLATES_LIST_TITLE).items.add({
      Title: data.Title,
      TextoConsulta: data.TextoConsulta,
      Status: data.Status
    });

    const rawData = (createdItemResult as { data?: Record<string, unknown> }).data;

    if (rawData) {
      return normalizeValidationTemplateItem(rawData);
    }

    const itemQueryable = (createdItemResult as { item?: { select: (...fields: string[]) => () => Promise<Record<string, unknown>> } }).item;

    if (itemQueryable) {
      const fetchedItem = await itemQueryable.select(...VALIDATION_TEMPLATE_SELECT_FIELDS)();
      return normalizeValidationTemplateItem(fetchedItem);
    }

    return normalizeValidationTemplateItem(createdItemResult as unknown as Record<string, unknown>);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Falha ao criar item na lista ValidarTemplates.';
    throw new Error(message);
  }
}

export async function getValidationTemplateItemById(itemId: number): Promise<ValidationTemplateItem> {
  const sp = getSP();

  if (!sp) {
    throw new Error('Contexto do SharePoint nao inicializado.');
  }

  try {
    const item = await sp.web.lists
      .getByTitle(VALIDATION_TEMPLATES_LIST_TITLE)
      .items
      .getById(itemId)
      .select(...VALIDATION_TEMPLATE_SELECT_FIELDS)();

    return normalizeValidationTemplateItem(item as Record<string, unknown>);
  } catch (error) {
    const message = error instanceof Error ? error.message : 'Falha ao consultar item na lista ValidarTemplates.';
    throw new Error(message);
  }
}
