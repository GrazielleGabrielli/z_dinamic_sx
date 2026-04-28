import type { SPFI } from '@pnp/sp';
import { getSPForWeb } from '../core/sp';
import { IFieldMetadata, FieldMappedType, IRawSPField } from './types';

export const SYSTEM_METADATA_FIELDS: IFieldMetadata[] = [
  {
    Id: 'system-created',
    Title: 'Criado',
    InternalName: 'Created',
    TypeAsString: 'DateTime',
    MappedType: 'datetime',
    Required: false,
    ReadOnlyField: true,
    Hidden: false,
    Description: '',
    DefaultValue: null,
  },
  {
    Id: 'system-modified',
    Title: 'Modificado',
    InternalName: 'Modified',
    TypeAsString: 'DateTime',
    MappedType: 'datetime',
    Required: false,
    ReadOnlyField: true,
    Hidden: false,
    Description: '',
    DefaultValue: null,
  },
  {
    Id: 'system-author',
    Title: 'Criado por',
    InternalName: 'Author',
    TypeAsString: 'User',
    MappedType: 'user',
    Required: false,
    ReadOnlyField: true,
    Hidden: false,
    Description: '',
    DefaultValue: null,
    LookupField: 'Title',
  },
  {
    Id: 'system-editor',
    Title: 'Modificado por',
    InternalName: 'Editor',
    TypeAsString: 'User',
    MappedType: 'user',
    Required: false,
    ReadOnlyField: true,
    Hidden: false,
    Description: '',
    DefaultValue: null,
    LookupField: 'Title',
  },
];

const FIELD_SELECT =
  'Id,Title,InternalName,TypeAsString,Required,ReadOnlyField,Hidden,Description,DefaultValue,Choices,LookupList,LookupField,AllowMultipleValues,MaxLength,RichText';

const SP_TYPE_MAP: Record<string, FieldMappedType> = {
  Text: 'text',
  Note: 'multiline',
  Choice: 'choice',
  MultiChoice: 'multichoice',
  Number: 'number',
  Currency: 'currency',
  Boolean: 'boolean',
  DateTime: 'datetime',
  Lookup: 'lookup',
  LookupMulti: 'lookupmulti',
  User: 'user',
  UserMulti: 'usermulti',
  URL: 'url',
  Calculated: 'calculated',
  TaxonomyFieldType: 'taxonomy',
  TaxonomyFieldTypeMulti: 'taxonomymulti',
};

const listRef = (sp: SPFI, titleOrId: string) => {
  const isGuid = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i.test(titleOrId);
  return isGuid
    ? sp.web.lists.getById(titleOrId)
    : sp.web.lists.getByTitle(titleOrId);
};

export class FieldsService {
  private spFor(webServerRelativeUrl?: string): SPFI {
    return getSPForWeb(webServerRelativeUrl);
  }

  mapSharePointFieldType(typeAsString: string): FieldMappedType {
    return SP_TYPE_MAP[typeAsString] ?? 'unknown';
  }

  private toFieldMetadata(raw: IRawSPField): IFieldMetadata {
    return {
      ...raw,
      MappedType: this.mapSharePointFieldType(raw.TypeAsString),
    };
  }

  async getFields(listTitleOrId: string, webServerRelativeUrl?: string): Promise<IFieldMetadata[]> {
    try {
      const sp = this.spFor(webServerRelativeUrl);
      const fields = await listRef(sp, listTitleOrId).fields
        .select(FIELD_SELECT)() as IRawSPField[];
      return fields.map(f => this.toFieldMetadata(f));
    } catch (e) {
      throw new Error(`FieldsService.getFields("${listTitleOrId}"): ${e}`);
    }
  }

  async getVisibleFields(listTitleOrId: string, webServerRelativeUrl?: string): Promise<IFieldMetadata[]> {
    try {
      const sp = this.spFor(webServerRelativeUrl);
      const fields = await listRef(sp, listTitleOrId).fields
        .filter('Hidden eq false and ReadOnlyField eq false')
        .select(FIELD_SELECT)() as IRawSPField[];
      return fields.map(f => this.toFieldMetadata(f));
    } catch (e) {
      throw new Error(`FieldsService.getVisibleFields("${listTitleOrId}"): ${e}`);
    }
  }

  async getFieldByInternalName(
    listTitleOrId: string,
    internalName: string,
    webServerRelativeUrl?: string
  ): Promise<IFieldMetadata> {
    try {
      const sp = this.spFor(webServerRelativeUrl);
      const field = await listRef(sp, listTitleOrId).fields
        .getByInternalNameOrTitle(internalName)
        .select(FIELD_SELECT)() as IRawSPField;
      return this.toFieldMetadata(field);
    } catch (e) {
      throw new Error(`FieldsService.getFieldByInternalName("${listTitleOrId}", "${internalName}"): ${e}`);
    }
  }

  /** Retorna opções de um campo Choice/MultiChoice */
  async getFieldOptions(
    listTitleOrId: string,
    fieldInternalName: string,
    webServerRelativeUrl?: string
  ): Promise<string[]> {
    try {
      const field = await this.getFieldByInternalName(listTitleOrId, fieldInternalName, webServerRelativeUrl);
      return field.Choices ?? [];
    } catch (e) {
      throw new Error(`FieldsService.getFieldOptions: ${e}`);
    }
  }
}
