import type { SPFI } from '@pnp/sp';
import { getSPForWeb, isSharePointListGuid, normalizeListGuid, buildWebPathCandidatesForListByGuid } from '../core/sp';
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
  return isSharePointListGuid(titleOrId)
    ? sp.web.lists.getById(normalizeListGuid(titleOrId))
    : sp.web.lists.getByTitle(titleOrId);
};

export function mergeSystemMetadataFields(meta: IFieldMetadata[]): IFieldMetadata[] {
  const seen = new Set(meta.map((m) => m.InternalName));
  const extra = SYSTEM_METADATA_FIELDS.filter((s) => !seen.has(s.InternalName));
  return meta.concat(extra);
}

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

  private async execGetFieldsAll(listTitleOrId: string, webServerRelativeUrl?: string): Promise<IFieldMetadata[]> {
    const sp = this.spFor(webServerRelativeUrl);
    const fields = await listRef(sp, listTitleOrId).fields
      .select(FIELD_SELECT)() as IRawSPField[];
    return fields.map(f => this.toFieldMetadata(f));
  }

  async getFields(listTitleOrId: string, webServerRelativeUrl?: string): Promise<IFieldMetadata[]> {
    try {
      if (!isSharePointListGuid(listTitleOrId)) {
        return await this.execGetFieldsAll(listTitleOrId, webServerRelativeUrl);
      }
      let last: unknown;
      for (const w of buildWebPathCandidatesForListByGuid(webServerRelativeUrl)) {
        try {
          return await this.execGetFieldsAll(listTitleOrId, w);
        } catch (e) {
          last = e;
        }
      }
      throw last;
    } catch (e) {
      throw new Error(`FieldsService.getFields("${listTitleOrId}"): ${e}`);
    }
  }

  private async execGetVisibleFields(listTitleOrId: string, webServerRelativeUrl?: string): Promise<IFieldMetadata[]> {
    const sp = this.spFor(webServerRelativeUrl);
    const fields = await listRef(sp, listTitleOrId).fields
      .filter('Hidden eq false and ReadOnlyField eq false')
      .select(FIELD_SELECT)() as IRawSPField[];
    return fields.map(f => this.toFieldMetadata(f));
  }

  async getVisibleFields(listTitleOrId: string, webServerRelativeUrl?: string): Promise<IFieldMetadata[]> {
    try {
      if (!isSharePointListGuid(listTitleOrId)) {
        return await this.execGetVisibleFields(listTitleOrId, webServerRelativeUrl);
      }
      let last: unknown;
      for (const w of buildWebPathCandidatesForListByGuid(webServerRelativeUrl)) {
        try {
          return await this.execGetVisibleFields(listTitleOrId, w);
        } catch (e) {
          last = e;
        }
      }
      throw last;
    } catch (e) {
      throw new Error(`FieldsService.getVisibleFields("${listTitleOrId}"): ${e}`);
    }
  }

  private async execGetFieldByInternalName(
    listTitleOrId: string,
    internalName: string,
    webServerRelativeUrl?: string
  ): Promise<IFieldMetadata> {
    const sp = this.spFor(webServerRelativeUrl);
    const field = await listRef(sp, listTitleOrId).fields
      .getByInternalNameOrTitle(internalName)
      .select(FIELD_SELECT)() as IRawSPField;
    return this.toFieldMetadata(field);
  }

  async getFieldByInternalName(
    listTitleOrId: string,
    internalName: string,
    webServerRelativeUrl?: string
  ): Promise<IFieldMetadata> {
    try {
      if (!isSharePointListGuid(listTitleOrId)) {
        return await this.execGetFieldByInternalName(listTitleOrId, internalName, webServerRelativeUrl);
      }
      let last: unknown;
      for (const w of buildWebPathCandidatesForListByGuid(webServerRelativeUrl)) {
        try {
          return await this.execGetFieldByInternalName(listTitleOrId, internalName, w);
        } catch (e) {
          last = e;
        }
      }
      throw last;
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
