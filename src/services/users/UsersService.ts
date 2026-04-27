import { getSP, getSPForWeb } from '../core/sp';
import { getGraph } from '../core/graph';
import { IUserDetails, IUserGroupMembership, IPeoplePickerResult } from './types';

export class UsersService {
  private webSp(webServerRelativeUrl?: string) {
    return webServerRelativeUrl?.trim() ? getSPForWeb(webServerRelativeUrl.trim()) : getSP();
  }

  private get sp() { return getSP(); }
  private get graph() { return getGraph(); }

  async getCurrentUser(): Promise<IUserDetails> {
    try {
      const user = await this.sp.web.currentUser
        .select('Id', 'Title', 'Email', 'LoginName', 'IsSiteAdmin')();
      return user as IUserDetails;
    } catch (e) {
      throw new Error(`UsersService.getCurrentUser: ${e}`);
    }
  }

  async getSiteUsers(): Promise<IUserDetails[]> {
    try {
      const users = await this.sp.web.siteUsers
        .select('Id', 'Title', 'Email', 'LoginName', 'IsSiteAdmin')();
      return users as IUserDetails[];
    } catch (e) {
      throw new Error(`UsersService.getSiteUsers: ${e}`);
    }
  }

  async getUserById(userId: number): Promise<IUserDetails> {
    try {
      const user = await this.sp.web.siteUsers
        .getById(userId)
        .select('Id', 'Title', 'Email', 'LoginName', 'IsSiteAdmin')();
      return user as IUserDetails;
    } catch (e) {
      throw new Error(`UsersService.getUserById(${userId}): ${e}`);
    }
  }

  /** Busca usuários via People Picker (autocomplete, campos Pessoa/Grupo) */
  async searchUsers(searchText: string, maxResults = 10): Promise<IPeoplePickerResult[]> {
    try {
      const results = await this.sp.profiles.clientPeoplePickerSearchUser({
        QueryString: searchText,
        MaximumEntitySuggestions: maxResults,
        AllowMultipleEntities: false,
        PrincipalType: 1, // User
      });
      return results as IPeoplePickerResult[];
    } catch (e) {
      throw new Error(`UsersService.searchUsers("${searchText}"): ${e}`);
    }
  }

  async getCurrentUserGroupIds(webServerRelativeUrl?: string): Promise<number[]> {
    try {
      const sp = this.webSp(webServerRelativeUrl);
      const raw = await sp.web.currentUser.groups.select('Id')();
      const arr = Array.isArray(raw) ? raw : [];
      const ids: number[] = [];
      for (let i = 0; i < arr.length; i++) {
        const id = typeof arr[i].Id === 'number' ? arr[i].Id : Number(arr[i].Id);
        if (isFinite(id) && id > 0) ids.push(id);
      }
      return ids;
    } catch (e) {
      throw new Error(`UsersService.getCurrentUserGroupIds: ${e}`);
    }
  }

  async ensureUserAndGetId(keyOrLoginName: string, webServerRelativeUrl?: string): Promise<number | undefined> {
    try {
      const sp = this.webSp(webServerRelativeUrl);
      const r = await sp.web.ensureUser(keyOrLoginName);
      const o = r as { Id?: number; data?: { Id?: number } };
      const id = o.Id ?? o.data?.Id;
      return typeof id === 'number' && isFinite(id) && id > 0 ? id : undefined;
    } catch {
      return undefined;
    }
  }

  async getUserGroups(loginName?: string): Promise<IUserGroupMembership[]> {
    const loadCurrentUserGroups = async (): Promise<IUserGroupMembership[]> => {
      const groups = await this.sp.web.currentUser.groups.select('Id', 'Title')();
      return groups as IUserGroupMembership[];
    };
    try {
      if (loginName) {
        try {
          const groups = await this.sp.web.siteUsers
            .getByLoginName(loginName)
            .groups.select('Id', 'Title')();
          return groups as IUserGroupMembership[];
        } catch {
          return await loadCurrentUserGroups();
        }
      }
      return await loadCurrentUserGroups();
    } catch (e) {
      throw new Error(`UsersService.getUserGroups: ${e}`);
    }
  }

  /** Retorna o usuário do Graph (dados mais completos: foto, cargo, dept.) */
  async getCurrentUserFromGraph(): Promise<Record<string, unknown>> {
    try {
      const me = await this.graph.me.select('id', 'displayName', 'mail', 'jobTitle', 'department')();
      return me as unknown as Record<string, unknown>;
    } catch (e) {
      throw new Error(`UsersService.getCurrentUserFromGraph: ${e}`);
    }
  }
}
