import { getSP } from '../core/sp';
import { IGroupDetails, IGroupMember } from './types';

export class GroupsService {
  private get sp() { return getSP(); }

  async getSiteGroups(): Promise<IGroupDetails[]> {
    try {
      const groups = await this.sp.web.siteGroups
        .select('Id', 'Title', 'Description', 'OwnerTitle', 'AllowMembersEditMembership',
                'OnlyAllowMembersViewMembership', 'AutoAcceptRequestToJoinLeave')();
      return groups as IGroupDetails[];
    } catch (e) {
      throw new Error(`GroupsService.getSiteGroups: ${e}`);
    }
  }

  async getGroupById(groupId: number): Promise<IGroupDetails> {
    try {
      const group = await this.sp.web.siteGroups
        .getById(groupId)
        .select('Id', 'Title', 'Description', 'OwnerTitle', 'AllowMembersEditMembership',
                'OnlyAllowMembersViewMembership', 'AutoAcceptRequestToJoinLeave')();
      return group as IGroupDetails;
    } catch (e) {
      throw new Error(`GroupsService.getGroupById(${groupId}): ${e}`);
    }
  }

  async getGroupUsers(groupId: number): Promise<IGroupMember[]> {
    try {
      const members = await this.sp.web.siteGroups
        .getById(groupId)
        .users
        .select('Id', 'Title', 'Email', 'LoginName')();
      return members as IGroupMember[];
    } catch (e) {
      throw new Error(`GroupsService.getGroupUsers(${groupId}): ${e}`);
    }
  }

  async isCurrentUserInGroup(groupName: string): Promise<boolean> {
    try {
      const currentUser = await this.sp.web.currentUser
        .select('Id')();
      const members = await this.sp.web.siteGroups
        .getByName(groupName)
        .users
        .select('Id')();
      return members.some((m: { Id: number }) => m.Id === currentUser.Id);
    } catch (e) {
      // grupo não encontrado ou sem permissão → false
      return false;
    }
  }

  async isUserInGroup(userId: number, groupId: number): Promise<boolean> {
    try {
      const members = await this.sp.web.siteGroups
        .getById(groupId)
        .users
        .select('Id')();
      return members.some((m: { Id: number }) => m.Id === userId);
    } catch (e) {
      return false;
    }
  }
}
