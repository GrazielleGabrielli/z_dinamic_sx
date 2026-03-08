export interface IGroupDetails {
  Id: number;
  Title: string;
  Description: string;
  OwnerTitle: string;
  AllowMembersEditMembership: boolean;
  OnlyAllowMembersViewMembership: boolean;
  AutoAcceptRequestToJoinLeave: boolean;
}

export interface IGroupMember {
  Id: number;
  Title: string;
  Email: string;
  LoginName: string;
}
