export interface IUserDetails {
  Id: number;
  Title: string;
  Email: string;
  LoginName: string;
  IsSiteAdmin: boolean;
}

export interface IUserGroupMembership {
  Id: number;
  Title: string;
}

export interface IPeoplePickerResult {
  Key: string;
  DisplayText: string;
  Description: string;
  EntityType: string;
  ProviderDisplayName: string;
  IsResolved: boolean;
}
