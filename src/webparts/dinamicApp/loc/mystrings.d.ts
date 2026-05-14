declare interface IDinamicAppWebPartStrings {
  PropertyPaneReadModeHint: string;
  PropertyPaneNoPermissionHint: string;
  PropertyPaneEditingGroupName: string;
  PropertyPaneButtonFlexViewWizard: string;
  PropertyPaneButtonPageComponents: string;
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'DinamicAppWebPartStrings' {
  const strings: IDinamicAppWebPartStrings;
  export = strings;
}
