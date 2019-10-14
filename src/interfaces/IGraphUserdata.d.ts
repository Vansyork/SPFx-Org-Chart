export interface IGraphUserdata {
  "@odata.context": string;
  value?: (ValueEntity)[] | null;
}
export interface ValueEntity {
  "@odata.type": string;
  id: string;
  businessPhones?: (string | null)[] | null;
  displayName: string;
  givenName: string;
  jobTitle: string;
  mail: string;
  mobilePhone?: null;
  officeLocation: string;
  preferredLanguage?: null;
  surname: string;
  userPrincipalName: string;
}
