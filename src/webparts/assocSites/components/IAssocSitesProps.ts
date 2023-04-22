import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, } from '@microsoft/sp-http';

export interface IAssocSitesProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context: WebPartContext;
  spHttpClient: SPHttpClient;

}
export interface IAssocSitesState {
  sites: any;
  response: any;

}
