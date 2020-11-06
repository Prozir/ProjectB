import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from '@microsoft/sp-http';

export interface IMcarFormProps {
  description: string;
  context:WebPartContext;
 // isexistingsolution:boolean;
  spHttpClient:SPHttpClient;
  siteUrl:string;
}
