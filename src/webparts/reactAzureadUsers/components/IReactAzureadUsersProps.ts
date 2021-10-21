import { MSGraphClient } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IReactAzureadUsersProps {
  description: string;
  graphClient: MSGraphClient;
  context: WebPartContext;  
}
