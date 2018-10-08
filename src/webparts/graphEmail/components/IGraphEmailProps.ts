import { MSGraphClient } from "@microsoft/sp-http";

export interface IGraphEmailProps {
  description: string;
  graphClient: MSGraphClient;
}
