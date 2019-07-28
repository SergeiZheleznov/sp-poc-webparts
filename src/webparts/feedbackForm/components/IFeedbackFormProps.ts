import { MSGraphClient } from "@microsoft/sp-http";

export interface IFeedbackFormProps {
  description: string;
  graphClient: MSGraphClient;
}
