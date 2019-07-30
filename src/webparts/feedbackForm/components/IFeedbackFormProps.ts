import { MSGraphClient } from '@microsoft/sp-http';

export interface IFeedbackFormProps {
  graphClient: MSGraphClient;
  targetEmail: string;
  messageSubject: string;
}
