import * as React from 'react';
import styles from './FeedbackForm.module.scss';
import { IFeedbackFormProps } from './IFeedbackFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  TextField,
  DefaultButton
} from 'office-ui-fabric-react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export default class FeedbackForm extends React.Component<IFeedbackFormProps, {}> {

  public render(): React.ReactElement<IFeedbackFormProps> {
    console.log(this.props.graphClient);

    return (
      <div className={ styles.feedbackForm }>
        <TextField label="Standard" multiline rows={3} name="text" />
        <DefaultButton onClick={this._sendMessage}>Send</DefaultButton>
      </div>
    );
  }

  private async _sendMessage() {

    const message = {
      subject:"Did you see last night's game?",
      importance:"Low",
      body:{
          contentType:"HTML",
          content:"They were <b>awesome</b>!"
      },
      toRecipients:[
          {
              emailAddress:{
                  address:"sksdes@zx0.onmicrosoft.com"
              }
          }
      ]
  };

  let res = await this.props.graphClient.api('/me/messages')
    .post({message : message});
  }

}
