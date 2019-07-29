import * as React from 'react';
import styles from './FeedbackForm.module.scss';
import { IFeedbackFormProps } from './IFeedbackFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  TextField,
  DefaultButton
} from 'office-ui-fabric-react';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IFeedbackFormState {
  isBusy: boolean;
}

export default class FeedbackForm extends React.Component<IFeedbackFormProps, IFeedbackFormState> {

  constructor(props){
    super(props);

    this.state = {
      isBusy: false
    };
  }

  public render(): React.ReactElement<IFeedbackFormProps> {

    return (
      <div className={ styles.feedbackForm }>
        <TextField label="Standard" multiline rows={3} name="text" />
        <DefaultButton disabled={this.state.isBusy} onClick={() => {this._sendMessage();}}>Send</DefaultButton>
      </div>
    );
  }

  private _me():void {
    this.props.graphClient.api('/me')
    .get((error, user: MicrosoftGraph.User, rawResponse?: any) => {
      console.log(user);
    });
  }

  private _sendMessage() {
    this.setState({isBusy:true});

    const message = {
      subject:"Did you see last night's game?",
      importance:"low",
      body:{
          contentType:"html",
          content:"They were <b>awesome</b>!"
      },
      toRecipients:[
          {
              emailAddress:{
                  address:"sksdes@zx0.onmicrosoft.com"
              }
          }
      ]
  } as MicrosoftGraph.Message;

  this.props.graphClient.api('/me/sendMail')
    .post({message : message}).then((value:any) => {
      this.setState({isBusy:false});
    },(error: any) => {
      console.log(error);
    });
  }

}
