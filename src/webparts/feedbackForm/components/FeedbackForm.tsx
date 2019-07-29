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
  message: string;
}

export default class FeedbackForm extends React.Component<IFeedbackFormProps, IFeedbackFormState> {

  constructor(props){
    super(props);

    this.state = {
      isBusy: false,
      message: ''
    };
  }

  public render(): React.ReactElement<IFeedbackFormProps> {

    return (
      <div className={ styles.feedbackForm }>
        <TextField label="Feedback" multiline rows={3} name="text" value={this.state.message} onChange={this._onChange} />
        <DefaultButton disabled={this.state.isBusy} onClick={this._sendMessage}>Send</DefaultButton>
      </div>
    );
  }

  private _onChange = (event: React.ChangeEvent<HTMLInputElement>) : void => {
    this.setState({message:event.target.value});
  }

  private _sendMessage = (event: React.MouseEvent<HTMLButtonElement, MouseEvent>) : void => {
    this.setState({isBusy:true});

    const msg = {
      subject:"Did you see last night's game?",
      importance:"low",
      body:{
          contentType:"text",
          content: this.state.message
      },
      toRecipients:[
          {
              emailAddress:{
                  address: this.props.targetEmail
              }
          }
      ]
  } as MicrosoftGraph.Message;

  this.props.graphClient.api('/me/sendMail')
    .post({
      message : msg
    }).then((value:any) => {
      this.setState({
        isBusy:false,
        message: ''
      });
    },(error: any) => {
      console.log(error);
    });
  }

}
