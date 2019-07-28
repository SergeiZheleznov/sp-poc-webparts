import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'FeedbackFormWebPartStrings';
import FeedbackForm from './components/FeedbackForm';
import { IFeedbackFormProps } from './components/IFeedbackFormProps';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IFeedbackFormWebPartProps {
  description: string;
}

export default class FeedbackFormWebPart extends BaseClientSideWebPart<IFeedbackFormWebPartProps> {
  private graphClient: MSGraphClient;

  public async onInit(): Promise<void> {
    console.log('init');
    return;
  }


  public render(): void {
    this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
        this.graphClient = client;
        console.log(client);

      let res = client.api('/me/sendMail')
        .post({
          "message": {
            "subject": "Meet for lunch?",
            "body": {
              "contentType": "Text",
              "content": "The new cafeteria is open."
            },
            "toRecipients": [
              {
                "emailAddress": {
                  "address": "sksdes@zx0.onmicrosoft.com"
                }
              }
            ]
          }
        }).then((value)=>{
          console.log(value);
        },(error) => {
          console.log(error);
        });

    });

    const element: React.ReactElement<IFeedbackFormProps > = React.createElement(
      FeedbackForm,
      {
        description: this.properties.description,
        graphClient: this.graphClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
