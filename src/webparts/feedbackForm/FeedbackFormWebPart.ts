import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'FeedbackFormWebPartStrings';
import FeedbackForm from './components/FeedbackForm';
import { IFeedbackFormProps } from './components/IFeedbackFormProps';
import { MSGraphClient } from '@microsoft/sp-http';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IFeedbackFormWebPartProps {
  targetEmail: string;
  messageSubject: string;
}

export default class FeedbackFormWebPart extends BaseClientSideWebPart<IFeedbackFormWebPartProps> {

  private _graphClient: MSGraphClient;

  public onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void ): void => {
      this.context.msGraphClientFactory
        .getClient()
        .then((cli: MSGraphClient): void => {
          this._graphClient = cli;
          resolve();
        }, err => reject(err));
    });
  }


  public render(): void {

    if (this.displayMode == DisplayMode.Edit) {
      // TODO: Add edit mode view
    }

    const element: React.ReactElement<IFeedbackFormProps > = React.createElement(
      FeedbackForm,
      {
        graphClient: this._graphClient,
        targetEmail: this.properties.targetEmail,
        messageSubject: this.properties.messageSubject
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
                PropertyPaneTextField('targetEmail', {
                  label: strings.TargetEmailFieldLabel
                }),
                PropertyPaneTextField('messageSubject', {
                  label: strings.MessageSubjectFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
