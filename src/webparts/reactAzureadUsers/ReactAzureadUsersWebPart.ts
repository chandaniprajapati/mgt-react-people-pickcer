import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactAzureadUsersWebPartStrings';
import ReactAzureadUsers from './components/ReactAzureadUsers';
import { IReactAzureadUsersProps } from './components/IReactAzureadUsersProps';
import { MSGraphClient } from '@microsoft/sp-http';
import { Providers, SharePointProvider } from '@microsoft/mgt-spfx';

export interface IReactAzureadUsersWebPartProps {
  description: string;
}

export default class ReactAzureadUsersWebPart extends BaseClientSideWebPart<IReactAzureadUsersWebPartProps> {

  private graphClient: MSGraphClient;

  public onInit(): Promise<void> {
    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          this.graphClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  public render(): void {
    const element: React.ReactElement<IReactAzureadUsersProps> = React.createElement(
      ReactAzureadUsers,
      {
        description: this.properties.description,
        graphClient: this.graphClient,
        context: this.context
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
