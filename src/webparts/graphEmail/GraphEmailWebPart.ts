import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphEmailWebPartStrings';
import GraphEmail from './components/GraphEmail';
import { IGraphEmailProps } from './components/IGraphEmailProps';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IGraphEmailWebPartProps {
  description: string;
}

export default class GraphEmailWebPart extends BaseClientSideWebPart<IGraphEmailWebPartProps> {
  private graphClient : MSGraphClient;

  public onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void)
    : void => {
      this.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          this.graphClient = client;
          resolve();
        }, err => reject(err));
    } );
  }

  public render(): void {
    const element: React.ReactElement<IGraphEmailProps > = React.createElement(
      GraphEmail,
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
