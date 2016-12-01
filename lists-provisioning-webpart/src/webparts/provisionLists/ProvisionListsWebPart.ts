import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'provisionListsStrings';
import ProvisionLists, { IProvisionListsProps } from './components/ProvisionLists';
import { IProvisionListsWebPartProps } from './IProvisionListsWebPartProps';

export default class ProvisionListsWebPart extends BaseClientSideWebPart<IProvisionListsWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    const element: React.ReactElement<IProvisionListsProps> = React.createElement(ProvisionLists, {
      provisionSourceListEndpointUrl: this.properties.provisionSourceListEndpointUrl,
      provisionDestinationListEndpointUrl: this.properties.provisionSourceListEndpointUrl,
      context : this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.UrlsGroupName,
              groupFields: [
                PropertyPaneTextField('provisionSourceListEndpointUrl', {
                  label: strings.ProvisionSourceListEndpointFieldLabel,
                  multiline: true
                }),
                PropertyPaneTextField('provisionDestinationListEndpointUrl', {
                  label: strings.ProvisionDestinationListEndpointFieldLabel,
                  multiline: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}