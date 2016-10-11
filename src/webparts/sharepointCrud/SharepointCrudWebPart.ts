import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import * as strings from 'sharepointCrudStrings';
import SharepointCrud, { ISharepointCrudProps } from './components/SharepointCrud';
import { ISharepointCrudWebPartProps } from './ISharepointCrudWebPartProps';

export default class SharepointCrudWebPart extends BaseClientSideWebPart<ISharepointCrudWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    const element: React.ReactElement<ISharepointCrudProps> = React.createElement(SharepointCrud, {
      description: this.properties.description,
      httpClient : this.context.httpClient,
      siteUrl : this.context.pageContext.web.absoluteUrl,
      listName : this.properties.listName
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
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label : strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
