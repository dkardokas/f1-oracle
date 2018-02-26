import * as React from 'react';
import * as ReactDom from 'react-dom';
import { 
  Version, 
  Environment,
  EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'F1EntryWebPartStrings';
import F1Entry from './components/F1Entry';
import { IF1EntryProps } from './components/IF1EntryProps';
import {
  SPHttpClient,
  SPHttpClientResponse   
 } from '@microsoft/sp-http';

export interface IF1EntryWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
 }
 
 export interface ISPList {
  Title: string;
  Id: string;
 }


export default class F1EntryWebPart extends BaseClientSideWebPart<IF1EntryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IF1EntryProps > = React.createElement(
      F1Entry,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
    this._renderListAsync();
  }

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }

   private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      return;
    }
    else if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
        <span class="ms-font-l">${item.Title}</span>`;
    });
 
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
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
