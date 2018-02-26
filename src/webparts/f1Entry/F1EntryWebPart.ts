import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  Version,
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'F1EntryWebPartStrings';
import F1Entry from './components/F1Entry';
import { IF1EntryProps } from './components/IF1EntryProps';


export interface IF1EntryWebPartProps {
  description: string;
}



export default class F1EntryWebPart extends BaseClientSideWebPart<IF1EntryWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IF1EntryProps> = React.createElement(
      F1Entry,
      {
        description: this.properties.description,
        context: this.context
      }
    );
    ReactDom.render(element, this.domElement);
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
