import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from '@microsoft/sp-webpart-base';

import * as strings from 'TodoWebPartStrings';
import Todo from './components/Todo';
import { ITodoProps } from './components/ITodoProps';

export interface ITodoWebPartProps {
  description: string;
}

export default class TodoWebPart extends BaseClientSideWebPart<ITodoWebPartProps> {
  public static context: WebPartContext;

  protected onInit(): Promise<void> {
    // Set global context
    TodoWebPart.context = this.context;
    return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<ITodoProps > = React.createElement(
      Todo,
      {
        description: this.properties.description
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
