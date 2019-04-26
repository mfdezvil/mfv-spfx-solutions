import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';

import * as strings from 'GraphConnectionTestWebPartStrings';
import GraphConnectionTest from './components/GraphConnectionTest';
import { IGraphConnectionTestProps } from './components/IGraphConnectionTestProps';
import { GraphDataOptions } from './utils/GraphDataOptions';

export interface IGraphConnectionTestWebPartProps {
  graphDataOptions: GraphDataOptions;
}

export default class GraphConnectionTestWebPart extends BaseClientSideWebPart<IGraphConnectionTestWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IGraphConnectionTestProps > = React.createElement(
      GraphConnectionTest,
      {
        selectedGraphDataOption: this.properties.graphDataOptions,
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
                PropertyPaneChoiceGroup('graphDataOptions', {
                  label: strings.GraphDataOptionsLabel,
                  options: [
                    {key: GraphDataOptions.currentUser, text: "Current user data"},
                    {key: GraphDataOptions.mails, text: "Mails data"},
                    {key: GraphDataOptions.calendar, text: "Calendar events data"},
                    {key: GraphDataOptions.teams, text: "User Teams data"}
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
