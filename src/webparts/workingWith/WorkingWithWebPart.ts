import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration
} from '@microsoft/sp-webpart-base';

import * as strings from 'WorkingWithWebPartStrings';
import { WorkingWith, IWorkingWithProps } from './components';
import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/propertyFields/number';
import { MSGraphClient } from '@microsoft/sp-http';

export interface IWorkingWithWebPartProps {
  title: string;
  nrOfContacts: number;
}

export default class WorkingWithWebPart extends BaseClientSideWebPart<IWorkingWithWebPartProps> {
  private graphClient: MSGraphClient;

  public onInit(): Promise<void> {
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
    const element: React.ReactElement<IWorkingWithProps > = React.createElement(
      WorkingWith,
      {
        title: this.properties.title,
        nrOfContacts: this.properties.nrOfContacts,
        graphClient: this.graphClient,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
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
                PropertyFieldNumber("nrOfContacts", {
                  key: "nrOfContacts",
                  label: strings.NrOfContactsToShow,
                  value: this.properties.nrOfContacts,
                  minValue: 1,
                  maxValue: 10
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
