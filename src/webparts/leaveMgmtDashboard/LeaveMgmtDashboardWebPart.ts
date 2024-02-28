import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';

import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  PropertyPaneToggle,
} from '@microsoft/sp-property-pane';

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'LeaveMgmtDashboardWebPartStrings';
import LeaveMgmtDashboard from './components/LeaveMgmtDashboard';
import { ILeaveMgmtDashboardProps } from './components/ILeaveMgmtDashboardProps';

export interface ILeaveMgmtDashboardWebPartProps {
  description: string;
 
}

export default class LeaveMgmtDashboardWebPart extends BaseClientSideWebPart<ILeaveMgmtDashboardWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ILeaveMgmtDashboardProps> = React.createElement(
      LeaveMgmtDashboard,
      {
        description: this.properties.description,
        context: this.context,
        siteurl: this.context.pageContext.web.absoluteUrl,
        userId: this.context.pageContext.legacyPageContext["userId"]
         
      
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
