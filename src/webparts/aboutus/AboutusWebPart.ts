import * as React from 'react';
import * as ReactDom from 'react-dom';
import { ISPFXContext } from '@pnp/common';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { sp } from "@pnp/sp";
// import 'core-js/es6/array';
// import 'es6-map/implement';
// import 'core-js/modules/es6.array.find';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';

import * as strings from 'AboutusWebPartStrings';
import Aboutus from './components/Aboutus';
import { IAboutusProps } from './components/IAboutusProps';
import { string } from 'prop-types';


export interface IAboutusWebPartProps {
  listId: string;
  accordionTitle: string;
  allowZeroExpanded: boolean;
  allowMultipleExpanded: boolean;
  [key: string]: any;
}

export default class AboutusWebPart extends BaseClientSideWebPart<IAboutusWebPartProps> {
  spfxContext: ISPFXContext;
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.spfxContext
      });
    });
  }

  public render(): void {
    const element: React.ReactElement<IAboutusProps> = React.createElement(
      Aboutus,
      {
        listId: this.properties.listId,
        accordionTitle: this.properties.accordionTitle,
        allowZeroExpanded: this.properties.allowZeroExpanded,
        allowMultipleExpanded: this.properties.allowMultipleExpanded,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.accordionTitle = value;
        },
        onConfigure: () => {
          this.context.propertyPane.open();
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

                PropertyPaneToggle('allowZeroExpanded', {
                  label: 'Allow zero expanded',
                  checked: this.properties.allowZeroExpanded,
                  key: 'allowZeroExpanded',
                }),
                PropertyPaneToggle('allowMultipleExpanded', {
                  label: 'Allow multiple expand',
                  checked: this.properties.allowMultipleExpanded,
                  key: 'allowMultipleExpanded',
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {

    /**
     * This section is used to determine when to refresh the pane options
     */
    let updateOnThese = [
      'allowZeroExpanded', 'allowMultipleExpanded', 'listId',
    ];
    //alert('props updated');
    console.log('onPropertyPaneFieldChanged:', propertyPath, oldValue, newValue);
    if (updateOnThese.indexOf(propertyPath) > -1) {
      this.properties[propertyPath] = newValue;
      this.context.propertyPane.refresh();

    } else { //This can be removed if it works

    }
    this.render();
  }
}
