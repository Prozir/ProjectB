import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as jQuery from "jquery";
import * as strings from 'McarFormWebPartStrings';
import McarForm from './components/McarForm';
import { IMcarFormProps } from './components/IMcarFormProps';

export interface IMcarFormWebPartProps {
  description: string;
}

export default class McarFormWebPart extends BaseClientSideWebPart <IMcarFormWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
   /* sp.setup({
    spfxContext: this.context
    });*/
    jQuery("#workbenchPageContent").prop("style", "max-width: 150%");
        jQuery(".SPCanvas-canvas").prop("style", "max-width: 150%");
        jQuery(".CanvasZone").prop("style", "max-width: 150%");             
    });
    }
  public render(): void {
    const element: React.ReactElement<IMcarFormProps> = React.createElement(
      McarForm,
      {
        description: this.properties.description,
        context:this.context,
        spHttpClient: this.context.spHttpClient,  
        siteUrl: this.context.pageContext.web.absoluteUrl 
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
