import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'SalesDataCardAdaptiveCardExtensionStrings';

export class SalesDataCardPropertyPane {
  public getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlFieldLabel,
                  description: strings.SiteUrlFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
