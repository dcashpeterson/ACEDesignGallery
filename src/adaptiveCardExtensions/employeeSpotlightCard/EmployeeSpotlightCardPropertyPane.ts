import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'EmployeeSpotlightCardAdaptiveCardExtensionStrings';

export class EmployeeSpotlightCardPropertyPane {
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
                  placeholder: 'https://tenant.sharepoint.com/sites/yoursite'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
