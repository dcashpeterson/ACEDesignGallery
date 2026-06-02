import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'CompanyAnnouncementsCardAdaptiveCardExtensionStrings';

export class CompanyAnnouncementsCardPropertyPane {
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
                PropertyPaneTextField('listSiteUrl', {
                  label: strings.ListSiteUrlFieldLabel,
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
