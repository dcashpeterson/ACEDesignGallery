import { IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from 'FacilitiesMaintenanceCardAdaptiveCardExtensionStrings';

export class FacilitiesMaintenanceCardPropertyPane {
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
                }),
                PropertyPaneTextField('imageUrl', {
                  label: strings.ImageUrlFieldLabel,
                  placeholder: 'https://...'
                }),
                PropertyPaneTextField('maintenanceGroupName', {
                  label: strings.MaintenanceGroupFieldLabel
                }),
                PropertyPaneTextField('managerGroupName', {
                  label: strings.ManagerGroupFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
