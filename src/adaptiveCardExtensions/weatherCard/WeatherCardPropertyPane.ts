import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown, PropertyPaneButton, PropertyPaneButtonType } from '@microsoft/sp-property-pane';
import * as strings from 'WeatherCardAdaptiveCardExtensionStrings';

export class WeatherCardPropertyPane {
  private _onSeedSampleData: () => void;

  constructor(onSeedSampleData: () => void) {
    this._onSeedSampleData = onSeedSampleData;
  }

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
                PropertyPaneDropdown('temperatureUnit', {
                  label: strings.TemperatureUnitFieldLabel,
                  options: [
                    { key: 'fahrenheit', text: 'Fahrenheit (°F)' },
                    { key: 'celsius',    text: 'Celsius (°C)' }
                  ]
                })
              ]
            },
            {
              groupName: strings.SampleDataGroupName,
              groupFields: [
                PropertyPaneButton('', {
                  text: strings.AddSampleDataButtonLabel,
                  buttonType: PropertyPaneButtonType.Normal,
                  onClick: this._onSeedSampleData
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
