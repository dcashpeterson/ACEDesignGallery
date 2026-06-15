import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import * as strings from 'StockTickerCardAdaptiveCardExtensionStrings';

export class StockTickerCardPropertyPane {
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
                PropertyPaneTextField('stockSymbol', {
                  label: strings.StockSymbolFieldLabel,
                  placeholder: 'AAPL'
                }),
                PropertyPaneTextField('apiKey', {
                  label: strings.ApiKeyFieldLabel,
                  placeholder: 'Get a free key at alphavantage.co'
                }),
                PropertyPaneDropdown('refreshIntervalMinutes', {
                  label: strings.RefreshIntervalFieldLabel,
                  options: [
                    { key: 5,  text: '5 minutes'  },
                    { key: 15, text: '15 minutes' },
                    { key: 30, text: '30 minutes' },
                    { key: 60, text: '60 minutes' }
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
