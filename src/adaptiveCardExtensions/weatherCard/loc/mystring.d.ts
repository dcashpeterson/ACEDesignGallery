declare interface IWeatherCardAdaptiveCardExtensionStrings {
  PropertyPaneDescription: string;
  TitleFieldLabel: string;
  ListSiteUrlFieldLabel: string;
  TemperatureUnitFieldLabel: string;
  SelectLocation: string;
  ChangeLocation: string;
  WeatherCardDescription: string;
  SampleDataGroupName: string;
  AddSampleDataButtonLabel: string;
}

declare module 'WeatherCardAdaptiveCardExtensionStrings' {
  const strings: IWeatherCardAdaptiveCardExtensionStrings;
  export = strings;
}
