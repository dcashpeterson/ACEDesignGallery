declare interface IStockTickerCardAdaptiveCardExtensionStrings {
  PropertyPaneDescription: string;
  TitleFieldLabel: string;
  StockSymbolFieldLabel: string;
  ApiKeyFieldLabel: string;
  RefreshIntervalFieldLabel: string;
  SelectRangeTitle: string;
  DetailViewTitle: string;
  Range7D: string;
  Range30D: string;
  Range1Y: string;
  Range7DShort: string;
  Range30DShort: string;
  Range1YShort: string;
  MonthAbrevLabels: string[];
  WeekDayAbreviations: string[];
}

declare module 'StockTickerCardAdaptiveCardExtensionStrings' {
  const strings: IStockTickerCardAdaptiveCardExtensionStrings;
  export = strings;
}
