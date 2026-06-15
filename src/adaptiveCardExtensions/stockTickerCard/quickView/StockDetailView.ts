import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IStockTickerCardProps,
  IStockTickerCardState
} from '../StockTickerCardAdaptiveCardExtension';
import template from './template/StockDetailViewTemplate.json';

interface IStockDetailViewData {
  symbol: string;
  priceDisplay: string;
  changeDisplay: string;
  changeColor: string;
  open: string;
  high: string;
  low: string;
  previousClose: string;
  lastUpdated: string;
}

export class StockDetailView extends BaseAdaptiveCardQuickView<
  IStockTickerCardProps,
  IStockTickerCardState,
  IStockDetailViewData
> {
  public get data(): IStockDetailViewData {
    const { quote } = this.state;

    if (!quote) {
      return {
        symbol: this.properties.stockSymbol || 'AAPL',
        priceDisplay: 'No data available',
        changeDisplay: '',
        changeColor: 'Default',
        open: '—',
        high: '—',
        low: '—',
        previousClose: '—',
        lastUpdated: '—'
      };
    }

    const sign = quote.change >= 0 ? '+' : '';
    const changeColor = quote.change >= 0 ? 'Good' : 'Attention';

    return {
      symbol: quote.symbol,
      priceDisplay: `$${quote.price.toFixed(2)}`,
      changeDisplay: `${sign}${quote.change.toFixed(2)} (${sign}${quote.changePercent.toFixed(2)}%)`,
      changeColor,
      open: `$${quote.open.toFixed(2)}`,
      high: `$${quote.high.toFixed(2)}`,
      low: `$${quote.low.toFixed(2)}`,
      previousClose: `$${quote.previousClose.toFixed(2)}`,
      lastUpdated: quote.lastUpdated.toLocaleTimeString()
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }
}
