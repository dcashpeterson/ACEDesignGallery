import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { RangeSelectorView } from './quickView/RangeSelectorView';
import { StockDetailView } from './quickView/StockDetailView';
import { StockTickerCardPropertyPane } from './StockTickerCardPropertyPane';
import { IStockQuote, IStockDataPoint, StockTimeRange } from '../../common/models/models';
import { StockTickerService } from '../../common/services/StockTickerService';
import strings from 'StockTickerCardAdaptiveCardExtensionStrings';

export interface IStockTickerCardProps {
  title: string;
  stockSymbol: string;
  apiKey: string;
  refreshIntervalMinutes: number;
}

export interface IStockTickerCardState {
  symbol: string;
  quote: IStockQuote | undefined;
  chartDataByRange: Partial<Record<StockTimeRange, IStockDataPoint[]>>;
  selectedRange: StockTimeRange;
  isLoading: boolean;
}

const CARD_VIEW_ID: string = 'StockTickerCard_CARD_VIEW';
export const RANGE_SELECTOR_VIEW_ID: string = 'StockTickerCard_RANGE_SELECTOR';
export const STOCK_DETAIL_VIEW_ID: string = 'StockTickerCard_DETAIL';

export let stockTickerCardInstance: StockTickerCardAdaptiveCardExtension;

export default class StockTickerCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IStockTickerCardProps,
  IStockTickerCardState
> {
  private LOG_SOURCE = '📈 StockTickerCardAdaptiveCardExtension';
  private _deferredPropertyPane: StockTickerCardPropertyPane | undefined;
  public service = new StockTickerService();
  private _refreshTimer: ReturnType<typeof setInterval> | undefined;

  public async onInit(): Promise<void> {
    stockTickerCardInstance = this;

    this.state = {
      symbol: this.properties.stockSymbol || 'MSFT',
      quote: undefined,
      chartDataByRange: {},
      selectedRange: strings.Range7DShort as StockTimeRange,
      isLoading: true
    };

    this.cardNavigator.register(CARD_VIEW_ID, () => new CardView());
    this.quickViewNavigator.register(RANGE_SELECTOR_VIEW_ID, () => new RangeSelectorView());
    this.quickViewNavigator.register(STOCK_DETAIL_VIEW_ID, () => new StockDetailView());

    this.service.Init(this.properties.apiKey || '');
    try {
      const quote = await this.service.getQuote(this.state.symbol);
      const chartData = await this.service.getHistoricalData(this.state.symbol, strings.Range7DShort as StockTimeRange);
      this.setState({ quote, chartDataByRange: { [strings.Range7DShort]: chartData }, isLoading: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
    }
    this._startRefreshTimer();

    return Promise.resolve();
  }

  public onDispose(): void {
    this._clearRefreshTimer();
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string): Promise<void> {
    if (propertyPath === 'stockSymbol' || propertyPath === 'apiKey') {
      const symbol = this.properties.stockSymbol || 'MSFT';
      this.service.Init(this.properties.apiKey || '');
      this.setState({ symbol, chartDataByRange: {}, selectedRange: strings.Range7DShort as StockTimeRange, isLoading: true });
      try {
        const quote = await this.service.getQuote(symbol);
        const chartData = await this.service.getHistoricalData(symbol, strings.Range7DShort as StockTimeRange);
        this.setState({ quote, chartDataByRange: { [strings.Range7DShort]: chartData }, isLoading: false });
      } catch (err) {
        console.error(`${this.LOG_SOURCE} (onPropertyPaneFieldChanged) - ${err}`);
      }
    }
    if (propertyPath === 'refreshIntervalMinutes') {
      this._clearRefreshTimer();
      this._startRefreshTimer();
    }
  }

  private _startRefreshTimer(): void {
    const intervalMs = (this.properties.refreshIntervalMinutes ?? 15) * 60_000;
    this._refreshTimer = setInterval(async () => {
      try {
        const { symbol } = this.state;
        const quote = await this.service.getQuote(symbol);
        const chartData = await this.service.getHistoricalData(symbol, strings.Range7DShort as StockTimeRange);
        this.setState({ quote, chartDataByRange: { ...this.state.chartDataByRange, [strings.Range7DShort]: chartData }, isLoading: false });
      } catch (err) {
        console.error(`${this.LOG_SOURCE} (_startRefreshTimer) - ${err}`);
      }
    }, intervalMs);
  }

  private _clearRefreshTimer(): void {
    if (this._refreshTimer !== undefined) {
      clearInterval(this._refreshTimer);
      this._refreshTimer = undefined;
    }
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'StockTickerCard-property-pane'*/
      './StockTickerCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.StockTickerCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
