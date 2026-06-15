import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IStockTickerCardProps,
  IStockTickerCardState,
  stockTickerCardInstance
} from '../StockTickerCardAdaptiveCardExtension';
import { StockTimeRange } from '../../../common/models/models';
import * as strings from 'StockTickerCardAdaptiveCardExtensionStrings';
import template from './template/RangeSelectorViewTemplate.json';

interface IRangeSelectorViewData {
  title: string;
}

export class RangeSelectorView extends BaseAdaptiveCardQuickView<
  IStockTickerCardProps,
  IStockTickerCardState,
  IRangeSelectorViewData
> {
  public get data(): IRangeSelectorViewData {
    return { title: strings.SelectRangeTitle };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as unknown as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    if (action.data?.id !== 'setRange') return;

    const range = action.data.range as StockTimeRange;
    if (!range) return;

    this.setState({ selectedRange: range });
    try {
      const quote = await stockTickerCardInstance.service.getQuote(this.state.symbol);
      const chartData = await stockTickerCardInstance.service.getHistoricalData(this.state.symbol, range);
      
      this.setState({
        quote,
        chartDataByRange: { ...this.state.chartDataByRange, [range]: chartData },
        isLoading: false
      });
    } catch (err) {
      console.error(`RangeSelectorView (onAction) - ${err}`);
    }

    this.quickViewNavigator.close();
  }
}
