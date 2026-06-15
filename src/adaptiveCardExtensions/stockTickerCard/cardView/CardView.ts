import {BaseComponentsCardView,ComponentsCardViewParameters,IDataPoint,IExternalLinkCardAction,IQuickViewCardAction,LineChartCardView} from '@microsoft/sp-adaptive-card-extension-base';
import {IStockTickerCardProps,IStockTickerCardState,RANGE_SELECTOR_VIEW_ID,STOCK_DETAIL_VIEW_ID} from '../StockTickerCardAdaptiveCardExtension';
import { StockTimeRange } from '../../../common/models/models';
import * as strings from 'StockTickerCardAdaptiveCardExtensionStrings';

function _xFormatter(date: Date, range: StockTimeRange): string {
  if (range === strings.Range7DShort) return strings.WeekDayAbreviations[date.getDay()];
  if (range === strings.Range30DShort) return `${strings.MonthAbrevLabels[date.getMonth()]} ${date.getDate()}`;
  return strings.MonthAbrevLabels[date.getMonth()];
}

function _rangeLabel(range: StockTimeRange): string {
  let retVal = '';
  switch (range) {
    case '30D':
      retVal = strings.Range30D;
      break;
    case '1Y':
      retVal = strings.Range1Y;
      break;
    case '7D':
    default:
      retVal = strings.Range7D;
      break;
  }
  return retVal;
}

export class CardView extends BaseComponentsCardView<
  IStockTickerCardProps,
  IStockTickerCardState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const chartPoints = this.state.chartDataByRange[this.state.selectedRange] ?? [];

    let headerText: string;
    if (this.state.isLoading) {
      headerText = 'Loading…';
    } else if (!this.state.quote) {
      headerText = 'Unable to load stock data';
    } else {
      const sign = this.state.quote.change >= 0 ? '+' : '';
      headerText = `$${this.state.quote.price.toFixed(2)}`;
    }

    const seriesData: IDataPoint<Date>[] = chartPoints.map(p => ({ x: p.date, y: p.price }));

    // Sample x-axis tick labels (up to 4)
    // Silly math to make sure that we divide the number of points by 4 and then get even intervals
    const step = seriesData.length > 4 ? Math.floor(seriesData.length / 4) : 1;
    const labelValues: Date[] = seriesData.filter((_, i) => i % step === 0).map(d => d.x).slice(0, 4);

    return LineChartCardView({
      cardBar: {
        componentName: 'cardBar',
        title: `${this.properties.title} · ${this.state.symbol.toUpperCase()}`
      },
      header: {
        componentName: 'text',
        text: headerText
      },
      body: {
        componentName: 'dataVisualization',
        dataVisualizationKind: 'line',
        series: [{
          data: seriesData,
          lastDataPointLabel: _rangeLabel(this.state.selectedRange)
        }],
        xAxis: {
          labelValues,
          formatter: (date: Date) => _xFormatter(date, this.state.selectedRange)
        }
      },
      footer: {
        componentName: 'cardButton',
        title: `Range: ${_rangeLabel(this.state.selectedRange)}`,
        style: 'positive',
        action: {
          type: 'QuickView',
          parameters: { view: RANGE_SELECTOR_VIEW_ID }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: STOCK_DETAIL_VIEW_ID }
    };
  }
}
