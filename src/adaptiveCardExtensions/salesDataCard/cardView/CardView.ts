import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  IDataPoint,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  LineChartCardView
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'SalesDataCardAdaptiveCardExtensionStrings';
import {
  ISalesDataCardProps,
  ISalesDataCardState,
  CHART_QUICK_VIEW_ID
} from '../SalesDataCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ISalesDataCardProps,
  ISalesDataCardState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { series, totalSales, isLoading } = this.state;

    const headerText = isLoading
      ? 'Loading sales data…'
      : series.length === 0
      ? 'No sales data'
      : `Total: $${totalSales.toLocaleString()}`;

      const seriesData : IDataPoint<Date>[] = [
  { x: new Date(2024, 1, 1), y: 1000 },
  { x: new Date(2024, 2, 1), y: 2400 },
  { x: new Date(2024, 3, 1), y: 2000 },
  { x: new Date(2024, 4, 1), y: 2900 },
  { x: new Date(2024, 5, 1), y: 3000 },
  { x: new Date(2024, 6, 1), y: 3100 }
];

const seriesData2 : IDataPoint<Date>[] = [
  { x: new Date(2024, 1, 1), y: 600 },
  { x: new Date(2024, 2, 1), y: 1200 },
  { x: new Date(2024, 3, 1), y: 3200 },
  { x: new Date(2024, 4, 1), y: 2800 },
  { x: new Date(2024, 5, 1), y: 3600 },
  { x: new Date(2024, 6, 1), y: 4500 }
];

const seriesData3 : IDataPoint<Date>[] = [
  { x: new Date(2024, 1, 1), y: 5200 },
  { x: new Date(2024, 2, 1), y: 1000 },
  { x: new Date(2024, 3, 1), y: 1800 },
  { x: new Date(2024, 4, 1), y: 2900 },
  { x: new Date(2024, 5, 1), y: 600 },
  { x: new Date(2024, 6, 1), y: 400 }
];

    // Map top-3 ISalesDataSeries to ILineChartSeries<Date>
    const top3 = series.slice(0, 3).map(s => ({
      data: s.dataPoints.map((dp): IDataPoint<Date> => ({ x: dp.date, y: dp.amount })),
      color: s.color,
      lastDataPointLabel: s.name
    }));

    // Build x-axis tick labels from first series (up to 4 ticks for Medium card)
    const firstData = top3[0]?.data ?? [];
    const step = firstData.length > 4 ? Math.floor(firstData.length / 4) : 1;
    const labelValues: Date[] = firstData.filter((_, i) => i % step === 0).map(d => d.x).slice(0, 4);

    // Construct a MaxThreeTuple — TypeScript infers literal array expressions as tuples
    const empty = { data: [] as IDataPoint<Date>[], lastDataPointLabel: '' };
    const seriesTuple =
      top3.length >= 3 ? [top3[0], top3[1], top3[2]] as [typeof top3[0], typeof top3[0], typeof top3[0]] :
      top3.length === 2 ? [top3[0], top3[1]] as [typeof top3[0], typeof top3[0]] :
      top3.length === 1 ? [top3[0]] as [typeof top3[0]] :
      [empty] as [typeof empty];

    return LineChartCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      body: {
        componentName: 'dataVisualization',
        dataVisualizationKind: 'line',
        series: seriesTuple
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: CHART_QUICK_VIEW_ID }
    };
  }
}
