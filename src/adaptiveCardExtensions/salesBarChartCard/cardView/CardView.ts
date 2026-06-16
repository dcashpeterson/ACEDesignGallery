import {
  BarChartCardView,
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  IExternalLinkCardAction,
  IBarChartSeries,
  IQuickViewCardAction,
  MaxThreeTuple
} from '@microsoft/sp-adaptive-card-extension-base';
import {
  ISalesBarChartCardProps,
  ISalesBarChartCardState,
  QUICK_VIEW_ID
} from '../SalesBarChartCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ISalesBarChartCardProps,
  ISalesBarChartCardState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { series, totalSales, isLoading } = this.state;

    const headerText = isLoading
      ? 'Loading sales data…'
      : series.length === 0
      ? 'No sales data'
      : `Total: $${totalSales.toLocaleString()}`;

    const barSeries = series
      .slice(0, 3)
      .map(s => ({
        data: s.dataPoints.map(dp => ({
          x: dp.date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' }),
          y: dp.amount
        })),
        name: s.name,
        color: s.color
      })) as MaxThreeTuple<IBarChartSeries<string>>;

    return BarChartCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      body: [
        {
          componentName: 'text',
          text: headerText
        },
        {
          componentName: 'dataVisualization',
          dataVisualizationKind: 'bar',
          series: barSeries
        }
      ]
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
      return {
        type: 'QuickView',
        parameters: { view: QUICK_VIEW_ID }
      };
    }
}
