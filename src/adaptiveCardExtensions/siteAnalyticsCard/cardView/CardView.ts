import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  IDataPoint,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  LineChartCardView
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'SiteAnalyticsCardAdaptiveCardExtensionStrings';
import {
  ISiteAnalyticsCardProps,
  ISiteAnalyticsCardState,
  CHART_VIEW_ID
} from '../SiteAnalyticsCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ISiteAnalyticsCardProps,
  ISiteAnalyticsCardState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { summary, isLoading } = this.state;

    let headerText: string;
    if (isLoading) {
      headerText = 'Loading analytics…';
    } else if (!summary) {
      headerText = 'Unable to load analytics';
    } else {
      headerText = `${summary.last7Views.toLocaleString()} views this week`;
    }
    const seriesData : IDataPoint<Date>[] = [
    ];
    this.state.summary.dailyStats.map(item =>{
      const [year, month, day] = item.date.split('-').map(Number);
      seriesData.push({ x: new Date(year,month-1,day), y: item.viewCount })
    })
    return LineChartCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
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
          lastDataPointLabel: ''
      }],
      xAxis: {
        labelValues: seriesData.map(d => d.x),
        formatter: (date: Date) => strings.WeekDayAbreviations[date.getDay()]
      }
      },
      footer: summary
        ? {
            componentName: 'cardButton',
            title: strings.QuickViewButton,
            style: 'positive',
            action: {
              type: 'QuickView',
              parameters: { view: CHART_VIEW_ID }
            }
          }
        : undefined
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: CHART_VIEW_ID }
    };
  }
}
