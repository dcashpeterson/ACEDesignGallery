import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/ChartViewTemplate.json';
import {
  ISiteAnalyticsCardProps,
  ISiteAnalyticsCardState,
  DETAIL_VIEW_ID
} from '../SiteAnalyticsCardAdaptiveCardExtension';

export interface IChartViewData {
  title: string;
  chartUrl: string;
  totalViews: string;
  uniqueVisitors: string;
}

export class ChartView extends BaseAdaptiveCardQuickView<
  ISiteAnalyticsCardProps,
  ISiteAnalyticsCardState,
  IChartViewData
> {
  private LOG_SOURCE: string = '📊 ChartView';

  public get data(): IChartViewData {
    const { summary, chartUrl } = this.state;
    return {
      title: this.properties.title,
      chartUrl: chartUrl ?? '',
      totalViews: (summary?.last7Views ?? 0).toLocaleString(),
      uniqueVisitors: (summary?.last7Visitors ?? 0).toLocaleString()
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit' && action.data?.id === 'details') {
        this.quickViewNavigator.push(DETAIL_VIEW_ID);
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
