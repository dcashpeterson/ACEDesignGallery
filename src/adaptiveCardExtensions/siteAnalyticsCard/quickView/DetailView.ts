import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/DetailViewTemplate.json';
import {
  ISiteAnalyticsCardProps,
  ISiteAnalyticsCardState
} from '../SiteAnalyticsCardAdaptiveCardExtension';

export interface IDetailViewData {
  title: string;
  siteName: string;
  last7Views: string;
  last7Visitors: string;
  avgDailyViews: string;
  allTimeViews: string;
  allTimeVisitors: string;
}

export class DetailView extends BaseAdaptiveCardQuickView<
  ISiteAnalyticsCardProps,
  ISiteAnalyticsCardState,
  IDetailViewData
> {
  private LOG_SOURCE: string = '📊 DetailView';

  public get data(): IDetailViewData {
    const { summary } = this.state;
    const siteUrl = this.properties.siteUrl || '';
    const siteName = siteUrl
      ? siteUrl.replace(/https?:\/\/[^/]+/, '').replace(/^\/sites\//, '') || siteUrl
      : 'Current site';

    const last7Views    = summary?.last7Views    ?? 0;
    const last7Visitors = summary?.last7Visitors ?? 0;
    const avgDaily      = last7Views > 0 ? Math.round(last7Views / 7) : 0;

    return {
      title: this.properties.title,
      siteName,
      last7Views:      last7Views.toLocaleString(),
      last7Visitors:   last7Visitors.toLocaleString(),
      avgDailyViews:   avgDaily.toLocaleString(),
      allTimeViews:    (summary?.allTimeViews    ?? 0).toLocaleString(),
      allTimeVisitors: (summary?.allTimeVisitors ?? 0).toLocaleString()
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit' && action.data?.id === 'back') {
        this.quickViewNavigator.pop();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
