import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { ChartView } from './quickView/ChartView';
import { DetailView } from './quickView/DetailView';
import { SiteAnalyticsPropertyPane } from './SiteAnalyticsPropertyPane';
import { ISiteAnalyticsSummary, ISiteActivityStat, SiteAnalyticsSummary } from '../../common/models/models';
import { SiteAnalyticsService } from '../../common/services/SiteAnalyticsService';

export interface ISiteAnalyticsCardProps {
  title: string;
  siteUrl: string;
}

export interface ISiteAnalyticsCardState {
  summary: SiteAnalyticsSummary;
  chartUrl: string;
  isLoading: boolean;
}

const CARD_VIEW_ID: string = 'SiteAnalyticsCard_CARD_VIEW';
export const CHART_VIEW_ID: string = 'SiteAnalyticsCard_CHART_VIEW';
export const DETAIL_VIEW_ID: string = 'SiteAnalyticsCard_DETAIL_VIEW';

export default class SiteAnalyticsCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ISiteAnalyticsCardProps,
  ISiteAnalyticsCardState
> {
  private LOG_SOURCE = '📊 SiteAnalyticsCardAdaptiveCardExtension';
  private _deferredPropertyPane: SiteAnalyticsPropertyPane | undefined;
  private _analyticsService = new SiteAnalyticsService(this.context);

  public async onInit(): Promise<void> {
    this.state = { summary: new SiteAnalyticsSummary(), chartUrl: '', isLoading: true };

    this.cardNavigator.register(CARD_VIEW_ID, () => new CardView());
    this.quickViewNavigator.register(CHART_VIEW_ID, () => new ChartView());
    this.quickViewNavigator.register(DETAIL_VIEW_ID, () => new DetailView());

    try {
      const siteUrl = this.properties.siteUrl || this.context.pageContext.web.absoluteUrl;
      await this._analyticsService.Init(siteUrl, this.context.pageContext, this.context.aadTokenProviderFactory);
      
      const summary = await this._analyticsService.getSiteAnalytics(siteUrl);
      const chartUrl = this._analyticsService._generateChartUrl(summary.dailyStats);
      this.setState({ summary, chartUrl, isLoading: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'SiteAnalyticsCard-property-pane'*/
      './SiteAnalyticsPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.SiteAnalyticsPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
