import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { ChartQuickView } from './quickView/ChartQuickView';
import { DetailQuickView } from './quickView/DetailQuickView';
import { SalesDataCardPropertyPane } from './SalesDataCardPropertyPane';
import { ISalesDataSeries } from '../../common/models/models';
import { SalesDataService } from '../../common/services/SalesDataService';
import { ChartRenderer } from './quickView/components/ChartRenderer';

export interface ISalesDataCardProps {
  title: string;
  siteUrl: string;
}

export interface ISalesDataCardState {
  series: ISalesDataSeries[];
  chartDataUrl: string;
  totalSales: number;
  topPerformerName: string;
  topPerformerTotal: number;
  daysBack: number;
  groupBy: 'salesperson' | 'region' | 'product';
  isLoading: boolean;
}

const CARD_VIEW_ID: string = 'SalesDataCard_CARD_VIEW';
export const CHART_QUICK_VIEW_ID: string = 'SalesDataCard_CHART_QUICK_VIEW';
export const DETAIL_QUICK_VIEW_ID: string = 'SalesDataCard_DETAIL_QUICK_VIEW';

// Module-level singleton so QuickViews can trigger data refreshes
// Safe per-bundle because config.json gives salesDataCard its own webpack bundle
export let salesDataCardInstance: SalesDataCardAdaptiveCardExtension;

export default class SalesDataCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ISalesDataCardProps,
  ISalesDataCardState
> {
  private LOG_SOURCE = '📈 SalesDataCardAdaptiveCardExtension';
  private _deferredPropertyPane: SalesDataCardPropertyPane | undefined;
  private _salesService = new SalesDataService();

  public async onInit(): Promise<void> {
    salesDataCardInstance = this;

    this.state = {
      series: [],
      chartDataUrl: '',
      totalSales: 0,
      topPerformerName: '',
      topPerformerTotal: 0,
      daysBack: 30,
      groupBy: 'region',
      isLoading: true
    };

    this.cardNavigator.register(CARD_VIEW_ID, () => new CardView());
    this.quickViewNavigator.register(CHART_QUICK_VIEW_ID, () => new ChartQuickView());
    this.quickViewNavigator.register(DETAIL_QUICK_VIEW_ID, () => new DetailQuickView());

    try {
      const siteUrl = this.properties.siteUrl || this.context.pageContext.web.absoluteUrl;
      await this._salesService.Init(siteUrl, this.context.pageContext);
      await this._refreshData(90, 'region');
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  public async refreshData(daysBack: number, groupBy: 'salesperson' | 'region' | 'product'): Promise<void> {
    this.setState({ isLoading: true, daysBack, groupBy });
    await this._refreshData(daysBack, groupBy);
  }

  private async _refreshData(daysBack: number, groupBy: 'salesperson' | 'region' | 'product'): Promise<void> {
    try {
      const items = await this._salesService.getSalesItems(daysBack);
      const series = this._salesService.buildSeries(items, daysBack, 8, groupBy);

      const totalSales = series.reduce((sum, s) => sum + s.total, 0);
      const topPerformer = series[0] ?? { name: '', total: 0 };

      const chartDataUrl = await ChartRenderer.render(series.slice(0, 5), daysBack);

      this.setState({
        series,
        chartDataUrl,
        totalSales,
        topPerformerName: topPerformer.name,
        topPerformerTotal: topPerformer.total,
        daysBack,
        groupBy,
        isLoading: false
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_refreshData) - ${err}`);
      this.setState({ isLoading: false });
    }
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'SalesDataCard-property-pane'*/
      './SalesDataCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.SalesDataCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
