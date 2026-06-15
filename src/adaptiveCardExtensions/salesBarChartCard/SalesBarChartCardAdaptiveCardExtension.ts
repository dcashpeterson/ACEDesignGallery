import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { SalesBarChartCardPropertyPane } from './SalesBarChartCardPropertyPane';
import { ISalesDataSeries } from '../../common/models/models';
import { SalesDataService } from '../../common/services/SalesDataService';

export interface ISalesBarChartCardProps {
  title: string;
  siteUrl: string;
}

export interface ISalesBarChartCardState {
  series: ISalesDataSeries[];
  totalSales: number;
  isLoading: boolean;
}

const CARD_VIEW_ID: string = 'SalesBarChartCard_CARD_VIEW';

export default class SalesBarChartCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ISalesBarChartCardProps,
  ISalesBarChartCardState
> {
  private LOG_SOURCE = '📊 SalesBarChartCardAdaptiveCardExtension';
  private _deferredPropertyPane: SalesBarChartCardPropertyPane | undefined;
  private _salesService = new SalesDataService();

  public async onInit(): Promise<void> {
    this.state = {
      series: [],
      totalSales: 0,
      isLoading: true
    };

    this.cardNavigator.register(CARD_VIEW_ID, () => new CardView());

    try {
      const siteUrl = this.properties.siteUrl || this.context.pageContext.web.absoluteUrl;
      await this._salesService.Init(siteUrl, this.context.pageContext);
      const items = await this._salesService.getSalesItems(90);
      const series = this._salesService.buildMonthlySeriesByRegion(items);
      const totalSales = series.reduce((sum, s) => sum + s.total, 0);
      this.setState({ series, totalSales, isLoading: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'SalesBarChartCard-property-pane'*/
      './SalesBarChartCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.SalesBarChartCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
