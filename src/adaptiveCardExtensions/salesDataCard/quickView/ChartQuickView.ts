import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/ChartQuickViewTemplate.json';
import {
  ISalesDataCardProps,
  ISalesDataCardState,
  DETAIL_QUICK_VIEW_ID,
  salesDataCardInstance
} from '../SalesDataCardAdaptiveCardExtension';

export interface IChartQuickViewData {
  title: string;
  chartDataUrl: string;
  totalSales: string;
  topPerformerName: string;
  topPerformerTotal: string;
  daysBack: number;
  groupBy: string;
  isLoading: boolean;
}

export class ChartQuickView extends BaseAdaptiveCardQuickView<
  ISalesDataCardProps,
  ISalesDataCardState,
  IChartQuickViewData
> {
  private LOG_SOURCE: string = '📈 ChartQuickView';

  public get data(): IChartQuickViewData {
    const { chartDataUrl, totalSales, topPerformerName, topPerformerTotal, daysBack, groupBy, isLoading } = this.state;
    return {
      title: this.properties.title,
      chartDataUrl: chartDataUrl ?? '',
      totalSales: `$${(totalSales ?? 0).toLocaleString()}`,
      topPerformerName: topPerformerName ?? '',
      topPerformerTotal: `$${(topPerformerTotal ?? 0).toLocaleString()}`,
      daysBack: daysBack ?? 30,
      groupBy: groupBy ?? 'salesperson',
      isLoading: isLoading ?? false
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type !== 'Submit') return;
      const { id, value } = (action.data ?? {}) as { id?: string; value?: string };

      if (id === 'details') {
        this.quickViewNavigator.push(DETAIL_QUICK_VIEW_ID);
        return;
      }

      if (id === 'range' && value) {
        const daysBack = parseInt(value, 10);
        if (!isNaN(daysBack)) {
          await salesDataCardInstance.refreshData(daysBack, this.state.groupBy);
        }
        return;
      }

      if (id === 'groupBy' && value) {
        await salesDataCardInstance.refreshData(this.state.daysBack, value as 'salesperson' | 'region' | 'product');
        return;
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
