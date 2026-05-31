import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/DetailQuickViewTemplate.json';
import { ISalesDataSeries } from '../../../common/models/models';
import {
  ISalesDataCardProps,
  ISalesDataCardState
} from '../SalesDataCardAdaptiveCardExtension';

const GROUP_BY_LABELS: Record<string, string> = {
  salesperson: 'Salesperson',
  region: 'Region',
  product: 'Product'
};

export interface IDetailQuickViewData {
  title: string;
  daysBack: number;
  groupByLabel: string;
  series: Array<{ name: string; totalFormatted: string }>;
}

export class DetailQuickView extends BaseAdaptiveCardQuickView<
  ISalesDataCardProps,
  ISalesDataCardState,
  IDetailQuickViewData
> {
  private LOG_SOURCE: string = '📈 DetailQuickView';

  public get data(): IDetailQuickViewData {
    const { series, daysBack, groupBy } = this.state;
    return {
      title: this.properties.title,
      daysBack: daysBack ?? 30,
      groupByLabel: GROUP_BY_LABELS[groupBy ?? 'salesperson'] ?? 'Salesperson',
      series: (series ?? []).map((s: ISalesDataSeries) => ({
        name: s.name,
        totalFormatted: `$${s.total.toLocaleString()}`
      }))
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit' && (action.data as { id?: string })?.id === 'back') {
        this.quickViewNavigator.pop();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
