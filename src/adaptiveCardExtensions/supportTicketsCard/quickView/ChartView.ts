import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/ChartViewTemplate.json';
import {ISupportTicketsCardAdaptiveCardExtensionProps,ISupportTicketsCardAdaptiveCardExtensionState,TICKETS_VIEW_REGISTRY_ID} from '../SupportTicketsCardAdaptiveCardExtension';

export interface IChartViewData {
  title: string;
  chartUrl: string;
  open: number;
  inProgress: number;
  resolved: number;
  escalated: number;
  total: number;
}

export class ChartView extends BaseAdaptiveCardQuickView<
  ISupportTicketsCardAdaptiveCardExtensionProps,
  ISupportTicketsCardAdaptiveCardExtensionState,
  IChartViewData
> {
  private LOG_SOURCE: string = '🎫 ChartView';

  public get data(): IChartViewData {
    const { counts, chartUrl } = this.state;
    return {
      title: this.properties.title,
      chartUrl: chartUrl ?? '',
      open: counts.open,
      inProgress: counts.inProgress,
      resolved: counts.resolved,
      escalated: counts.escalated,
      total: counts.open + counts.inProgress + counts.resolved + counts.escalated
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit' && action.data?.id === 'viewTickets') {
        this.quickViewNavigator.push(TICKETS_VIEW_REGISTRY_ID);
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
