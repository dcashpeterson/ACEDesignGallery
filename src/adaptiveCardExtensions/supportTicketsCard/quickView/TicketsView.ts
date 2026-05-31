import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/TicketsViewTemplate.json';
import { ISupportTicket } from '../../../common/models/models';
import {
  ISupportTicketsCardAdaptiveCardExtensionProps,
  ISupportTicketsCardAdaptiveCardExtensionState
} from '../SupportTicketsCardAdaptiveCardExtension';

export interface ITicketsViewData {
  tickets: ISupportTicket[];
}

export class TicketsView extends BaseAdaptiveCardQuickView<
  ISupportTicketsCardAdaptiveCardExtensionProps,
  ISupportTicketsCardAdaptiveCardExtensionState,
  ITicketsViewData
> {
  private LOG_SOURCE: string = '🎫 TicketsView';

  public get data(): ITicketsViewData {
    return {
      tickets: this.state.tickets
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
