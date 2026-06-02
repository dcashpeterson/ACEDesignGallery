import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/NewTicketFormViewTemplate.json';
import {
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState,
  THANK_YOU_VIEW_ID,
  myTicketsCardInstance
} from '../MyTicketsCardAdaptiveCardExtension';

export interface INewTicketFormViewData {
  isSubmitting: boolean;
}

export class NewTicketFormView extends BaseAdaptiveCardQuickView<
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState,
  INewTicketFormViewData
> {
  private LOG_SOURCE: string = '🎫 NewTicketFormView';

  public get data(): INewTicketFormViewData {
    return {
      isSubmitting: this.state.isSubmitting
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type !== 'Submit') return;
      const { id, ticketTitle, ticketPriority, ticketDescription } =
        (action.data ?? {}) as {
          id?: string;
          ticketTitle?: string;
          ticketPriority?: string;
          ticketDescription?: string;
        };

      if (id === 'submit') {
        const title = (ticketTitle ?? '').trim();
        if (!title) return;
        await myTicketsCardInstance.submitTicket(
          title,
          ticketPriority ?? 'Medium',
          ticketDescription ?? ''
        );
        this.quickViewNavigator.push(THANK_YOU_VIEW_ID);
      } else if (id === 'cancel') {
        this.quickViewNavigator.pop();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
      myTicketsCardInstance.setState({ isSubmitting: false });
    }
  }
}
