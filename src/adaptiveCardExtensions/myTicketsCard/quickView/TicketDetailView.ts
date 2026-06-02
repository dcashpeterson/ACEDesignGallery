import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/TicketDetailViewTemplate.json';
import {
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState
} from '../MyTicketsCardAdaptiveCardExtension';

export interface ITicketDetailViewData {
  title: string;
  status: string;
  priority: string;
  description: string;
  submittedDate: string;
  assignedTo: string;
}

export class TicketDetailView extends BaseAdaptiveCardQuickView<
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState,
  ITicketDetailViewData
> {
  private LOG_SOURCE: string = '🎫 TicketDetailView';

  public get data(): ITicketDetailViewData {
    const ticket = this.state.selectedTicket;
    return {
      title: ticket?.title ?? '',
      status: ticket?.status ?? '',
      priority: ticket?.priority ?? '',
      description: ticket?.description ?? '',
      submittedDate: ticket?.submittedDate ?? '',
      assignedTo: ticket?.assignedTo ?? ''
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
