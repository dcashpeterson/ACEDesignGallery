import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/MyTicketsListViewTemplate.json';
import { ISupportTicket } from '../../../common/models/models';
import {
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState,
  TICKET_DETAIL_VIEW_ID,
  myTicketsCardInstance
} from '../MyTicketsCardAdaptiveCardExtension';

export interface IMyTicketsListViewData {
  tickets: ISupportTicket[];
}

export class MyTicketsListView extends BaseAdaptiveCardQuickView<
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState,
  IMyTicketsListViewData
> {
  private LOG_SOURCE: string = '🎫 MyTicketsListView';

  private _formatDate(isoDate: string): string {
    if (!isoDate) return '';
    try {
      return new Date(isoDate).toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
    } catch {
      return isoDate;
    }
  }

  public get data(): IMyTicketsListViewData {
    return {
      tickets: this.state.myTickets.map(t => ({ ...t, submittedDate: this._formatDate(t.submittedDate) }))
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type !== 'Submit') return;
      const { id, ticketId, title, status, priority, description, submittedDate, assignedTo } =
        (action.data ?? {}) as {
          id?: string;
          ticketId?: string;
          title?: string;
          status?: string;
          priority?: string;
          description?: string;
          submittedDate?: string;
          assignedTo?: string;
        };

      if (id === 'select' && ticketId) {
        const selectedTicket: ISupportTicket = {
          id: ticketId,
          title: title ?? '',
          status: status ?? '',
          priority: priority ?? '',
          description: description ?? '',
          submittedDate: submittedDate ?? '',
          assignedTo: assignedTo ?? ''
        };
        myTicketsCardInstance.setState({ selectedTicket });
        this.quickViewNavigator.push(TICKET_DETAIL_VIEW_ID);
      } else if (id === 'back') {
        this.quickViewNavigator.pop();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
