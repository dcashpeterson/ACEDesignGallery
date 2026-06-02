import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { MyTicketsListView } from './quickView/MyTicketsListView';
import { TicketDetailView } from './quickView/TicketDetailView';
import { NewTicketFormView } from './quickView/NewTicketFormView';
import { ThankYouView } from './quickView/ThankYouView';
import { MyTicketsCardPropertyPane } from './MyTicketsCardPropertyPane';
import { ISupportTicket } from '../../common/models/models';
import { SupportTicketsService } from '../../common/services/SupportTicketsService';

export interface IMyTicketsCardAdaptiveCardExtensionProps {
  title: string;
  listSiteUrl: string;
}

export interface IMyTicketsCardAdaptiveCardExtensionState {
  myTickets: ISupportTicket[];
  myTicketCount: number;
  selectedTicket: ISupportTicket | undefined;
  isLoading: boolean;
  isSubmitting: boolean;
}

const CARD_VIEW_REGISTRY_ID: string = 'MyTicketsCard_CARD_VIEW';
export const MY_TICKETS_LIST_VIEW_ID: string = 'MyTicketsCard_MY_TICKETS_LIST_VIEW';
export const TICKET_DETAIL_VIEW_ID: string = 'MyTicketsCard_TICKET_DETAIL_VIEW';
export const NEW_TICKET_FORM_VIEW_ID: string = 'MyTicketsCard_NEW_TICKET_FORM_VIEW';
export const THANK_YOU_VIEW_ID: string = 'MyTicketsCard_THANK_YOU_VIEW';

export let myTicketsCardInstance: MyTicketsCardAdaptiveCardExtension;

export default class MyTicketsCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState
> {
  private LOG_SOURCE = '🎫 MyTicketsCardAdaptiveCardExtension';
  private _deferredPropertyPane: MyTicketsCardPropertyPane | undefined;
  private _ticketsService = new SupportTicketsService();

  public async onInit(): Promise<void> {
    myTicketsCardInstance = this;

    this.state = {
      myTickets: [],
      myTicketCount: 0,
      selectedTicket: undefined,
      isLoading: true,
      isSubmitting: false
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(MY_TICKETS_LIST_VIEW_ID, () => new MyTicketsListView());
    this.quickViewNavigator.register(TICKET_DETAIL_VIEW_ID, () => new TicketDetailView());
    this.quickViewNavigator.register(NEW_TICKET_FORM_VIEW_ID, () => new NewTicketFormView());
    this.quickViewNavigator.register(THANK_YOU_VIEW_ID, () => new ThankYouView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      await this._ticketsService.Init(siteUrl, this.context.pageContext);
      await this.refreshData();
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  public async refreshData(): Promise<void> {
    try {
      const authorId = (this.context.pageContext.legacyPageContext as { userId: number }).userId;
      const myTickets = await this._ticketsService.getMyTickets(authorId);
      this.setState({ myTickets, myTicketCount: myTickets.length, isLoading: false, isSubmitting: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (refreshData) - ${err}`);
      this.setState({ isLoading: false, isSubmitting: false });
    }
  }

  public async submitTicket(title: string, priority: string, description: string): Promise<void> {
    this.setState({ isSubmitting: true });
    await this._ticketsService.createTicket(title, priority, description);
    await this.refreshData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'MyTicketsCard-property-pane'*/
      './MyTicketsCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.MyTicketsCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
