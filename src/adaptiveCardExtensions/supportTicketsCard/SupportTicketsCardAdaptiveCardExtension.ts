import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { ChartView } from './quickView/ChartView';
import { TicketsView } from './quickView/TicketsView';
import { SupportTicketsCardPropertyPane } from './SupportTicketsCardPropertyPane';
import { ISupportTicket, ITicketStatusCounts } from '../../common/models/models';
import { SupportTicketsService } from '../../common/services/SupportTicketsService';

export interface ISupportTicketsCardAdaptiveCardExtensionProps {
  title: string;
  listSiteUrl: string;
}

export interface ISupportTicketsCardAdaptiveCardExtensionState {
  tickets: ISupportTicket[];
  counts: ITicketStatusCounts;
  chartUrl: string;
  isLoading: boolean;
}

const CARD_VIEW_REGISTRY_ID: string = 'SupportTicketsCard_CARD_VIEW';
export const CHART_VIEW_REGISTRY_ID: string = 'SupportTicketsCard_CHART_VIEW';
export const TICKETS_VIEW_REGISTRY_ID: string = 'SupportTicketsCard_TICKETS_VIEW';

export default class SupportTicketsCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ISupportTicketsCardAdaptiveCardExtensionProps,
  ISupportTicketsCardAdaptiveCardExtensionState
> {
  private LOG_SOURCE = '🎫 SupportTicketsCardAdaptiveCardExtension';
  private _deferredPropertyPane: SupportTicketsCardPropertyPane | undefined;
  private _ticketsService = new SupportTicketsService();

  public async onInit(): Promise<void> {
    this.state = {
      tickets: [],
      counts: { open: 0, inProgress: 0, resolved: 0, escalated: 0 },
      chartUrl: '',
      isLoading: true
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(CHART_VIEW_REGISTRY_ID, () => new ChartView());
    this.quickViewNavigator.register(TICKETS_VIEW_REGISTRY_ID, () => new TicketsView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      await this._ticketsService.Init(siteUrl, this.context.pageContext);
      const tickets = await this._ticketsService.getTickets();
      const counts = this._ticketsService.getStatusCounts(tickets);
      const chartUrl = this._ticketsService.generateDonutChartUrl(counts);
      this.setState({ tickets, counts, chartUrl, isLoading: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'SupportTicketsCard-property-pane'*/
      './SupportTicketsCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.SupportTicketsCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
