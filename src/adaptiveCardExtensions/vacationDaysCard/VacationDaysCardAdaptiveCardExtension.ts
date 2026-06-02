import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { RequestsView } from './quickView/RequestsView';
import { NewRequestView } from './quickView/NewRequestView';
import { ThankYouView } from './quickView/ThankYouView';
import { VacationDaysCardPropertyPane } from './VacationDaysCardPropertyPane';
import { IVacationRequest, IVacationSummary } from '../../common/models/models';
import { VacationDaysService } from '../../common/services/VacationDaysService';

export interface IVacationDaysCardAdaptiveCardExtensionProps {
  title: string;
  listSiteUrl: string;
  totalVacationDays: string;
}

export interface IVacationDaysCardAdaptiveCardExtensionState {
  requests: IVacationRequest[];
  summary: IVacationSummary;
  isLoading: boolean;
  isSubmitting: boolean;
}

const CARD_VIEW_REGISTRY_ID: string = 'VacationDaysCard_CARD_VIEW';
export const REQUESTS_VIEW_REGISTRY_ID: string = 'VacationDaysCard_REQUESTS_VIEW';
export const NEW_REQUEST_VIEW_REGISTRY_ID: string = 'VacationDaysCard_NEW_REQUEST_VIEW';
export const THANK_YOU_VIEW_REGISTRY_ID: string = 'VacationDaysCard_THANK_YOU_VIEW';

export let vacationDaysCardInstance: VacationDaysCardAdaptiveCardExtension;

export default class VacationDaysCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState
> {
  private LOG_SOURCE = '🏖️ VacationDaysCardAdaptiveCardExtension';
  private _deferredPropertyPane: VacationDaysCardPropertyPane | undefined;
  private _vacationService = new VacationDaysService();

  public async onInit(): Promise<void> {
    vacationDaysCardInstance = this;

    this.state = {
      requests: [],
      summary: { used: 0, available: 20, total: 20 },
      isLoading: true,
      isSubmitting: false
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(REQUESTS_VIEW_REGISTRY_ID, () => new RequestsView());
    this.quickViewNavigator.register(NEW_REQUEST_VIEW_REGISTRY_ID, () => new NewRequestView());
    this.quickViewNavigator.register(THANK_YOU_VIEW_REGISTRY_ID, () => new ThankYouView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      await this._vacationService.Init(siteUrl, this.context.pageContext);
      await this.refreshData();
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  public async refreshData(): Promise<void> {
    try {
      const totalDays = parseInt(this.properties.totalVacationDays, 10) || 20;
      const authorId = (this.context.pageContext.legacyPageContext as { userId: number }).userId;
      const requests = await this._vacationService.getMyRequests(authorId);
      const summary = this._vacationService.getVacationSummary(requests, totalDays);
      this.setState({ requests, summary, isLoading: false, isSubmitting: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (refreshData) - ${err}`);
      this.setState({ isLoading: false, isSubmitting: false });
    }
  }

  public async submitRequest(title: string, startDate: string, endDate: string, days: number, notes: string): Promise<void> {
    this.setState({ isSubmitting: true });
    await this._vacationService.createRequest(title, startDate, endDate, days, notes);
    await this.refreshData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'VacationDaysCard-property-pane'*/
      './VacationDaysCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.VacationDaysCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
