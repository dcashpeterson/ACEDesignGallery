import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { MyRequestsView } from './quickView/MyRequestsView';
import { RequestDetailView } from './quickView/RequestDetailView';
import { NewRequestStep1View } from './quickView/NewRequestStep1View';
import { NewRequestStep2View } from './quickView/NewRequestStep2View';
import { SubmitConfirmationView } from './quickView/SubmitConfirmationView';
import { MaintenanceQueueView } from './quickView/MaintenanceQueueView';
import { MaintenanceUpdateView } from './quickView/MaintenanceUpdateView';
import { ManagerQueueView } from './quickView/ManagerQueueView';
import { ManagerValidateView } from './quickView/ManagerValidateView';
import { FacilitiesMaintenanceCardPropertyPane } from './FacilitiesMaintenanceCardPropertyPane';
import { IFacilitiesMaintenanceRequest } from '../../common/models/models';
import { FacilitiesMaintenanceService } from '../../common/services/FacilitiesMaintenanceService';

export interface IFacilitiesMaintenanceCardAdaptiveCardExtensionProps {
  title: string;
  listSiteUrl: string;
  imageUrl: string;
  maintenanceGroupName: string;
  managerGroupName: string;
}

export interface IFacilitiesMaintenanceCardAdaptiveCardExtensionState {
  userRole: 'manager' | 'maintenance' | 'user';
  currentUserId: number;
  myRequests: IFacilitiesMaintenanceRequest[];
  maintenanceQueue: IFacilitiesMaintenanceRequest[];
  managerQueue: IFacilitiesMaintenanceRequest[];
  selectedRequest: IFacilitiesMaintenanceRequest | undefined;
  formStep1: { title: string; category: string; location: string };
  isLoading: boolean;
  isSubmitting: boolean;
}

const CARD_VIEW_REGISTRY_ID: string = 'FacilitiesMaintenanceCard_CARD_VIEW';
export const FACILITIES_MY_REQUESTS_VIEW_ID: string = 'FacilitiesMaintenanceCard_MY_REQUESTS';
export const FACILITIES_REQUEST_DETAIL_VIEW_ID: string = 'FacilitiesMaintenanceCard_REQUEST_DETAIL';
export const FACILITIES_NEW_REQUEST_STEP1_VIEW_ID: string = 'FacilitiesMaintenanceCard_NEW_REQUEST_STEP1';
export const FACILITIES_NEW_REQUEST_STEP2_VIEW_ID: string = 'FacilitiesMaintenanceCard_NEW_REQUEST_STEP2';
export const FACILITIES_SUBMIT_CONFIRMATION_VIEW_ID: string = 'FacilitiesMaintenanceCard_SUBMIT_CONFIRMATION';
export const FACILITIES_MAINTENANCE_QUEUE_VIEW_ID: string = 'FacilitiesMaintenanceCard_MAINTENANCE_QUEUE';
export const FACILITIES_MAINTENANCE_UPDATE_VIEW_ID: string = 'FacilitiesMaintenanceCard_MAINTENANCE_UPDATE';
export const FACILITIES_MANAGER_QUEUE_VIEW_ID: string = 'FacilitiesMaintenanceCard_MANAGER_QUEUE';
export const FACILITIES_MANAGER_VALIDATE_VIEW_ID: string = 'FacilitiesMaintenanceCard_MANAGER_VALIDATE';

export let facilitiesMaintenanceCardInstance: FacilitiesMaintenanceCardAdaptiveCardExtension;

export default class FacilitiesMaintenanceCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState
> {
  private LOG_SOURCE = '🔧 FacilitiesMaintenanceCardAdaptiveCardExtension';
  private _deferredPropertyPane: FacilitiesMaintenanceCardPropertyPane | undefined;
  private _service = new FacilitiesMaintenanceService();

  public async onInit(): Promise<void> {
    facilitiesMaintenanceCardInstance = this;

    this.state = {
      userRole: 'user',
      currentUserId: 0,
      myRequests: [],
      maintenanceQueue: [],
      managerQueue: [],
      selectedRequest: undefined,
      formStep1: { title: '', category: 'HVAC', location: '' },
      isLoading: true,
      isSubmitting: false
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(FACILITIES_MY_REQUESTS_VIEW_ID, () => new MyRequestsView());
    this.quickViewNavigator.register(FACILITIES_REQUEST_DETAIL_VIEW_ID, () => new RequestDetailView());
    this.quickViewNavigator.register(FACILITIES_NEW_REQUEST_STEP1_VIEW_ID, () => new NewRequestStep1View());
    this.quickViewNavigator.register(FACILITIES_NEW_REQUEST_STEP2_VIEW_ID, () => new NewRequestStep2View());
    this.quickViewNavigator.register(FACILITIES_SUBMIT_CONFIRMATION_VIEW_ID, () => new SubmitConfirmationView());
    this.quickViewNavigator.register(FACILITIES_MAINTENANCE_QUEUE_VIEW_ID, () => new MaintenanceQueueView());
    this.quickViewNavigator.register(FACILITIES_MAINTENANCE_UPDATE_VIEW_ID, () => new MaintenanceUpdateView());
    this.quickViewNavigator.register(FACILITIES_MANAGER_QUEUE_VIEW_ID, () => new ManagerQueueView());
    this.quickViewNavigator.register(FACILITIES_MANAGER_VALIDATE_VIEW_ID, () => new ManagerValidateView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      await this._service.Init(siteUrl, this.context.pageContext);
      await this.refreshData();
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  public async refreshData(): Promise<void> {
    try {
      const userId = (this.context.pageContext.legacyPageContext as { userId: number }).userId;
      const maintenanceGroup = this.properties.maintenanceGroupName || 'Facilities Maintenance Staff';
      const managerGroup = this.properties.managerGroupName || 'Facilities Managers';

      const userRole = await this._service.getUserRole(maintenanceGroup, managerGroup);
      const myRequests = await this._service.getMyRequests(userId);

      const maintenanceQueue = (userRole === 'maintenance' || userRole === 'manager')
        ? await this._service.getMaintenanceQueue(userId)
        : [];

      const managerQueue = userRole === 'manager'
        ? await this._service.getManagerQueue(userId)
        : [];

      this.setState({ userRole, currentUserId: userId, myRequests, maintenanceQueue, managerQueue, isLoading: false, isSubmitting: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (refreshData) - ${err}`);
      this.setState({ isLoading: false, isSubmitting: false });
    }
  }

  public async submitRequest(title: string, category: string, location: string, priority: string, description: string): Promise<void> {
    this.setState({ isSubmitting: true });
    const userId = (this.context.pageContext.legacyPageContext as { userId: number }).userId;
    const currentUser = this.context.pageContext.user.displayName;
    await this._service.createRequest({ title, category, location, priority, description, requestedById: userId, requestedByName: currentUser });
    await this.refreshData();
  }

  public async updateMaintenanceStatus(id: string, status: string, notes: string): Promise<void> {
    await this._service.updateStatus(id, status, notes);
    await this.refreshData();
  }

  public async validateCompletion(id: string, notes: string, approved: boolean): Promise<void> {
    await this._service.validateCompletion(id, notes, approved);
    await this.refreshData();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'FacilitiesMaintenanceCard-property-pane'*/
      './FacilitiesMaintenanceCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.FacilitiesMaintenanceCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
