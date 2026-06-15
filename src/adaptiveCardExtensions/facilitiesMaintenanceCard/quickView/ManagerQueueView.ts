import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  FACILITIES_MY_REQUESTS_VIEW_ID,
  FACILITIES_MANAGER_VALIDATE_VIEW_ID,
  facilitiesMaintenanceCardInstance
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';
import { IFacilitiesMaintenanceRequest } from '../../../common/models/models';

import * as template from './template/ManagerQueueViewTemplate.json';

interface IManagerQueueViewData {
  items: {
    id: string;
    title: string;
    status: string;
    category: string;
    location: string;
    priority: string;
    requestedDate: string;
    requestedByName: string;
    maintenanceNotes: string;
    isOverdue: boolean;
  }[];
  hasItems: boolean;
}

const SEVEN_DAYS_MS = 7 * 24 * 60 * 60 * 1000;

export class ManagerQueueView extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  IManagerQueueViewData
> {
  public get data(): IManagerQueueViewData {
    const { managerQueue } = this.state;
    const now = Date.now();
    return {
      items: managerQueue.map((r: IFacilitiesMaintenanceRequest) => {
        const requestedMs = r.requestedDate ? new Date(r.requestedDate).getTime() : 0;
        const isOverdue = r.status === 'New' && requestedMs > 0 && (now - requestedMs) > SEVEN_DAYS_MS;
        return {
          id: r.id,
          title: r.title,
          status: r.status,
          category: r.category,
          location: r.location,
          priority: r.priority,
          requestedDate: r.requestedDate,
          requestedByName: r.requestedByName,
          maintenanceNotes: r.maintenanceNotes,
          isOverdue
        };
      }),
      hasItems: managerQueue.length > 0
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    const { id, requestId, title, status, category, location, priority, requestedDate, requestedByName, maintenanceNotes } = action.data;

    if (id === 'validate' && requestId) {
      facilitiesMaintenanceCardInstance.setState({
        selectedRequest: {
          id: requestId,
          title: title ?? '',
          status: status ?? '',
          category: category ?? '',
          location: location ?? '',
          priority: priority ?? '',
          description: '',
          requestedDate: requestedDate ?? '',
          requestedByName: requestedByName ?? '',
          requestedById: 0,
          maintenanceNotes: maintenanceNotes ?? '',
          maintenanceUpdatedDate: '',
          assignedToName: '',
          assignedToId: 0,
          assignedManagerName: '',
          assignedManagerId: 0,
          managerNotes: '',
          completedDate: ''
        } as IFacilitiesMaintenanceRequest
      });
      this.quickViewNavigator.push(FACILITIES_MANAGER_VALIDATE_VIEW_ID);
    } else if (id === 'view-mine') {
      this.quickViewNavigator.push(FACILITIES_MY_REQUESTS_VIEW_ID);
    }
  }
}
