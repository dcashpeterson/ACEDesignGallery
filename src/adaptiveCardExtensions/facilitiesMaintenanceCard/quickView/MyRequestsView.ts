import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  FACILITIES_NEW_REQUEST_STEP1_VIEW_ID,
  FACILITIES_REQUEST_DETAIL_VIEW_ID,
  FACILITIES_MAINTENANCE_QUEUE_VIEW_ID,
  FACILITIES_MANAGER_QUEUE_VIEW_ID,
  facilitiesMaintenanceCardInstance
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';
import { IFacilitiesMaintenanceRequest } from '../../../common/models/models';

import * as template from './template/MyRequestsViewTemplate.json';

interface IMyRequestsViewData {
  requests: {
    id: string;
    title: string;
    status: string;
    category: string;
    location: string;
    priority: string;
    requestedDate: string;
    requestedByName: string;
    maintenanceNotes: string;
    maintenanceUpdatedDate: string;
    assignedToName: string;
    assignedManagerName: string;
    managerNotes: string;
    completedDate: string;
  }[];
  hasRequests: boolean;
  userRole: string;
}

export class MyRequestsView extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  IMyRequestsViewData
> {
  public get data(): IMyRequestsViewData {
    const { myRequests, userRole } = this.state;
    return {
      requests: myRequests.map((r: IFacilitiesMaintenanceRequest) => ({
        id: r.id,
        title: r.title,
        status: r.status,
        category: r.category,
        location: r.location,
        priority: r.priority,
        requestedDate: r.requestedDate,
        requestedByName: r.requestedByName,
        maintenanceNotes: r.maintenanceNotes,
        maintenanceUpdatedDate: r.maintenanceUpdatedDate,
        assignedToName: r.assignedToName,
        assignedManagerName: r.assignedManagerName,
        managerNotes: r.managerNotes,
        completedDate: r.completedDate
      })),
      hasRequests: myRequests.length > 0,
      userRole
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    const { id, requestId, title, status, category, location, priority, requestedDate, requestedByName, maintenanceNotes, maintenanceUpdatedDate, assignedToName, assignedManagerName, managerNotes, completedDate } = action.data;

    if (id === 'new-request') {
      this.quickViewNavigator.push(FACILITIES_NEW_REQUEST_STEP1_VIEW_ID);
    } else if (id === 'select' && requestId) {
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
          maintenanceUpdatedDate: maintenanceUpdatedDate ?? '',
          assignedToName: assignedToName ?? '',
          assignedToId: 0,
          assignedManagerName: assignedManagerName ?? '',
          assignedManagerId: 0,
          managerNotes: managerNotes ?? '',
          completedDate: completedDate ?? ''
        } as IFacilitiesMaintenanceRequest
      });
      this.quickViewNavigator.push(FACILITIES_REQUEST_DETAIL_VIEW_ID);
    } else if (id === 'my-queue') {
      this.quickViewNavigator.push(FACILITIES_MAINTENANCE_QUEUE_VIEW_ID);
    } else if (id === 'manager-queue') {
      this.quickViewNavigator.push(FACILITIES_MANAGER_QUEUE_VIEW_ID);
    }
  }
}
