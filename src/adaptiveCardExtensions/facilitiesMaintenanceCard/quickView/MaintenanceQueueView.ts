import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  FACILITIES_MY_REQUESTS_VIEW_ID,
  FACILITIES_MAINTENANCE_UPDATE_VIEW_ID,
  facilitiesMaintenanceCardInstance
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';
import { IFacilitiesMaintenanceRequest } from '../../../common/models/models';

import * as template from './template/MaintenanceQueueViewTemplate.json';

interface IMaintenanceQueueViewData {
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
  }[];
  hasItems: boolean;
}

export class MaintenanceQueueView extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  IMaintenanceQueueViewData
> {
  public get data(): IMaintenanceQueueViewData {
    const { maintenanceQueue } = this.state;
    return {
      items: maintenanceQueue.map((r: IFacilitiesMaintenanceRequest) => ({
        id: r.id,
        title: r.title,
        status: r.status,
        category: r.category,
        location: r.location,
        priority: r.priority,
        requestedDate: r.requestedDate,
        requestedByName: r.requestedByName,
        maintenanceNotes: r.maintenanceNotes
      })),
      hasItems: maintenanceQueue.length > 0
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    const { id, requestId, title, status, category, location, priority, requestedDate, requestedByName, maintenanceNotes } = action.data;

    if (id === 'update' && requestId) {
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
      this.quickViewNavigator.push(FACILITIES_MAINTENANCE_UPDATE_VIEW_ID);
    } else if (id === 'view-mine') {
      this.quickViewNavigator.push(FACILITIES_MY_REQUESTS_VIEW_ID);
    }
  }
}
