import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';

import * as template from './template/RequestDetailViewTemplate.json';

interface IRequestDetailViewData {
  id: string;
  title: string;
  status: string;
  category: string;
  location: string;
  priority: string;
  description: string;
  requestedDate: string;
  maintenanceNotes: string;
  maintenanceUpdatedDate: string;
  managerNotes: string;
  completedDate: string;
  hasMaintenanceNotes: boolean;
  hasManagerNotes: boolean;
  hasCompletedDate: boolean;
}

export class RequestDetailView extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  IRequestDetailViewData
> {
  public get data(): IRequestDetailViewData {
    const r = this.state.selectedRequest;
    if (!r) {
      return {
        id: '', title: 'No request selected', status: '', category: '', location: '',
        priority: '', description: '', requestedDate: '', maintenanceNotes: '',
        maintenanceUpdatedDate: '', managerNotes: '', completedDate: '',
        hasMaintenanceNotes: false, hasManagerNotes: false, hasCompletedDate: false
      };
    }
    return {
      id: r.id,
      title: r.title,
      status: r.status,
      category: r.category,
      location: r.location,
      priority: r.priority,
      description: r.description,
      requestedDate: r.requestedDate,
      maintenanceNotes: r.maintenanceNotes,
      maintenanceUpdatedDate: r.maintenanceUpdatedDate,
      managerNotes: r.managerNotes,
      completedDate: r.completedDate,
      hasMaintenanceNotes: !!r.maintenanceNotes,
      hasManagerNotes: !!r.managerNotes,
      hasCompletedDate: !!r.completedDate
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    if (action.data?.id === 'back') {
      this.quickViewNavigator.pop();
    }
  }
}
