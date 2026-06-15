import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  facilitiesMaintenanceCardInstance
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';

import * as template from './template/MaintenanceUpdateViewTemplate.json';

interface IMaintenanceUpdateViewData {
  requestId: string;
  title: string;
  category: string;
  location: string;
  priority: string;
  requestedByName: string;
  currentStatus: string;
  currentNotes: string;
}

export class MaintenanceUpdateView extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  IMaintenanceUpdateViewData
> {
  public get data(): IMaintenanceUpdateViewData {
    const r = this.state.selectedRequest;
    return {
      requestId: r?.id ?? '',
      title: r?.title ?? '',
      category: r?.category ?? '',
      location: r?.location ?? '',
      priority: r?.priority ?? '',
      requestedByName: r?.requestedByName ?? '',
      currentStatus: r?.status ?? 'In Progress',
      currentNotes: r?.maintenanceNotes ?? ''
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    const { id, requestId, updateStatus, maintenanceNotes } = action.data;

    if (id === 'update' && requestId) {
      await facilitiesMaintenanceCardInstance.updateMaintenanceStatus(
        requestId,
        updateStatus ?? 'In Progress',
        maintenanceNotes ?? ''
      );
      this.quickViewNavigator.pop();
    } else if (id === 'cancel') {
      this.quickViewNavigator.pop();
    }
  }
}
