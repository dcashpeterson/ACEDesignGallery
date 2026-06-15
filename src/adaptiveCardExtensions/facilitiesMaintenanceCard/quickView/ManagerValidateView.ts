import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  facilitiesMaintenanceCardInstance
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';

import * as template from './template/ManagerValidateViewTemplate.json';

interface IManagerValidateViewData {
  requestId: string;
  title: string;
  status: string;
  category: string;
  location: string;
  priority: string;
  requestedByName: string;
  maintenanceNotes: string;
  hasMaintenanceNotes: boolean;
}

export class ManagerValidateView extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  IManagerValidateViewData
> {
  public get data(): IManagerValidateViewData {
    const r = this.state.selectedRequest;
    return {
      requestId: r?.id ?? '',
      title: r?.title ?? '',
      status: r?.status ?? '',
      category: r?.category ?? '',
      location: r?.location ?? '',
      priority: r?.priority ?? '',
      requestedByName: r?.requestedByName ?? '',
      maintenanceNotes: r?.maintenanceNotes ?? '',
      hasMaintenanceNotes: !!(r?.maintenanceNotes)
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    const { id, requestId, managerNotes } = action.data;

    if ((id === 'approve' || id === 'reject') && requestId) {
      await facilitiesMaintenanceCardInstance.validateCompletion(
        requestId,
        managerNotes ?? '',
        id === 'approve'
      );
      this.quickViewNavigator.pop();
    } else if (id === 'cancel') {
      this.quickViewNavigator.pop();
    }
  }
}
