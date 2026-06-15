import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  FACILITIES_MY_REQUESTS_VIEW_ID
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';

import * as template from './template/SubmitConfirmationViewTemplate.json';

interface ISubmitConfirmationViewData {
  requestCount: number;
}

export class SubmitConfirmationView extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  ISubmitConfirmationViewData
> {
  public get data(): ISubmitConfirmationViewData {
    return { requestCount: this.state.myRequests.length };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    const { id } = action.data;

    if (id === 'view-requests') {
      this.quickViewNavigator.push(FACILITIES_MY_REQUESTS_VIEW_ID);
    } else if (id === 'close') {
      this.quickViewNavigator.close();
    }
  }
}
