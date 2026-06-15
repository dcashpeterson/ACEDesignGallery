import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  FACILITIES_SUBMIT_CONFIRMATION_VIEW_ID,
  facilitiesMaintenanceCardInstance
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';

import * as template from './template/NewRequestStep2ViewTemplate.json';

interface INewRequestStep2ViewData {
  isSubmitting: boolean;
}

export class NewRequestStep2View extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  INewRequestStep2ViewData
> {
  public get data(): INewRequestStep2ViewData {
    return { isSubmitting: this.state.isSubmitting };
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    const { id, requestPriority, requestDescription } = action.data;

    if (id === 'submit') {
      const { title, category, location } = this.state.formStep1;
      await facilitiesMaintenanceCardInstance.submitRequest(
        title,
        category,
        location,
        requestPriority ?? 'Medium',
        requestDescription ?? ''
      );
      this.quickViewNavigator.push(FACILITIES_SUBMIT_CONFIRMATION_VIEW_ID);
    } else if (id === 'back') {
      this.quickViewNavigator.pop();
    }
  }
}
