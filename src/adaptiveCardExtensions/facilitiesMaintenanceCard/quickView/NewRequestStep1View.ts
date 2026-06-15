import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  FACILITIES_NEW_REQUEST_STEP2_VIEW_ID,
  facilitiesMaintenanceCardInstance
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';

import * as template from './template/NewRequestStep1ViewTemplate.json';

export class NewRequestStep1View extends BaseAdaptiveCardQuickView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  Record<string, never>
> {
  public get data(): Record<string, never> {
    return {};
  }

  public get template(): ISPFxAdaptiveCard {
    return template;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    if (action.type !== 'Submit') return;
    const { id, requestTitle, requestCategory, requestLocation } = action.data;

    if (id === 'next') {
      const title = (requestTitle ?? '').trim();
      const location = (requestLocation ?? '').trim();
      if (!title || !location) return;
      facilitiesMaintenanceCardInstance.setState({
        formStep1: {
          title,
          category: requestCategory ?? 'HVAC',
          location
        }
      });
      this.quickViewNavigator.push(FACILITIES_NEW_REQUEST_STEP2_VIEW_ID);
    } else if (id === 'cancel') {
      this.quickViewNavigator.pop();
    }
  }
}
