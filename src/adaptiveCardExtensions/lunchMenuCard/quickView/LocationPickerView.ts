import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/LocationPickerTemplate.json';
import {
  ILunchMenuCardAdaptiveCardExtensionProps,
  ILunchMenuCardAdaptiveCardExtensionState,
  MENU_VIEW_REGISTRY_ID
} from '../LunchMenuCardAdaptiveCardExtension';
import { ILocation } from '../../../common/models/models';


export interface ILocationPickerViewData {
  locations: ILocation[];
}

export class LocationPickerView extends BaseAdaptiveCardQuickView<
  ILunchMenuCardAdaptiveCardExtensionProps,
  ILunchMenuCardAdaptiveCardExtensionState,
  ILocationPickerViewData
> {
  private LOG_SOURCE: string = '🍽️ LocationPickerView';

  public get data(): ILocationPickerViewData {
    return {
      locations: this.state.locations
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit') {
        const { id, location } = action.data;
        if (id === 'selectLocation' && location) {
          this.setState({ selectedLocation: location });
          this.quickViewNavigator.push(MENU_VIEW_REGISTRY_ID);
        }
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
