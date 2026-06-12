import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/LocationPickerTemplate.json';
import {
  IWeatherCardAdaptiveCardExtensionProps,
  IWeatherCardAdaptiveCardExtensionState
} from '../WeatherCardAdaptiveCardExtension';
import { IWeatherLocation } from '../../../common/models/models';

export interface ILocationPickerViewData {
  locations: IWeatherLocation[];
}

export class LocationPickerView extends BaseAdaptiveCardQuickView<
  IWeatherCardAdaptiveCardExtensionProps,
  IWeatherCardAdaptiveCardExtensionState,
  ILocationPickerViewData
> {
  private LOG_SOURCE: string = '🌤️ LocationPickerView';

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
          localStorage.setItem('WeatherCard_SelectedLocation', location);
          this.setState({ selectedLocationName: location });
          this.quickViewNavigator.close();
        }
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
