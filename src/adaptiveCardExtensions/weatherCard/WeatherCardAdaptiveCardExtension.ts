import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { LocationPickerView } from './quickView/LocationPickerView';
import { WeatherCardPropertyPane } from './WeatherCardPropertyPane';
import { IWeatherLocation, IWeatherData } from '../../common/models/models';
import { WeatherService } from '../../common/services/WeatherService';

export interface IWeatherCardAdaptiveCardExtensionProps {
  title: string;
  listSiteUrl: string;
  temperatureUnit: string;
}

export interface IWeatherCardAdaptiveCardExtensionState {
  locations: IWeatherLocation[];
  selectedLocationName: string;
  weatherByLocation: { [name: string]: IWeatherData };
  isLoading: boolean;
}

const CARD_VIEW_REGISTRY_ID: string = 'WeatherCard_CARD_VIEW';
export const LOCATION_PICKER_VIEW_REGISTRY_ID: string = 'WeatherCard_LOCATION_PICKER_VIEW';

export default class WeatherCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IWeatherCardAdaptiveCardExtensionProps,
  IWeatherCardAdaptiveCardExtensionState
> {
  private LOG_SOURCE = '🌤️ WeatherCardAdaptiveCardExtension';
  private _deferredPropertyPane: WeatherCardPropertyPane | undefined;
  private _weatherService = new WeatherService();

  public async onInit(): Promise<void> {
    this.state = {
      locations: [],
      selectedLocationName: '',
      weatherByLocation: {},
      isLoading: true
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(LOCATION_PICKER_VIEW_REGISTRY_ID, () => new LocationPickerView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      const unit = this.properties.temperatureUnit || 'fahrenheit';
      await this._weatherService.Init(siteUrl, this.context.pageContext);
      const locations = await this._weatherService.getLocations();
      const weatherByLocation = await this._weatherService.getWeatherForAllLocations(locations, unit);
      const stored = localStorage.getItem('WeatherCard_SelectedLocation');
      const selectedLocationName = (stored && locations.some(l => l.name === stored))
        ? stored
        : (locations[0]?.name ?? '');
      this.setState({
        locations,
        selectedLocationName,
        weatherByLocation,
        isLoading: false
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'WeatherCard-property-pane'*/
      './WeatherCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.WeatherCardPropertyPane(this._addSampleData.bind(this));
    });
  }

  private async _addSampleData(): Promise<void> {
    try {
      await this._weatherService.seedSampleLocations();
      const unit = this.properties.temperatureUnit || 'fahrenheit';
      const locations = await this._weatherService.getLocations();
      const weatherByLocation = await this._weatherService.getWeatherForAllLocations(locations, unit);
      this.setState({
        locations,
        selectedLocationName: this.state.selectedLocationName || (locations[0]?.name ?? ''),
        weatherByLocation
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_addSampleData) - ${err}`);
    }
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
