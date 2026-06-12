import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  ImageCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'WeatherCardAdaptiveCardExtensionStrings';
import {
  IWeatherCardAdaptiveCardExtensionProps,
  IWeatherCardAdaptiveCardExtensionState,
  LOCATION_PICKER_VIEW_REGISTRY_ID
} from '../WeatherCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IWeatherCardAdaptiveCardExtensionProps,
  IWeatherCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { selectedLocationName, weatherByLocation, isLoading } = this.state;
    const weather = weatherByLocation[selectedLocationName];
    const tempUnit = this.properties.temperatureUnit === 'celsius' ? 'C' : 'F';

    let headerText: string;
    if (isLoading) {
      headerText = 'Loading weather...';
    } else if (!selectedLocationName) {
      headerText = strings.SelectLocation;
    } else if (weather) {
      headerText = `${selectedLocationName} · ${Math.round(weather.temperature)}°${tempUnit} — ${weather.description}`;
    } else {
      headerText = selectedLocationName;
    }

    return ImageCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: headerText
      },
      image: {
        url: weather?.iconUrl ?? '',
        altText: weather?.description ?? 'Weather'
      },
      footer: {
        componentName: 'cardButton' as const,
        title: selectedLocationName ? strings.ChangeLocation : strings.SelectLocation,
        style: 'positive',
        action: {
          type: 'QuickView' as const,
          parameters: {
            view: LOCATION_PICKER_VIEW_REGISTRY_ID
          }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: {
        view: LOCATION_PICKER_VIEW_REGISTRY_ID
      }
    };
  }
}
