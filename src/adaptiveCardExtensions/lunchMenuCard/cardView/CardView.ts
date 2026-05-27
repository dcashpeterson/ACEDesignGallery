import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'LunchMenuCardAdaptiveCardExtensionStrings';
import {
  ILunchMenuCardAdaptiveCardExtensionProps,
  ILunchMenuCardAdaptiveCardExtensionState,
  LOCATION_PICKER_VIEW_REGISTRY_ID,
  MENU_VIEW_REGISTRY_ID
} from '../LunchMenuCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ILunchMenuCardAdaptiveCardExtensionProps,
  ILunchMenuCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { selectedLocation } = this.state;
    const hasLocation = !!selectedLocation;

    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: hasLocation ? `${selectedLocation} Menu` : strings.SelectLocation
      },
      footer: {
        componentName: 'cardButton' as const,
        title: hasLocation ? strings.ViewMenu : strings.ChooseLocation,
        style: 'positive',
        action: {
          type: 'QuickView' as const,
          parameters: {
            view: hasLocation ? MENU_VIEW_REGISTRY_ID : LOCATION_PICKER_VIEW_REGISTRY_ID
          }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    const { selectedLocation } = this.state;
    return {
      type: 'QuickView',
      parameters: {
        view: selectedLocation ? MENU_VIEW_REGISTRY_ID : LOCATION_PICKER_VIEW_REGISTRY_ID
      }
    };
  }
}
