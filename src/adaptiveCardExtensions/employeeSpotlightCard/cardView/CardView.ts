import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  ImageCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'EmployeeSpotlightCardAdaptiveCardExtensionStrings';
import {
  IEmployeeSpotlightCardProps,
  IEmployeeSpotlightCardState,
  QUICK_VIEW_REGISTRY_ID
} from '../EmployeeSpotlightCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IEmployeeSpotlightCardProps,
  IEmployeeSpotlightCardState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { spotlights, currentIndex } = this.state;
    const spotlight = spotlights[currentIndex];
    const hasSpotlight = !!spotlight;

    return ImageCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      image: {
        url: hasSpotlight ? spotlight.employeeImageUrl : '',
        altText: hasSpotlight ? spotlight.employeeName : ''
      },
      header: {
        componentName: 'text',
        text: hasSpotlight ? spotlight.employeeName : strings.NoSpotlights
      },
      footer: {
        componentName: 'cardButton',
        title: hasSpotlight ? spotlight.spotlightType : strings.ReadStory,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: {
        view: QUICK_VIEW_REGISTRY_ID
      }
    };
  }
}
