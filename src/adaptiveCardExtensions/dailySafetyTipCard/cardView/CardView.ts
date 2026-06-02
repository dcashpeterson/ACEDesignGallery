import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  PrimaryTextCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'DailySafetyTipCardAdaptiveCardExtensionStrings';
import {
  IDailySafetyTipCardProps,
  IDailySafetyTipCardState,
  DETAIL_VIEW_REGISTRY_ID
} from '../DailySafetyTipCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IDailySafetyTipCardProps,
  IDailySafetyTipCardState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { todaysTip, isLoading } = this.state;

    if (isLoading || !todaysTip) {
      return PrimaryTextCardView({
        cardBar: { componentName: 'cardBar', title: this.properties.title },
        header: {
          componentName: 'text',
          text: isLoading ? strings.Loading : strings.NoTip
        },
        body: { componentName: 'text', text: '' },
        footer: undefined
      });
    }

    return PrimaryTextCardView({
      cardBar: { componentName: 'cardBar', title: this.properties.title },
      header: { componentName: 'text', text: todaysTip.title },
      body: { componentName: 'text', text: todaysTip.summary },
      footer: {
        componentName: 'cardButton',
        title: strings.LearnMore,
        action: {
          type: 'QuickView',
          parameters: { view: DETAIL_VIEW_REGISTRY_ID }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    if (!this.state.todaysTip) return undefined;
    return {
      type: 'QuickView',
      parameters: { view: DETAIL_VIEW_REGISTRY_ID }
    };
  }
}
