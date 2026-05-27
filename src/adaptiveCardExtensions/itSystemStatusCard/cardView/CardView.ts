import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'ITSystemStatusCardAdaptiveCardExtensionStrings';
import {
  IITSystemStatusCardAdaptiveCardExtensionProps,
  IITSystemStatusCardAdaptiveCardExtensionState,
  QUICK_VIEW_REGISTRY_ID
} from '../ITSystemStatusCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IITSystemStatusCardAdaptiveCardExtensionProps,
  IITSystemStatusCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const impacted = this.state.services.filter((s: { statusType: string }) => s.statusType !== '🟢 Service Operational');
    const headerText = impacted.length === 0
      ? 'All systems operational'
      : `${impacted.length} service${impacted.length !== 1 ? 's' : ''} impacted`;
    
    const footer = impacted.length === 0
      ? undefined
      : {
          componentName: 'cardButton' as const,
          title: strings.QuickViewButton,
          action: {
            type: 'QuickView' as const,
            parameters: {
              view: QUICK_VIEW_REGISTRY_ID
            }
          }
        };

    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: headerText
      },
      footer: footer
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
