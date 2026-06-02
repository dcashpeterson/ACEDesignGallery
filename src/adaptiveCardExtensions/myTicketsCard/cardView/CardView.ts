import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'MyTicketsCardAdaptiveCardExtensionStrings';
import {
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState,
  MY_TICKETS_LIST_VIEW_ID,
  NEW_TICKET_FORM_VIEW_ID
} from '../MyTicketsCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { myTicketCount, isLoading } = this.state;
    const count = myTicketCount;
    const headerText = isLoading
      ? 'Loading…'
      : `${count} ticket${count !== 1 ? 's' : ''} submitted by you`;

    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: headerText
      },
      footer: [
        {
          componentName: 'cardButton',
          title: strings.QuickViewButton,
          style: 'positive',
          action: {
            type: 'QuickView',
            parameters: { view: MY_TICKETS_LIST_VIEW_ID }
          }
        },
        {
          componentName: 'cardButton',
          title: 'Submit New Ticket',
          action: {
            type: 'QuickView',
            parameters: { view: NEW_TICKET_FORM_VIEW_ID }
          }
        }
      ]
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: MY_TICKETS_LIST_VIEW_ID }
    };
  }
}
