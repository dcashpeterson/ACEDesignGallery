import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/ThankYouViewTemplate.json';
import {
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState
} from '../MyTicketsCardAdaptiveCardExtension';

export interface IThankYouViewData {
  myTicketCount: number;
}

export class ThankYouView extends BaseAdaptiveCardQuickView<
  IMyTicketsCardAdaptiveCardExtensionProps,
  IMyTicketsCardAdaptiveCardExtensionState,
  IThankYouViewData
> {
  private LOG_SOURCE: string = '🎫 ThankYouView';

  public get data(): IThankYouViewData {
    return {
      myTicketCount: this.state.myTicketCount
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit' && (action.data as { id?: string })?.id === 'close') {
        this.quickViewNavigator.close();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
