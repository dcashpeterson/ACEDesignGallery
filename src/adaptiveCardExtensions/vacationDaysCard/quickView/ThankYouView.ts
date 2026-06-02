import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/ThankYouViewTemplate.json';
import {
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState,
  vacationDaysCardInstance
} from '../VacationDaysCardAdaptiveCardExtension';

export interface IThankYouViewData {
  requestCount: number;
}

export class ThankYouView extends BaseAdaptiveCardQuickView<
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState,
  IThankYouViewData
> {
  private LOG_SOURCE: string = '🏖️ ThankYouView';

  public get data(): IThankYouViewData {
    return {
      requestCount: this.state.requests.length
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit' && (action.data as { id?: string })?.id === 'close') {
        await vacationDaysCardInstance.refreshData();
        this.quickViewNavigator.close();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
