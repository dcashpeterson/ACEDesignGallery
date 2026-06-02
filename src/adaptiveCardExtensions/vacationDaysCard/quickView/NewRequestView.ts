import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/NewRequestViewTemplate.json';
import {
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState,
  THANK_YOU_VIEW_REGISTRY_ID,
  vacationDaysCardInstance
} from '../VacationDaysCardAdaptiveCardExtension';

export interface INewRequestViewData {
  isSubmitting: boolean;
}

export class NewRequestView extends BaseAdaptiveCardQuickView<
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState,
  INewRequestViewData
> {
  private LOG_SOURCE: string = '🏖️ NewRequestView';

  public get data(): INewRequestViewData {
    return {
      isSubmitting: this.state.isSubmitting
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type !== 'Submit') return;
      const { id, requestTitle, startDate, endDate, requestDays, requestNotes } =
        (action.data ?? {}) as {
          id?: string;
          requestTitle?: string;
          startDate?: string;
          endDate?: string;
          requestDays?: string;
          requestNotes?: string;
        };

      if (id === 'submitRequest') {
        const title = (requestTitle ?? '').trim();
        if (!title) return;
        const days = parseInt(requestDays ?? '0', 10) || 0;
        await vacationDaysCardInstance.submitRequest(
          title,
          startDate ?? '',
          endDate ?? '',
          days,
          requestNotes ?? ''
        );
        this.quickViewNavigator.push(THANK_YOU_VIEW_REGISTRY_ID);
      } else if (id === 'cancel') {
        this.quickViewNavigator.pop();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
      vacationDaysCardInstance.setState({ isSubmitting: false });
    }
  }
}
