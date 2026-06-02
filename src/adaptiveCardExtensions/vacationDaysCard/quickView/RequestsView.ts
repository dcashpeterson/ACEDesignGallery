import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/RequestsViewTemplate.json';
import { IVacationRequest } from '../../../common/models/models';
import {
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState,
  NEW_REQUEST_VIEW_REGISTRY_ID
} from '../VacationDaysCardAdaptiveCardExtension';

export interface IRequestsViewData {
  title: string;
  usedDays: number;
  availableDays: number;
  totalDays: number;
  requests: IVacationRequest[];
}

export class RequestsView extends BaseAdaptiveCardQuickView<
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState,
  IRequestsViewData
> {
  private LOG_SOURCE: string = '🏖️ RequestsView';

  public get data(): IRequestsViewData {
    const { summary, requests } = this.state;
    return {
      title: this.properties.title,
      usedDays: summary.used,
      availableDays: summary.available,
      totalDays: summary.total,
      requests
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit' && action.data?.id === 'newRequest') {
        this.quickViewNavigator.push(NEW_REQUEST_VIEW_REGISTRY_ID);
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
