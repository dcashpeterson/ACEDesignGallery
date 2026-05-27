import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/QuickViewTemplate.json';
import { IServiceHealth } from '../../../common/models/models';
import {
  IITSystemStatusCardAdaptiveCardExtensionProps,
  IITSystemStatusCardAdaptiveCardExtensionState,
  ISSUE_VIEW_REGISTRY_ID,
  NEW_VIEW_REGISTRY_ID
} from '../ITSystemStatusCardAdaptiveCardExtension';

export interface IQuickViewData {
  title: string;
  services: IServiceHealth[];
}

export class QuickView extends BaseAdaptiveCardQuickView<
  IITSystemStatusCardAdaptiveCardExtensionProps,
  IITSystemStatusCardAdaptiveCardExtensionState,
  IQuickViewData
> {
  private LOG_SOURCE: string = '🔴 QuickView';

  public get data(): IQuickViewData {
    return {
      title: this.properties.title,
      services: this.state.services
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit') {
        const { id, service } = action.data;
        if (id === 'issues') {
          this.setState({ selectedService: service });
          this.quickViewNavigator.push(ISSUE_VIEW_REGISTRY_ID);
        }
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (render) - ${err}`);
    }
  }
}
