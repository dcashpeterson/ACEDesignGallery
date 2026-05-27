import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/IssueView.json';
import { IServiceHealthIssue } from '../../../common/models/models';
import {
  IITSystemStatusCardAdaptiveCardExtensionProps,
  IITSystemStatusCardAdaptiveCardExtensionState
} from '../ITSystemStatusCardAdaptiveCardExtension';

export interface IIssueViewData {
  title: string;
  selectedService: string;
  issues: IServiceHealthIssue[];
}

export class IssueView extends BaseAdaptiveCardQuickView<
  IITSystemStatusCardAdaptiveCardExtensionProps,
  IITSystemStatusCardAdaptiveCardExtensionState,
  IIssueViewData
> {
  private LOG_SOURCE: string = '🔴 IssueView';

  public get data(): IIssueViewData {
    const issues = this.state.issues.filter(item => item.service === this.state.selectedService);
    return {
      title: this.properties.title,
      selectedService: this.state.selectedService,
      issues: issues
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      // if (action.type === 'Submit') {
      //   const { id, service } = action.data;
      //   if (id === 'issues') {
      //     this.setState({ selectedService: service });
      //     this.quickViewNavigator.pop(true);
      //   }
        
      // }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (render) - ${err}`);
    }
  }
}