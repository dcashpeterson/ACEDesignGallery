import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/SafetyTipDetailViewTemplate.json';
import { ISafetyTip } from '../../../common/models/models';
import {
  IDailySafetyTipCardProps,
  IDailySafetyTipCardState
} from '../DailySafetyTipCardAdaptiveCardExtension';

export interface ISafetyTipDetailViewData {
  tip: ISafetyTip;
  hasPolicyLink: boolean;
}

const EMPTY_TIP: ISafetyTip = {
  id: '',
  title: '',
  summary: '',
  guidance: '',
  category: '',
  policyUrl: '',
  policyTitle: ''
};

export class SafetyTipDetailView extends BaseAdaptiveCardQuickView<
  IDailySafetyTipCardProps,
  IDailySafetyTipCardState,
  ISafetyTipDetailViewData
> {
  public get data(): ISafetyTipDetailViewData {
    const tip = this.state.todaysTip ?? EMPTY_TIP;
    return {
      tip,
      hasPolicyLink: !!tip.policyUrl
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }
}
