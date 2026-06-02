import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { SafetyTipDetailView } from './quickView/SafetyTipDetailView';
import { DailySafetyTipCardPropertyPane } from './DailySafetyTipCardPropertyPane';
import { ISafetyTip } from '../../common/models/models';
import { SafetyTipsService } from '../../common/services/SafetyTipsService';

export interface IDailySafetyTipCardProps {
  title: string;
  listSiteUrl: string;
}

export interface IDailySafetyTipCardState {
  tips: ISafetyTip[];
  todaysTip: ISafetyTip | undefined;
  isLoading: boolean;
}

const CARD_VIEW_REGISTRY_ID: string = 'DailySafetyTip_CARD_VIEW';
export const DETAIL_VIEW_REGISTRY_ID: string = 'DailySafetyTip_DETAIL_VIEW';

export let dailySafetyTipCardInstance: DailySafetyTipCardAdaptiveCardExtension;

export default class DailySafetyTipCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IDailySafetyTipCardProps,
  IDailySafetyTipCardState
> {
  private LOG_SOURCE = '🦺 DailySafetyTipCardAdaptiveCardExtension';
  private _deferredPropertyPane: DailySafetyTipCardPropertyPane | undefined;
  private _safetyTipsService = new SafetyTipsService();

  public async onInit(): Promise<void> {
    dailySafetyTipCardInstance = this;
    this.state = { tips: [], todaysTip: undefined, isLoading: true };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(DETAIL_VIEW_REGISTRY_ID, () => new SafetyTipDetailView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      await this._safetyTipsService.Init(siteUrl, this.context.pageContext);
      const tips = await this._safetyTipsService.getSafetyTips();
      const todaysTip = this._pickDailyTip(tips);
      this.setState({ tips, todaysTip, isLoading: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  private _pickDailyTip(tips: ISafetyTip[]): ISafetyTip | undefined {
    if (tips.length === 0) return undefined;
    const today = new Date();
    const dayOfYear = Math.floor(
      (today.getTime() - new Date(today.getFullYear(), 0, 0).getTime()) / 86400000
    );
    return tips[dayOfYear % tips.length];
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'DailySafetyTipCard-property-pane'*/
      './DailySafetyTipCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.DailySafetyTipCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
