import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { EmployeeSpotlightCardPropertyPane } from './EmployeeSpotlightCardPropertyPane';
import { IEmployeeSpotlight } from '../../common/models/models';
import { EmployeeSpotlightService } from '../../common/services/EmployeeSpotlightService';

export interface IEmployeeSpotlightCardProps {
  title: string;
  siteUrl: string;
}

export interface IEmployeeSpotlightCardState {
  spotlights: IEmployeeSpotlight[];
  currentIndex: number;
}

const CARD_VIEW_REGISTRY_ID: string = 'EmployeeSpotlightCard_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'EmployeeSpotlightCard_QUICK_VIEW';

export default class EmployeeSpotlightCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IEmployeeSpotlightCardProps,
  IEmployeeSpotlightCardState
> {
  private LOG_SOURCE = '🌟 EmployeeSpotlightCardAdaptiveCardExtension';
  private _deferredPropertyPane: EmployeeSpotlightCardPropertyPane | undefined;
  private _spotlightService = new EmployeeSpotlightService();

  public async onInit(): Promise<void> {
    this.state = { spotlights: [], currentIndex: 0 };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    try {
      const siteUrl = this.properties.siteUrl || this.context.pageContext.web.absoluteUrl;
      await this._spotlightService.Init(siteUrl, this.context.pageContext);
      const spotlights = await this._spotlightService.getSpotlights();
      this.setState({ spotlights });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
    }

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'EmployeeSpotlightCard-property-pane'*/
      './EmployeeSpotlightCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.EmployeeSpotlightCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
