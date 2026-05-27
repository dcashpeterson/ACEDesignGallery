import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { LocationPickerView } from './quickView/LocationPickerView';
import { MenuView } from './quickView/MenuView';
import { LunchMenuCardPropertyPane } from './LunchMenuCardPropertyPane';
import { ILunchMenuItem, ILocation } from '../../common/models/models';
import { LunchMenuService } from '../../common/services/LunchMenuService';

export interface ILunchMenuCardAdaptiveCardExtensionProps {
  title: string;
  listSiteUrl: string;
}

export interface ILunchMenuCardAdaptiveCardExtensionState {
  locations: ILocation[];
  menuItems: ILunchMenuItem[];
  selectedLocation: string;
  listReady: boolean;
}

const CARD_VIEW_REGISTRY_ID: string = 'LunchMenuCard_CARD_VIEW';
export const LOCATION_PICKER_VIEW_REGISTRY_ID: string = 'LunchMenuCard_LOCATION_PICKER_VIEW';
export const MENU_VIEW_REGISTRY_ID: string = 'LunchMenuCard_MENU_VIEW';

export default class LunchMenuCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ILunchMenuCardAdaptiveCardExtensionProps,
  ILunchMenuCardAdaptiveCardExtensionState
> {
  private LOG_SOURCE = '🍽️ LunchMenuCardAdaptiveCardExtension';
  private _deferredPropertyPane: LunchMenuCardPropertyPane | undefined;
  private _menuService = new LunchMenuService();

  public async onInit(): Promise<void> {
    this.state = { locations: [], menuItems: [], selectedLocation: '', listReady: false };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(LOCATION_PICKER_VIEW_REGISTRY_ID, () => new LocationPickerView());
    this.quickViewNavigator.register(MENU_VIEW_REGISTRY_ID, () => new MenuView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      await this._menuService.Init(siteUrl, this.context.pageContext);
      const locations = await this._menuService.getLocations();
      const menuItems = await this._menuService.getAllItems();
      this.setState({ locations, menuItems, listReady: true });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
    }

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'LunchMenuCard-property-pane'*/
      './LunchMenuCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.LunchMenuCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
