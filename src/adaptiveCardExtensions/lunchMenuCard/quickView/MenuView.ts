import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/MenuViewTemplate.json';
import { ILunchMenuItem } from '../../../common/models/models';
import {
  ILunchMenuCardAdaptiveCardExtensionProps,
  ILunchMenuCardAdaptiveCardExtensionState,
  LOCATION_PICKER_VIEW_REGISTRY_ID
} from '../LunchMenuCardAdaptiveCardExtension';

export interface IMenuViewData {
  selectedLocation: string;
  menuItems: ILunchMenuItem[];
}

export class MenuView extends BaseAdaptiveCardQuickView<
  ILunchMenuCardAdaptiveCardExtensionProps,
  ILunchMenuCardAdaptiveCardExtensionState,
  IMenuViewData
> {
  private LOG_SOURCE: string = '🍽️ MenuView';

  public get data(): IMenuViewData {
    const menuItems = this.state.menuItems.filter(item => item.location === this.state.selectedLocation);
    return {
      selectedLocation: this.state.selectedLocation,
      menuItems: menuItems.filter(i => i.location === this.state.selectedLocation)
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
      try {
        if (action.type === 'Submit') {
          const { id } = action.data;
          if (id === 'back') {
            this.setState({ selectedLocation: '' });
            this.quickViewNavigator.pop();
          }
        }
      } catch (err) {
        console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
      }
    }
}
