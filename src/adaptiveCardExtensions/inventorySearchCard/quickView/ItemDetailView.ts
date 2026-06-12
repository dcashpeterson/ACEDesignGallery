import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/ItemDetailViewTemplate.json';
import {
  IInventorySearchCardProps,
  IInventorySearchCardState
} from '../InventorySearchCardAdaptiveCardExtension';

export interface IItemDetailViewData {
  id: string;
  title: string;
  sku: string;
  category: string;
  quantity: number;
  location: string;
  description: string;
  unitPrice: string;
  isLowStock: boolean;
}

export class ItemDetailView extends BaseAdaptiveCardQuickView<
  IInventorySearchCardProps,
  IInventorySearchCardState,
  IItemDetailViewData
> {
  private LOG_SOURCE: string = '📦 ItemDetailView';

  public get data(): IItemDetailViewData {
    const item = this.state.selectedItem;
    const price = item?.unitPrice ?? 0;
    return {
      id: item?.id ?? '',
      title: item?.title ?? '',
      sku: item?.sku ?? '',
      category: item?.category ?? '',
      quantity: item?.quantity ?? 0,
      location: item?.location ?? '',
      description: item?.description ?? '',
      unitPrice: `$${price.toFixed(2)}`,
      isLowStock: (item?.quantity ?? 0) < 10
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type === 'Submit' && (action.data as { id?: string })?.id === 'back') {
        this.quickViewNavigator.pop();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
