import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/SearchResultsViewTemplate.json';
import { IInventoryItem } from '../../../common/models/models';
import {
  IInventorySearchCardProps,
  IInventorySearchCardState,
  ITEM_DETAIL_VIEW_ID,
  inventorySearchCardInstance
} from '../InventorySearchCardAdaptiveCardExtension';

export interface ISearchResultsViewData {
  query: string;
  results: IInventoryItem[];
  hasResults: boolean;
  hasQuery: boolean;
}

export class SearchResultsView extends BaseAdaptiveCardQuickView<
  IInventorySearchCardProps,
  IInventorySearchCardState,
  ISearchResultsViewData
> {
  private LOG_SOURCE: string = '📦 SearchResultsView';

  public get data(): ISearchResultsViewData {
    const { allItems, searchQuery } = this.state;
    const q = (searchQuery ?? '').toLowerCase().trim();

    const results = q
      ? allItems.filter(item =>
          item.title.toLowerCase().includes(q) ||
          item.sku.toLowerCase().includes(q)
        )
      : allItems;

    return {
      query: searchQuery ?? '',
      results,
      hasResults: results.length > 0,
      hasQuery: q.length > 0
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    try {
      if (action.type !== 'Submit') return;
      const data = (action.data ?? {}) as { id?: string; itemId?: string };

      if (data.id === 'selectItem' && data.itemId) {
        const item = this.state.allItems.find((i: IInventoryItem) => i.id === data.itemId);
        if (item) {
          inventorySearchCardInstance.selectItem(item);
          this.quickViewNavigator.push(ITEM_DETAIL_VIEW_ID);
        }
      } else if (data.id === 'back') {
        this.quickViewNavigator.pop();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onAction) - ${err}`);
    }
  }
}
