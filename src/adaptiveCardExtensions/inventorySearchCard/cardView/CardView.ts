import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  SearchCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'InventorySearchCardAdaptiveCardExtensionStrings';
import {
  IInventorySearchCardProps,
  IInventorySearchCardState,
  SEARCH_RESULTS_VIEW_ID,
  inventorySearchCardInstance
} from '../InventorySearchCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IInventorySearchCardProps,
  IInventorySearchCardState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { searchQuery, allItems, isLoading } = this.state;

    const resultCount = allItems.filter(item => {
      if (!searchQuery) return false;
      const q = searchQuery.toLowerCase();
      return item.title.toLowerCase().includes(q) || item.sku.toLowerCase().includes(q);
    }).length;

    const footerText = isLoading
      ? 'Loading inventory...'
      : searchQuery
        ? `${resultCount} result${resultCount !== 1 ? 's' : ''} found`
        : `${allItems.length} items in inventory`;

    return SearchCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: strings.CardHeader
      },
      body: {
        componentName: 'searchBox',
        placeholder: strings.SearchPlaceholder,
        value: searchQuery,
        button: {
          action: {
            type: 'QuickView',
            parameters: { view: SEARCH_RESULTS_VIEW_ID }
          }
        },
        onChange: (newValue?: string) => {
          inventorySearchCardInstance.search(newValue ?? '');
        }
      },
      footer: {
        componentName: 'searchFooter',
        title: strings.FooterTitle,
        text: footerText,
        onSelection: {
          type: 'QuickView',
          parameters: { view: SEARCH_RESULTS_VIEW_ID }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: SEARCH_RESULTS_VIEW_ID }
    };
  }
}
