import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { SearchResultsView } from './quickView/SearchResultsView';
import { ItemDetailView } from './quickView/ItemDetailView';
import { IInventoryItem } from '../../common/models/models';
import { InventoryService } from '../../common/services/InventoryService';

export interface IInventorySearchCardProps {
  title: string;
  listSiteUrl: string;
}

export interface IInventorySearchCardState {
  allItems: IInventoryItem[];
  searchQuery: string;
  selectedItem: IInventoryItem | undefined;
  isLoading: boolean;
}

export const CARD_VIEW_REGISTRY_ID: string = 'InventorySearchCard_CARD_VIEW';
export const SEARCH_RESULTS_VIEW_ID: string = 'InventorySearchCard_SEARCH_RESULTS_VIEW';
export const ITEM_DETAIL_VIEW_ID: string = 'InventorySearchCard_ITEM_DETAIL_VIEW';

export let inventorySearchCardInstance: InventorySearchCardAdaptiveCardExtension;

export default class InventorySearchCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IInventorySearchCardProps,
  IInventorySearchCardState
> {
  private LOG_SOURCE = '📦 InventorySearchCardAdaptiveCardExtension';
  private _deferredPropertyPane: any;
  private _inventoryService = new InventoryService();

  public async onInit(): Promise<void> {
    inventorySearchCardInstance = this;

    this.state = {
      allItems: [],
      searchQuery: '',
      selectedItem: undefined,
      isLoading: true
    };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(SEARCH_RESULTS_VIEW_ID, () => new SearchResultsView());
    this.quickViewNavigator.register(ITEM_DETAIL_VIEW_ID, () => new ItemDetailView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      await this._inventoryService.Init(siteUrl, this.context.pageContext);
      const allItems = await this._inventoryService.getAllItems();
      this.setState({ allItems, isLoading: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  public search(query: string): void {
    this.setState({ searchQuery: query });
  }

  public selectItem(item: IInventoryItem): void {
    this.setState({ selectedItem: item });
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'InventorySearchCard-property-pane'*/
      './InventorySearchCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.InventorySearchCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
