import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { IInventoryItem, InventoryItemFields } from '../models/models';
import { Lists } from '../models/enums';
import { ConfigService } from './ConfigService';

const LIST_DESCRIPTION = 'Inventory items with stock levels, locations, and pricing';

export class InventoryService {
  private LOG_SOURCE: string = '📦 InventoryService';
  private _sp!: SPFI;
  private _configService = new ConfigService();

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      await this._configService.Init(siteUrl, pageContext);
      const created = await this._configService._configList(Lists.INVENTORY, LIST_DESCRIPTION, InventoryItemFields);
      if (created) {
        await this._seedSampleData();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getAllItems(): Promise<IInventoryItem[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle(Lists.INVENTORY)
        .items.select('Id', 'Title', 'SKU', 'Category', 'Quantity', 'Location', 'Description', 'Unit_x0020_Price')
        .orderBy('Title', true)
        .top(500)();

      if (items.length === 0) return this._getSampleData();

      return items.map((item: any) => this._mapItem(item));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getAllItems) - ${err}`);
      return this._getSampleData();
    }
  }

  private _mapItem(item: any): IInventoryItem {
    return {
      id: String(item.Id),
      title: item.Title ?? '',
      sku: item.SKU ?? '',
      category: item.Category ?? '',
      quantity: item.Quantity ?? 0,
      location: item.Location ?? '',
      description: item.Description ?? '',
      unitPrice: item.UnitPrice ?? 0
    };
  }

  private async _seedSampleData(): Promise<void> {
    try {
      const list = this._sp.web.lists.getByTitle(Lists.INVENTORY);
      const samples = this._getSampleData();
      for (const item of samples) {
        await list.items.add({
          Title: item.title,
          SKU: item.sku,
          Category: item.category,
          Quantity: item.quantity,
          Location: item.location,
          Description: item.description,
          UnitPrice: item.unitPrice
        });
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_seedSampleData) - ${err}`);
    }
  }

  private _getSampleData(): IInventoryItem[] {
    return [
      { id: '1',  title: 'Laptop - Dell Latitude 5540',    sku: 'ELEC-0001', category: 'Electronics',      quantity: 12,  location: 'Warehouse A - Shelf E1', description: '15.6" business laptop, Intel Core i7, 16GB RAM, 512GB SSD',    unitPrice: 1249.99 },
      { id: '2',  title: 'Wireless Keyboard',              sku: 'ELEC-0002', category: 'Electronics',      quantity: 45,  location: 'Warehouse A - Shelf E2', description: 'Full-size wireless keyboard with number pad, Bluetooth 5.0',   unitPrice: 49.95  },
      { id: '3',  title: 'USB-C Docking Station',          sku: 'ELEC-0003', category: 'Electronics',      quantity: 8,   location: 'Warehouse A - Shelf E3', description: 'Triple-display USB-C dock with 100W charging pass-through',    unitPrice: 199.00 },
      { id: '4',  title: 'Monitor - 27" 4K Display',       sku: 'ELEC-0004', category: 'Electronics',      quantity: 5,   location: 'Warehouse A - Shelf E4', description: '27-inch 4K IPS monitor, USB-C, 60Hz, VESA compatible',         unitPrice: 399.00 },
      { id: '5',  title: 'Power Drill - Cordless 18V',     sku: 'TOOL-0001', category: 'Tools',            quantity: 7,   location: 'Warehouse B - Shelf T1', description: '18V lithium-ion cordless drill, 2-speed, includes 2 batteries', unitPrice: 89.99  },
      { id: '6',  title: 'Socket Wrench Set - 40pc',       sku: 'TOOL-0002', category: 'Tools',            quantity: 15,  location: 'Warehouse B - Shelf T2', description: 'SAE and metric socket set, 1/4" and 3/8" drive, chrome vanadium', unitPrice: 54.99 },
      { id: '7',  title: 'Tape Measure - 25ft',            sku: 'TOOL-0003', category: 'Tools',            quantity: 30,  location: 'Warehouse B - Shelf T3', description: 'Heavy-duty 25-foot retractable tape measure with belt clip',    unitPrice: 14.99  },
      { id: '8',  title: 'Copy Paper - Case (10 Reams)',   sku: 'OFFC-0001', category: 'Office Supplies',  quantity: 60,  location: 'Storage Room 1 - Bay O1', description: '8.5x11 white copy paper, 20lb, 500 sheets per ream, case of 10', unitPrice: 42.50 },
      { id: '9',  title: 'Ballpoint Pens - Box of 36',     sku: 'OFFC-0002', category: 'Office Supplies',  quantity: 120, location: 'Storage Room 1 - Bay O2', description: 'Medium point blue ballpoint pens, retractable, box of 36',     unitPrice: 12.99  },
      { id: '10', title: 'Sticky Notes - 12 Pack',         sku: 'OFFC-0003', category: 'Office Supplies',  quantity: 85,  location: 'Storage Room 1 - Bay O3', description: '3x3 yellow sticky notes, 100 sheets per pad, 12-pack',        unitPrice: 8.49   },
      { id: '11', title: 'Cable Ties - Pack of 100',       sku: 'HDWR-0001', category: 'Hardware',         quantity: 200, location: 'Warehouse C - Bin H1',    description: 'Nylon cable ties, 8-inch, self-locking, 100-pack',            unitPrice: 6.99   },
      { id: '12', title: 'Hex Bolt Set - M6 Assortment',  sku: 'HDWR-0002', category: 'Hardware',         quantity: 3,   location: 'Warehouse C - Bin H2',    description: '240-piece M6 hex bolt, nut, and washer assortment kit',       unitPrice: 22.50  },
      { id: '13', title: 'Extension Cord - 25ft 12AWG',   sku: 'HDWR-0003', category: 'Hardware',         quantity: 18,  location: 'Warehouse C - Bin H3',    description: '25-foot 12AWG 3-outlet heavy-duty extension cord, grounded',  unitPrice: 29.99  },
      { id: '14', title: 'Hard Hat - Type II',             sku: 'SAFE-0001', category: 'Safety Equipment', quantity: 22,  location: 'Safety Cabinet - Row S1', description: 'ANSI/ISEA Z89.1 Type II Class C vented hard hat, white',      unitPrice: 24.95  },
      { id: '15', title: 'Safety Glasses - Anti-Fog',      sku: 'SAFE-0002', category: 'Safety Equipment', quantity: 9,   location: 'Safety Cabinet - Row S2', description: 'ANSI Z87.1 anti-fog safety glasses, clear lens, wraparound',  unitPrice: 11.99  }
    ];
  }
}
