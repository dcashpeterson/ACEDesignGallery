import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import "@pnp/sp/items";
import "@pnp/sp/attachments"; 
import { ILunchMenuItem, ILocation, LunchMenuFields } from '../models/models';
import { ConfigService } from "./ConfigService";
import { Lists } from "../models/enums";
import strings from "LunchMenuCardAdaptiveCardExtensionStrings";

export class LunchMenuService {
  private _sp!: SPFI;
  private _configService = new ConfigService();
  private LOG_SOURCE = '🍽️ LunchMenuService';
  private _siteUrl:string = "";

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      this._siteUrl = siteUrl;
      await this._configService.Init(siteUrl,pageContext);
      const configReady = await this._configService._configList(Lists.DEMOLUNCHLIST,strings.LunchMenuDescription,LunchMenuFields);
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getLocations(): Promise<ILocation[]> {
    let retVal: ILocation[] = [];
    try {
      const items = await this._sp.web.lists.getByTitle(Lists.DEMOLUNCHLIST).items.select('ID','Location', 'Image')();
      items.forEach(item => {
        if (!retVal.find(l => l.name === item.Location)) {
          const imageData = JSON.parse(item.Image);
          retVal.push({ name: item.Location, imageUrl: `${this._siteUrl}/Lists/${Lists.DEMOLUNCHLIST}/Attachments/${item.ID}/${imageData.fileName}` });
        }
      });
      retVal.sort((a, b) => a.name.localeCompare(b.name));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getLocations) - ${err}`);
    }
    return retVal;
  }

  public async getAllItems(): Promise<ILunchMenuItem[]> {
    let retVal: ILunchMenuItem[] = [];
    try {
      const menuItems = await this._sp.web.lists.getByTitle(Lists.DEMOLUNCHLIST).items.select('Id', 'Title', 'Location', 'Description', 'Category', 'Price').top(500)();
      
      menuItems.map(item =>{
        
        return retVal.push(
          {
          id: String(item.Id),
          title: item.Title ?? '',
          location: item.Location ?? '',
          description: item.Description ?? '',
          category: item.Category ?? '',
          price: item.Price ?? 0,
          image:''
          });
        });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getAllItems) - ${err}`);
    }
    return retVal;
  }
}
