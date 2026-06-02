import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { IVacationRequest, IVacationSummary, VacationRequestFields } from "../models/models";
import { Lists } from "../models/enums";
import { ConfigService } from "./ConfigService";

const LIST_DESCRIPTION = "Employee time off requests by status";

export class VacationDaysService {
  private LOG_SOURCE: string = '🏖️ VacationDaysService';
  private _sp!: SPFI;
  private _configService = new ConfigService();

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      await this._configService.Init(siteUrl, pageContext);
      await this._configService._configList(Lists.VACATIONREQUESTS, LIST_DESCRIPTION, VacationRequestFields);
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getMyRequests(authorId: number): Promise<IVacationRequest[]> {
    const retVal: IVacationRequest[] = [];
    try {
      const items = await this._sp.web.lists
        .getByTitle(Lists.VACATIONREQUESTS)
        .items.filter(`AuthorId eq ${authorId}`)
        .select('Id', 'Title', 'Start_x0020_Date', 'End_x0020_Date', 'Number_x0020_of_x0020_Days', 'Status', 'Notes')
        .orderBy('Start_x0020_Date', false)
        .top(100)();

      items.forEach(item => {
        retVal.push({
          id: String(item.Id),
          title: item.Title ?? '',
          startDate: item.Start_x0020_Date ? item.Start_x0020_Date.substring(0, 10) : '',
          endDate: item.End_x0020_Date ? item.End_x0020_Date.substring(0, 10) : '',
          days: item.Number_x0020_of_x0020_Days ?? 0,
          status: item.Status ?? 'Submitted',
          notes: item.Notes ?? ''
        });
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getMyRequests) - ${err}`);
    }
    return retVal;
  }

  public async createRequest(title: string, startDate: string, endDate: string, days: number, notes: string): Promise<void> {
    try {
      await this._sp.web.lists
        .getByTitle(Lists.VACATIONREQUESTS)
        .items.add({
          Title: title,
          Start_x0020_Date: startDate ? startDate + 'T00:00:00Z' : null,
          End_x0020_Date: endDate ? endDate + 'T00:00:00Z' : null,
          Number_x0020_of_x0020_Days: days,
          Status: 'Submitted',
          Notes: notes
        });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (createRequest) - ${err}`);
      throw err;
    }
  }

  public getVacationSummary(requests: IVacationRequest[], totalDays: number): IVacationSummary {
    const used = requests
      .filter(r => r.status === 'Approved')
      .reduce((sum, r) => sum + r.days, 0);
    const available = Math.max(0, totalDays - used);
    return { used, available, total: totalDays };
  }
}
