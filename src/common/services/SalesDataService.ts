import { SPFx, spfi, SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/batching";
import { ISalesDataItem, ISalesDataPoint, ISalesDataSeries, SalesDataFields } from "../models/models";
import { Lists } from "../models/enums";
import { ConfigService } from "./ConfigService";

const LIST_DESCRIPTION = "Sales data by salesperson, region, and product";
const SERIES_COLORS = [
  "#0078d4", "#107c10", "#d13438", "#ffb900",
  "#8764b8", "#038387", "#ca5010", "#00b7c3"
];

function pad2(n: number): string {
  return n < 10 ? "0" + String(n) : String(n);
}

export class SalesDataService {
  private LOG_SOURCE: string = "📈 SalesDataService";
  private _sp!: SPFI;
  private _configService = new ConfigService();

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      await this._configService.Init(siteUrl, pageContext);
      await this._configService._configList(Lists.SALESDATA, LIST_DESCRIPTION, SalesDataFields);
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getSalesItems(daysBack: number = 90): Promise<ISalesDataItem[]> {
    const retVal: ISalesDataItem[] = [];
    try {
      const cutoff = new Date();
      cutoff.setDate(cutoff.getDate() - daysBack);
      const cutoffIso = cutoff.toISOString();

      const items = await this._sp.web.lists
        .getByTitle(Lists.SALESDATA)
        .items.select("Id", "Title", "Sale_x0020_Date", "Amount", "Salesperson", "Region", "Product")
        .filter(`Sale_x0020_Date ge '${cutoffIso}'`)
        .orderBy("Sale_x0020_Date", true)
        .top(2000)();

      items.forEach((item) => {
        retVal.push({
          id: String(item.Id),
          title: item.Title ?? "",
          saleDate: item.Sale_x0020_Date ?? "",
          amount: item.Amount ?? 0,
          salesperson: item.Salesperson ?? "Unknown",
          region: item.Region ?? "",
          product: item.Product ?? ""
        });
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getSalesItems) - ${err}`);
    }
    return retVal;
  }

  public buildSeries(items: ISalesDataItem[], daysBack: number, topN: number = 8, groupBy: 'salesperson' | 'region' | 'product' = 'salesperson'): ISalesDataSeries[] {
    
    const groupByValues:string[] = [];
    const series: ISalesDataSeries[] = [];
    for (const item of items) {
      if (!groupByValues.find(i => i === item[groupBy])) {
          groupByValues.push(item[groupBy]);
        }
    }

    for (const item of groupByValues) {
      const groupItems = items.filter(i => i[groupBy] === item);
      if (groupItems.length > 0) {
        const dataPoints: ISalesDataPoint[] = [];
        groupItems.map(gi => {
          const date = new Date(gi.saleDate);
          dataPoints.push({ date, amount: gi.amount });
        });
        series.push({ name: item, color: SERIES_COLORS[series.length % SERIES_COLORS.length], dataPoints, total: 0 });
      }
    }

    // const totals = new Map<string, number>();
    // const byPerson = new Map<string, Map<string, number>>();

    // const cutoff = new Date();
    // cutoff.setDate(cutoff.getDate() - daysBack);
    // cutoff.setHours(0, 0, 0, 0);

    // for (const item of items) {
    //   const date = new Date(item.saleDate);
    //   if (date < cutoff) continue;

    //   const dateKey = this._bucketKey(date, daysBack);
    //   const person = (groupBy === 'region' ? item.region : groupBy === 'product' ? item.product : item.salesperson) || "Unknown";

    //   totals.set(person, (totals.get(person) ?? 0) + item.amount);

    //   if (!byPerson.has(person)) byPerson.set(person, new Map());
    //   const dateMap = byPerson.get(person)!;
    //   dateMap.set(dateKey, (dateMap.get(dateKey) ?? 0) + item.amount);
    // }

    // // Sort by total desc, take top N
    // const sorted = Array.from(totals.entries())
    //   .sort((a, b) => b[1] - a[1])
    //   .slice(0, topN);

    // const allBuckets = this._buildBuckets(daysBack);

    // return sorted.map(([name, total], idx) => {
    //   const dateMap = byPerson.get(name) ?? new Map<string, number>();
    //   const dataPoints: ISalesDataPoint[] = allBuckets.map(({ key, date }) => ({
    //     date,
    //     amount: dateMap.get(key) ?? 0
    //   }));
    //   return { name, color: SERIES_COLORS[idx % SERIES_COLORS.length], dataPoints, total };
    // });

    return series;
  }

  private _bucketKey(date: Date, daysBack: number): string {
    if (daysBack <= 30) {
      return `${date.getFullYear()}-${pad2(date.getMonth() + 1)}-${pad2(date.getDate())}`;
    }
    // Weekly bucket — key off the Sunday of that week
    const startOfWeek = new Date(date);
    startOfWeek.setDate(date.getDate() - date.getDay());
    return `${startOfWeek.getFullYear()}-${pad2(startOfWeek.getMonth() + 1)}-${pad2(startOfWeek.getDate())}`;
  }

  private _buildBuckets(daysBack: number): { key: string; date: Date }[] {
    const buckets: { key: string; date: Date }[] = [];
    const seen = new Set<string>();
    const end = new Date();

    const current = new Date();
    current.setDate(end.getDate() - daysBack);
    current.setHours(0, 0, 0, 0);

    const step = daysBack <= 30 ? 1 : 7;

    while (current <= end) {
      const key = this._bucketKey(current, daysBack);
      if (!seen.has(key)) {
        seen.add(key);
        const bucketDate = new Date(current);
        if (daysBack > 30) {
          bucketDate.setDate(current.getDate() - current.getDay());
        }
        buckets.push({ key, date: bucketDate });
      }
      current.setDate(current.getDate() + step);
    }

    return buckets;
  }
}
