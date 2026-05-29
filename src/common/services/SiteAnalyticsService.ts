import { SPFx as spSPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import { GraphFI, graphfi, SPFx as graphSPFx, GraphQueryable, graphGet } from "@pnp/graph";
import "@pnp/graph/sites";
import "@pnp/graph/analytics";
import { ItemActivityStat as IItemActivityStat } from "@microsoft/microsoft-graph-types";
import { AdaptiveCardExtensionContext } from "@microsoft/sp-adaptive-card-extension-base";
import { ISiteAnalyticsSummary, ISiteActivityStat, SiteAnalyticsSummary } from "../models/models";
import strings from "SiteAnalyticsCardAdaptiveCardExtensionStrings";

export class SiteAnalyticsService {
  private LOG_SOURCE: string = '📊 SiteAnalyticsService';
  private _sp!: SPFI;
  private _graph!: GraphFI;
  private _rootSite!: string;
  private _ready: boolean = false;

  constructor(private context: AdaptiveCardExtensionContext) {}

  public async Init(rootSite: string, pageContext: any, aadTokenProviderFactory: any): Promise<void> {
    try {
      this._rootSite = rootSite;
      this._sp = spfi(this._rootSite).using(spSPFx({ pageContext }));
      this._graph = graphfi().using(graphSPFx({ aadTokenProviderFactory }));
      this._ready = true;
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getSiteAnalytics(siteUrl: string): Promise<SiteAnalyticsSummary> {
    let retVal: SiteAnalyticsSummary = new SiteAnalyticsSummary();
    try {
      const url: URL = new URL(siteUrl || this._rootSite);
      const site = url.pathname === '/'
        ? this._graph.sites.getById(url.hostname)
        : await this._graph.sites.getByUrl(url.hostname, url.pathname);

      const last7: IItemActivityStat = await site.analytics({timeRange: "lastSevenDays"});
      const last7Activity = await this._getActivityStats(site, last7.startDateTime, last7.endDateTime);
      const allTime: IItemActivityStat = await site.analytics({timeRange: "allTime"});
      
      const startDate = new Date();
      startDate.setDate(startDate.getDate() - 6);

      let dailyStats: ISiteActivityStat[] = [];
      if(last7Activity && last7Activity.length >0){
        dailyStats = this._buildDailyStats(last7Activity, startDate, 7);
      }
      

      retVal.last7Views = last7.access?.actionCount   ?? 0;
      retVal.last7Visitors = last7.access?.actorCount    ?? 0;
      retVal.allTimeViews = allTime.access?.actionCount  ?? 0;
      retVal.allTimeVisitors = allTime.access?.actorCount   ?? 0;
      retVal.dailyStats = dailyStats;

    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getSiteAnalytics) - ${err}`);
    }

    return retVal;
  }

  private async _getActivityStats(site: any, startDate: string | null | undefined, endDate: string | null | undefined): Promise<ISiteActivityStat[]> {
    let retVal: ISiteActivityStat[] = [];
    try{
      const query = `getActivitiesByInterval(startDateTime='${startDate}',endDateTime='${endDate}',interval='day')`;
      const activity = await graphGet(GraphQueryable(site, query));
      retVal = activity;
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_getActivityStats) - ${err}`);
    }
    return retVal;
  }

  private _buildDailyStats(raw: any[], start: Date, days: number): ISiteActivityStat[] {
    const retVal: ISiteActivityStat[] = [];
    try{
      const rawStats = [];
      for (let i = 0; i < raw.length; i++) {
        const stat = raw[i];
        const key = (stat.startDateTime as string)?.substring(0, 10) ?? '';
        rawStats.push({
          key: key,
          value: {
            views: (stat.access?.actionCount as number) ?? 0,
            visitors: (stat.access?.actorCount as number) ?? 0
          }
        });
      }

      const dayLabels = strings.WeekDayAbreviations;
      const d = new Date(start);
      for (let i = 0; i < days; i++) {
         d.setDate(d.getDate() + 1);
         const key = d.toISOString().substring(0, 10);
         const entry = rawStats.find(item => item.key === key);
         retVal.push({label: dayLabels[d.getDay()], date: entry?.key ?? key, viewCount: entry?.value?.views ?? 0, visitorCount: entry?.value?.visitors ?? 0});
       }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_buildDailyStats) - ${err}`);
    }
    return retVal;
  }

  public _generateChartUrl(stats: ISiteActivityStat[]): string {
    if (!stats || stats.length === 0) return '';

    const W = 380, H = 140, padX = 8, padBot = 24, barGap = 4;
    const max = Math.max(...stats.map(s => s.viewCount), 1);
    const barW = (W - padX * 2) / stats.length;

    const bars = stats.map((s, i) => {
      const barH = Math.max(Math.round(((H - padBot - 4) * s.viewCount) / max), s.viewCount > 0 ? 2 : 0);
      const x = padX + i * barW + barGap / 2;
      const y = H - padBot - barH;
      const labelX = x + (barW - barGap) / 2;
      return [
        `<rect x="${x}" y="${y}" width="${barW - barGap}" height="${barH}" fill="#0078d4" rx="3"/>`,
        s.viewCount > 0
          ? (barH >= 16
              ? `<text x="${labelX}" y="${y + 11}" text-anchor="middle" font-family="Segoe UI,sans-serif" font-size="9" fill="#ffffff">${s.viewCount}</text>`
              : `<text x="${labelX}" y="${y - 3}" text-anchor="middle" font-family="Segoe UI,sans-serif" font-size="9" fill="#323130">${s.viewCount}</text>`)
          : '',
        `<text x="${labelX}" y="${H - 6}" text-anchor="middle" font-family="Segoe UI,sans-serif" font-size="9" fill="#605e5c">${s.label}</text>`
      ].join('');
    }).join('');

    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" style="background:#f3f2f1;border-radius:4px">${bars}</svg>`;
    return `data:image/svg+xml;base64,${btoa(svg)}`;
  }
}
