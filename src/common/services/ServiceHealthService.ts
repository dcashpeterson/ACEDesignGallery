import { SPFx as spSPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import { GraphFI, graphfi, SPFx as graphSPFx } from "@pnp/graph";
import { AdaptiveCardExtensionContext } from '@microsoft/sp-adaptive-card-extension-base';
import '@pnp/graph/admin';
import { IServiceHealth, IServiceHealthIssue } from '../models/models';
import { statusType } from "../models/enums";

export class ServiceHealthService {
  private LOG_SOURCE: string = '🔴 ServiceHealthService';
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
      console.error(`${this.LOG_SOURCE} (render) - ${err}`);
    }
  }

  public async getServices(): Promise<IServiceHealth[]> {
    let retVal: IServiceHealth[] = [];
    try {
      const overviews = await this._graph.admin.serviceAnnouncements.healthOverviews();

      if (overviews && overviews.length > 0) {
        overviews.map(item => { 
          if(item.status && item.status !== 'serviceOperational') {
            //const issues: IServiceHealthIssue[] = await this.getIssuesByService(item.service);
            return retVal.push(
              {
                id: item.id ?? '',
                title: item.service ?? '',
                service: item.service ?? '',
                statusType: statusType[item.status as keyof typeof statusType] ?? 0 as unknown as statusType
              });
          }
        });
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (render) - ${err}`);
    }

    return retVal;
  }

  public async getIssuesByService(service?: string): Promise<IServiceHealthIssue[]> {
    let retVal: IServiceHealthIssue[] = [];
    try {
      const issues = await this._graph.admin.serviceAnnouncements.issues();

      if (issues && issues.length > 0) {
        const filtered = service
          ? issues.filter(i => i.service === service)
          : issues;

        filtered.map(item => {
          return retVal.push({
            id: item.id ?? '',
            title: item.title ?? '',
            service: item.service ?? '',
            statusType: statusType[item.status as keyof typeof statusType] ?? 0 as unknown as statusType,
            classification: item.classification ?? '',
            feature: item.feature ?? '',
            featureGroup: item.featureGroup ?? '',
            impactDescription: item.impactDescription ?? '',
            isResolved: item.isResolved ?? false,
            origin: item.origin ?? '',
            startDateTime: item.startDateTime ?? '',
            endDateTime: item.endDateTime ?? '',
            lastModifiedDateTime: item.lastModifiedDateTime ?? '',
          });
        });
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (render) - ${err}`);
    }

    return retVal;
  }
}
