import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { ISupportTicket, ITicketStatusCounts, SupportTicketFields } from "../models/models";
import { Lists } from "../models/enums";
import { ConfigService } from "./ConfigService";

const LIST_DESCRIPTION = "Support ticket tracking by status";

export class SupportTicketsService {
  private LOG_SOURCE: string = '🎫 SupportTicketsService';
  private _sp!: SPFI;
  private _configService = new ConfigService();

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      await this._configService.Init(siteUrl, pageContext);
      await this._configService._configList(Lists.SUPPORTTICKETS, LIST_DESCRIPTION, SupportTicketFields);
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getTickets(): Promise<ISupportTicket[]> {
    const retVal: ISupportTicket[] = [];
    try {
      const items = await this._sp.web.lists
        .getByTitle(Lists.SUPPORTTICKETS)
        .items.select('Id', 'Title', 'Status', 'Priority', 'Assigned_x0020_To/Title', 'Description', 'Submitted_x0020_Date')
        .expand('Assigned_x0020_To')
        .top(500)();

      items.forEach(item => {
        retVal.push({
          id: String(item.Id),
          title: item.Title ?? '',
          status: item.Status ?? 'Open',
          priority: item.Priority ?? 'Medium',
          assignedTo: item.Assigned_x0020_To?.Title ?? '',
          description: item.Description ?? '',
          submittedDate: item.Submitted_x0020_Date ?? ''
        });
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getTickets) - ${err}`);
    }
    return retVal;
  }

  public getStatusCounts(tickets: ISupportTicket[]): ITicketStatusCounts {
    return {
      open: tickets.filter(t => t.status === 'Open').length,
      inProgress: tickets.filter(t => t.status === 'In Progress').length,
      resolved: tickets.filter(t => t.status === 'Resolved').length,
      escalated: tickets.filter(t => t.status === 'Escalated').length
    };
  }

  public async getMyTickets(authorId: number): Promise<ISupportTicket[]> {
    const retVal: ISupportTicket[] = [];
    try {
      const items = await this._sp.web.lists
        .getByTitle(Lists.SUPPORTTICKETS)
        .items.filter(`AuthorId eq ${authorId}`)
        .select('Id', 'Title', 'Status', 'Priority', 'Assigned_x0020_To/Title', 'Description', 'Submitted_x0020_Date')
        .expand('Assigned_x0020_To')
        .top(500)();

      items.forEach(item => {
        retVal.push({
          id: String(item.Id),
          title: item.Title ?? '',
          status: item.Status ?? 'Open',
          priority: item.Priority ?? 'Medium',
          assignedTo: item.Assigned_x0020_To?.Title ?? '',
          description: item.Description ?? '',
          submittedDate: item.Submitted_x0020_Date ?? ''
        });
      });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getMyTickets) - ${err}`);
    }
    return retVal;
  }

  public async createTicket(title: string, priority: string, description: string): Promise<void> {
    try {
      await this._sp.web.lists
        .getByTitle(Lists.SUPPORTTICKETS)
        .items.add({
          Title: title,
          Status: 'Open',
          Priority: priority,
          Description: description,
          Submitted_x0020_Date: new Date().toISOString()
        });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (createTicket) - ${err}`);
      throw err;
    }
  }

  public generateDonutChartUrl(counts: ITicketStatusCounts): string {
    const total = counts.open + counts.inProgress + counts.resolved + counts.escalated;
    if (total === 0) return '';

    const W = 320, H = 180;
    const cx = 90, cy = 90, r = 72, innerR = 42;

    const segments = [
      { label: 'Open', value: counts.open, color: '#0078d4' },
      { label: 'In Progress', value: counts.inProgress, color: '#ffb900' },
      { label: 'Resolved', value: counts.resolved, color: '#107c10' },
      { label: 'Escalated', value: counts.escalated, color: '#d13438' }
    ];

    let paths = '';
    let startAngle = -Math.PI / 2;

    for (const seg of segments) {
      if (seg.value === 0) continue;
      const sliceAngle = (seg.value / total) * 2 * Math.PI;
      const endAngle = startAngle + sliceAngle;

      const x1 = cx + r * Math.cos(startAngle);
      const y1 = cy + r * Math.sin(startAngle);
      const x2 = cx + r * Math.cos(endAngle);
      const y2 = cy + r * Math.sin(endAngle);
      const xi1 = cx + innerR * Math.cos(endAngle);
      const yi1 = cy + innerR * Math.sin(endAngle);
      const xi2 = cx + innerR * Math.cos(startAngle);
      const yi2 = cy + innerR * Math.sin(startAngle);
      const largeArc = sliceAngle > Math.PI ? 1 : 0;

      paths += `<path d="M${x1.toFixed(2)},${y1.toFixed(2)} A${r},${r} 0 ${largeArc},1 ${x2.toFixed(2)},${y2.toFixed(2)} L${xi1.toFixed(2)},${yi1.toFixed(2)} A${innerR},${innerR} 0 ${largeArc},0 ${xi2.toFixed(2)},${yi2.toFixed(2)} Z" fill="${seg.color}"/>`;

      startAngle = endAngle;
    }

    const centerText = `<text x="${cx}" y="${cy - 8}" text-anchor="middle" font-family="Segoe UI,sans-serif" font-size="20" font-weight="bold" fill="#323130">${counts.open}</text><text x="${cx}" y="${cy + 12}" text-anchor="middle" font-family="Segoe UI,sans-serif" font-size="11" fill="#605e5c">open</text>`;

    let legendY = 28;
    let legend = '';
    for (const seg of segments) {
      legend += `<rect x="196" y="${legendY}" width="10" height="10" fill="${seg.color}" rx="2"/><text x="212" y="${legendY + 9}" font-family="Segoe UI,sans-serif" font-size="11" fill="#323130">${seg.label} (${seg.value})</text>`;
      legendY += 24;
    }

    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" style="background:#f3f2f1;border-radius:4px">${paths}${centerText}${legend}</svg>`;
    return `data:image/svg+xml;base64,${btoa(svg)}`;
  }
}
