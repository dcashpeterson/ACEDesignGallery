import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import { ICompanyAnnouncement, CompanyAnnouncementsFields } from '../models/models';
import { Lists } from '../models/enums';
import { ConfigService } from './ConfigService';

const LIST_DESCRIPTION = 'Company-wide announcements and updates';

export class CompanyAnnouncementsService {
  private LOG_SOURCE = '📢 CompanyAnnouncementsService';
  private _sp!: SPFI;
  private _configService = new ConfigService();

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      await this._configService.Init(siteUrl, pageContext);
      const created = await this._configService._configList(Lists.COMPANYANNOUNCEMENTS, LIST_DESCRIPTION, CompanyAnnouncementsFields);
      if (created) {
        await this._seedSampleData();
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getAnnouncements(): Promise<ICompanyAnnouncement[]> {
    try {
      const items = await this._sp.web.lists
        .getByTitle(Lists.COMPANYANNOUNCEMENTS)
        .items.select('Id', 'Title', 'AnnouncementBody', 'Category', 'Published_x0020_Date')
        .orderBy('Published_x0020_Date', false)
        .top(20)();

      if (items.length === 0) return this._getSampleData();

      return items.map((item: any) => this._mapItem(item));
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (getAnnouncements) - ${err}`);
      return this._getSampleData();
    }
  }

  private async _seedSampleData(): Promise<void> {
    try {
      const list = this._sp.web.lists.getByTitle(Lists.COMPANYANNOUNCEMENTS);
      const samples = this._getSampleData();
      for (const item of samples) {
        await list.items.add({
          Title: item.title,
          AnnouncementBody: item.body,
          Category: item.category,
          Published_x0020_Date: item.publishedDate
        });
      }
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (_seedSampleData) - ${err}`);
    }
  }

  private _mapItem(item: any): ICompanyAnnouncement {
    return {
      id: String(item.Id),
      title: item.Title ?? '',
      body: this._stripHtml(item.AnnouncementBody ?? ''),
      publishedDate: item.PublishedDate ?? '',
      publishedDateFormatted: this._formatDate(item.PublishedDate ?? ''),
      category: item.Category ?? ''
    };
  }

  private _formatDate(isoDate: string): string {
    if (!isoDate) return '';
    try {
      return new Date(isoDate).toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    } catch {
      return isoDate;
    }
  }

  private _stripHtml(html: string): string {
    return html.replace(/<[^>]*>/g, '').trim();
  }

  private _getSampleData(): ICompanyAnnouncement[] {
    return [
      {
        id: '1',
        title: 'Q2 All-Hands Meeting Scheduled',
        body: 'Join us on June 15th for our quarterly all-hands meeting. We will review Q1 results, share our Q2 roadmap, and celebrate recent wins as a team. All employees are encouraged to attend in person or via Teams. Lunch will be provided for those joining on-site.',
        publishedDate: '2026-06-01T00:00:00Z',
        publishedDateFormatted: 'June 1, 2026',
        category: 'General'
      },
      {
        id: '2',
        title: 'New Employee Benefits Portal Launched',
        body: 'We are excited to announce the launch of our new Employee Benefits Portal. The new portal makes it easier than ever to review your benefits, update dependents, submit claims, and access wellness resources. Log in with your company credentials at benefits.company.com. HR will host drop-in sessions this week to answer any questions.',
        publishedDate: '2026-05-20T00:00:00Z',
        publishedDateFormatted: 'May 20, 2026',
        category: 'HR'
      },
      {
        id: '3',
        title: 'Office Renovation Complete — West Wing Now Open',
        body: 'The West Wing renovation is now complete! Teams who were temporarily relocated can begin moving back to their permanent workspaces starting Monday. The renovated floor includes updated collaboration spaces, phone booths for focus work, and a fully refreshed kitchen area. A walk-through tour will be available Friday from 12–2pm.',
        publishedDate: '2026-05-10T00:00:00Z',
        publishedDateFormatted: 'May 10, 2026',
        category: 'Operations'
      },
      {
        id: '4',
        title: 'Updated Remote Work Policy Effective July 1',
        body: 'Following feedback from our annual engagement survey, we have updated our Remote Work Policy effective July 1, 2026. Key changes include expanded flexibility for fully remote roles, a revised in-office minimum for hybrid employees, and clearer guidance on home office equipment reimbursement. Please review the full policy document in the HR portal.',
        publishedDate: '2026-04-28T00:00:00Z',
        publishedDateFormatted: 'April 28, 2026',
        category: 'General'
      },
      {
        id: '5',
        title: 'Planned IT Maintenance Window — June 7',
        body: 'Our IT team will perform scheduled maintenance on Saturday, June 7, from 10pm to 2am. During this window, email, SharePoint, and Teams may experience intermittent downtime. Please save and close any open documents before 10pm Friday. Critical system alerts will be communicated via SMS. Contact the IT helpdesk at ext. 5000 if you have questions.',
        publishedDate: '2026-04-15T00:00:00Z',
        publishedDateFormatted: 'April 15, 2026',
        category: 'IT'
      }
    ];
  }
}
