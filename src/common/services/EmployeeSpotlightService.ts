import { SPFx, spfi, SPFI } from "@pnp/sp";
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/fields';
import '@pnp/sp/content-types';
import { IEmployeeSpotlight, EmployeeSpotlightFields } from '../models/models';
import { CalendarType, DateTimeFieldFormatType, DateTimeFieldFriendlyFormatType, UrlFieldFormatType } from "@pnp/sp/fields";

const LOG_SOURCE = '🌟 EmployeeSpotlightService';

export class EmployeeSpotlightService {
  private _sp!: SPFI;
  private _siteUrl: string = '';
  private readonly CT_PARENT_ID = '0x0101009D1CB255DA76424F860D91F20E6C4118';
  private readonly CT_NAME = 'Employee Spotlight';
  private readonly SITE_PAGES = 'Site Pages';

  public async Init(siteUrl: string, pageContext: any): Promise<void> {
    try {
      this._sp = spfi(siteUrl).using(SPFx({ pageContext }));
      this._siteUrl = siteUrl;
      await this.provisionContentType();
    } catch (err) {
      console.error(`${LOG_SOURCE} (Init) - ${err}`);
    }
  }

  public async getSpotlights(): Promise<IEmployeeSpotlight[]> {
    try {
      const items = await this._sp.web.lists.getByTitle(this.SITE_PAGES).items
        .select('Id', 'Title', 'SpotlightEmployeeName', 'SpotlightType', 'SpotlightEmployeeImage', 'SpotlightDate', 'SpotlightArticle', 'FileRef')
        .filter(`startswith(ContentTypeId,'${this.CT_PARENT_ID}') and SpotlightEmployeeName ne null`)
        .orderBy('SpotlightDate', false)
        .top(20)();

      if (items.length === 0) return this._getSampleData();
      return items.map((item: any) => this._mapItem(item));
    } catch (err) {
      console.error(`${LOG_SOURCE} (getSpotlights) - ${err}`);
      return this._getSampleData();
    }
  }

  private async provisionContentType(): Promise<void> {
    try {
      const webInfo = await this._sp.web();
      const groupName = `_${webInfo.Title}`;

      // Step A: Add web-level site columns if not already present
      let webFields: any[] = [];
      try {
        webFields = await this._sp.web.fields.select('InternalName')();
      } catch { /* ignore */ }
      const existingFieldNames = new Set(webFields.map((f: any) => f.InternalName));

      for (const fieldDef of EmployeeSpotlightFields) {
        if (existingFieldNames.has(fieldDef.internalName)) continue;
        try {
          if (fieldDef.props.FieldTypeKind === 2) {
            await this._sp.web.fields.addText(fieldDef.internalName, { Group: groupName, Title: fieldDef.displayName });
          } else if (fieldDef.props.FieldTypeKind === 6) {
            await this._sp.web.fields.addChoice(fieldDef.internalName, {
              Choices: fieldDef.props.choices ?? [],
              EditFormat: fieldDef.props.editFormat,
              Group: groupName,
              Title: fieldDef.displayName
            });
          } else if (fieldDef.props.FieldTypeKind === 11) {
            await this._sp.web.fields.addUrl(fieldDef.internalName, {
              DisplayFormat: UrlFieldFormatType.Hyperlink,
              Group: groupName,
              Title: fieldDef.displayName
            });
          } else if (fieldDef.props.FieldTypeKind === 4) {
            await this._sp.web.fields.addDateTime(fieldDef.internalName, {
              DisplayFormat: DateTimeFieldFormatType.DateOnly,
              DateTimeCalendarType: CalendarType.Gregorian,
              FriendlyDisplayFormat: DateTimeFieldFriendlyFormatType.Disabled,
              Group: groupName,
              Title: fieldDef.displayName
            });
          } else if (fieldDef.props.FieldTypeKind === 3) {
            await this._sp.web.fields.createFieldAsXml(
              `<Field Type="Note" Name="${fieldDef.internalName}" DisplayName="${fieldDef.displayName}" Required="FALSE" RichText="${fieldDef.props.richText}" RichTextMode="FullHtml" Group="${groupName}" />`
            );
          }
        } catch (fieldErr) {
          console.warn(`${LOG_SOURCE} (provisionContentType) - field ${fieldDef.internalName}: ${fieldErr}`);
        }
      }

      // Step B: Create web-level content type inheriting from Site Page
      let webCTs: any[] = [];
      try {
        webCTs = await this._sp.web.contentTypes.select('StringId', 'Name')();
      } catch { /* ignore */ }

      const ctExists = webCTs.some((ct: any) => ct.Name === this.CT_NAME);
      if (!ctExists) {
        try {
          await this._sp.web.contentTypes.add(this.CT_PARENT_ID, this.CT_NAME, 'Employee recognition story');
          // Refresh after adding
          webCTs = await this._sp.web.contentTypes.select('StringId', 'Name')();
        } catch (ctErr) {
          console.warn(`${LOG_SOURCE} (provisionContentType) - create CT: ${ctErr}`);
        }
      }

      // Step C: Add FieldLinks to the content type
      const ctInfo = webCTs.find((ct: any) => ct.Name === this.CT_NAME);
      if (ctInfo) {
        try {
          const ct = this._sp.web.contentTypes.getById(ctInfo.StringId);
          const existingLinks = await ct.fieldLinks.select('Name')();
          const existingLinkNames = new Set(existingLinks.map((fl: any) => fl.Name));

          for (const fieldDef of EmployeeSpotlightFields) {
            if (!existingLinkNames.has(fieldDef.internalName)) {
              try {
                await (ct.fieldLinks as any).add({ FieldInternalName: fieldDef.internalName });
              } catch (flErr) {
                console.warn(`${LOG_SOURCE} (provisionContentType) - fieldLink ${fieldDef.internalName}: ${flErr}`);
              }
            }
          }
        } catch (linkErr) {
          console.warn(`${LOG_SOURCE} (provisionContentType) - FieldLinks: ${linkErr}`);
        }

        // Step D: Enable content types on Site Pages library and add the CT
        try {
          const lib = this._sp.web.lists.getByTitle(this.SITE_PAGES);
          const libInfo = await lib.select('ContentTypesEnabled')();
          if (!libInfo.ContentTypesEnabled) {
            await lib.update({ ContentTypesEnabled: true });
          }
          const listCTs = await lib.contentTypes.select('Name')();
          if (!listCTs.some((ct: any) => ct.Name === this.CT_NAME)) {
            await lib.contentTypes.addAvailableContentType(ctInfo.StringId);
          }
        } catch (libErr) {
          console.warn(`${LOG_SOURCE} (provisionContentType) - Site Pages library: ${libErr}`);
        }
      }
    } catch (err) {
      console.error(`${LOG_SOURCE} (provisionContentType) - ${err}`);
    }
  }

  private _mapItem(item: any): IEmployeeSpotlight {
    const imageField = item.SpotlightEmployeeImage;
    const imageUrl = imageField ? (imageField.Url ?? imageField) : '';
    return {
      id: String(item.Id),
      employeeName: item.SpotlightEmployeeName ?? item.Title ?? '',
      spotlightType: item.SpotlightType ?? '',
      description: item.SpotlightType ?? '',
      employeeImageUrl: imageUrl,
      spotlightDate: item.SpotlightDate ?? '',
      spotlightDateFormatted: this._formatDate(item.SpotlightDate ?? ''),
      fullArticle: this._stripHtml(item.SpotlightArticle ?? ''),
      pageUrl: item.FileRef ? `${this._siteUrl}${item.FileRef}` : '#'
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

  private _getSampleData(): IEmployeeSpotlight[] {
    return [
      {
        id: '1',
        employeeName: 'Sarah Mitchell',
        spotlightType: 'Promotion',
        description: 'Promoted to Senior Engineering Manager',
        employeeImageUrl: 'https://randomuser.me/api/portraits/women/44.jpg',
        spotlightDate: '2026-05-01T00:00:00Z',
        spotlightDateFormatted: 'May 1, 2026',
        fullArticle: 'We are thrilled to announce that Sarah Mitchell has been promoted to Senior Engineering Manager on the Platform Infrastructure team. Over the past four years, Sarah has led three major cloud migrations, mentored over a dozen junior engineers, and consistently delivered complex projects ahead of schedule. Please join us in congratulating Sarah on this well-deserved achievement!',
        pageUrl: '#'
      },
      {
        id: '2',
        employeeName: 'James Okonkwo',
        spotlightType: 'Award',
        description: 'Q1 2026 Innovation Award recipient',
        employeeImageUrl: 'https://randomuser.me/api/portraits/men/32.jpg',
        spotlightDate: '2026-04-15T00:00:00Z',
        spotlightDateFormatted: 'April 15, 2026',
        fullArticle: 'Congratulations to James Okonkwo, this quarter\'s Innovation Award winner! James proposed and prototyped an AI-assisted triage system for our support queue that reduced average ticket resolution time by 38%. His idea started as a weekend hackathon project and is now being evaluated for full production rollout. James embodies our value of continuous improvement—always looking for smarter, faster ways to solve real customer problems.',
        pageUrl: '#'
      },
      {
        id: '3',
        employeeName: 'Linda Forsythe',
        spotlightType: 'Retirement',
        description: 'Celebrating 28 years of exceptional service',
        employeeImageUrl: 'https://randomuser.me/api/portraits/women/68.jpg',
        spotlightDate: '2026-03-28T00:00:00Z',
        spotlightDateFormatted: 'March 28, 2026',
        fullArticle: 'After 28 remarkable years, Linda Forsythe is embarking on a well-earned retirement. Linda joined our company as a Customer Success Representative in 1998 and rose through the ranks to become Director of Client Relations. She personally built the client onboarding program that still forms the backbone of our new customer experience today. Linda leaves behind a legacy of thousands of clients who felt genuinely cared for. We wish her all the best in her next chapter!',
        pageUrl: '#'
      },
      {
        id: '4',
        employeeName: 'Marcus Rivera',
        spotlightType: 'New Baby',
        description: 'Welcoming baby Eliana Rivera',
        employeeImageUrl: 'https://randomuser.me/api/portraits/men/55.jpg',
        spotlightDate: '2026-04-02T00:00:00Z',
        spotlightDateFormatted: 'April 2, 2026',
        fullArticle: 'The Rivera family has a new addition! Marcus Rivera, Lead UX Designer on our Product team, and his partner Sofia are overjoyed to welcome baby Eliana Rivera, born March 30th, 2026, weighing 7 lbs 4 oz and measuring 20 inches. Both mom and baby are doing wonderfully. Marcus will be on parental leave through June and will return refreshed and ready to design beautiful things. Send your congratulations his way!',
        pageUrl: '#'
      },
      {
        id: '5',
        employeeName: 'Priya Nambiar',
        spotlightType: 'Work Anniversary',
        description: 'Celebrating 10 years with the company',
        employeeImageUrl: 'https://randomuser.me/api/portraits/women/21.jpg',
        spotlightDate: '2026-05-10T00:00:00Z',
        spotlightDateFormatted: 'May 10, 2026',
        fullArticle: 'This month marks a decade of excellence from Priya Nambiar, our Principal Data Scientist. When Priya joined in 2016, we had a single analytics dashboard; today she leads a team of eight and oversees a data platform that processes over 50 million events daily. Along the way she has published two internal research papers, spoken at three industry conferences, and mentored the next generation of data practitioners. Here\'s to ten more years of insight and impact, Priya!',
        pageUrl: '#'
      }
    ];
  }
}
