import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { AnnouncementDetailView } from './quickView/AnnouncementDetailView';
import { CompanyAnnouncementsCardPropertyPane } from './CompanyAnnouncementsCardPropertyPane';
import { ICompanyAnnouncement } from '../../common/models/models';
import { CompanyAnnouncementsService } from '../../common/services/CompanyAnnouncementsService';

export interface ICompanyAnnouncementsCardProps {
  title: string;
  listSiteUrl: string;
}

export interface ICompanyAnnouncementsCardState {
  announcements: ICompanyAnnouncement[];
  currentIndex: number;
  isLoading: boolean;
}

const CARD_VIEW_REGISTRY_ID: string = 'CompanyAnnouncements_CARD_VIEW';
export const DETAIL_VIEW_REGISTRY_ID: string = 'CompanyAnnouncements_DETAIL_VIEW';

export default class CompanyAnnouncementsCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ICompanyAnnouncementsCardProps,
  ICompanyAnnouncementsCardState
> {
  private LOG_SOURCE = '📢 CompanyAnnouncementsCardAdaptiveCardExtension';
  private _deferredPropertyPane: CompanyAnnouncementsCardPropertyPane | undefined;
  private _announcementsService = new CompanyAnnouncementsService();

  public async onInit(): Promise<void> {
    this.state = { announcements: [], currentIndex: 0, isLoading: true };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(DETAIL_VIEW_REGISTRY_ID, () => new AnnouncementDetailView());

    try {
      const siteUrl = this.properties.listSiteUrl || this.context.pageContext.web.absoluteUrl;
      await this._announcementsService.Init(siteUrl, this.context.pageContext);
      const announcements = await this._announcementsService.getAnnouncements();
      this.setState({ announcements, isLoading: false });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (onInit) - ${err}`);
      this.setState({ isLoading: false });
    }

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'CompanyAnnouncementsCard-property-pane'*/
      './CompanyAnnouncementsCardPropertyPane'
    ).then((component) => {
      this._deferredPropertyPane = new component.CompanyAnnouncementsCardPropertyPane();
    });
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
