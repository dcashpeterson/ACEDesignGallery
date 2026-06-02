import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/AnnouncementDetailViewTemplate.json';
import { ICompanyAnnouncement } from '../../../common/models/models';
import {
  ICompanyAnnouncementsCardProps,
  ICompanyAnnouncementsCardState
} from '../CompanyAnnouncementsCardAdaptiveCardExtension';

export interface IAnnouncementDetailViewData {
  announcement: ICompanyAnnouncement;
  hasPrev: boolean;
  hasNext: boolean;
  currentPosition: string;
  showNavigation: boolean;
}

const EMPTY_ANNOUNCEMENT: ICompanyAnnouncement = {
  id: '',
  title: '',
  body: '',
  publishedDate: '',
  publishedDateFormatted: '',
  category: ''
};

export class AnnouncementDetailView extends BaseAdaptiveCardQuickView<
  ICompanyAnnouncementsCardProps,
  ICompanyAnnouncementsCardState,
  IAnnouncementDetailViewData
> {
  private LOG_SOURCE = '📢 AnnouncementDetailView';

  public get data(): IAnnouncementDetailViewData {
    const { announcements, currentIndex } = this.state;
    const announcement = announcements[currentIndex] ?? EMPTY_ANNOUNCEMENT;
    return {
      announcement,
      hasPrev: currentIndex > 0,
      hasNext: currentIndex < announcements.length - 1,
      currentPosition: `${currentIndex + 1} of ${announcements.length}`,
      showNavigation: announcements.length > 1
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    const { announcements, currentIndex } = this.state;
    if (action.type === 'Submit') {
      if (action.data?.id === 'prev' && currentIndex > 0) {
        await this.setState({ currentIndex: currentIndex - 1 });
      } else if (action.data?.id === 'next' && currentIndex < announcements.length - 1) {
        await this.setState({ currentIndex: currentIndex + 1 });
      }
    }
  }
}
