import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView, ISubmitActionArguments } from '@microsoft/sp-adaptive-card-extension-base';
import template from './template/QuickViewTemplate.json';
import { IEmployeeSpotlight } from '../../../common/models/models';
import {
  IEmployeeSpotlightCardProps,
  IEmployeeSpotlightCardState
} from '../EmployeeSpotlightCardAdaptiveCardExtension';

export interface ISpotlightQuickViewData {
  spotlight: IEmployeeSpotlight;
  hasPageUrl: boolean;
  hasPrev: boolean;
  hasNext: boolean;
  currentPosition: string;
  showNavigation: boolean;
}

const EMPTY_SPOTLIGHT: IEmployeeSpotlight = {
  id: '',
  employeeName: '',
  spotlightType: '',
  description: '',
  employeeImageUrl: '',
  spotlightDate: '',
  spotlightDateFormatted: '',
  fullArticle: '',
  pageUrl: ''
};

export class QuickView extends BaseAdaptiveCardQuickView<
  IEmployeeSpotlightCardProps,
  IEmployeeSpotlightCardState,
  ISpotlightQuickViewData
> {
  private LOG_SOURCE = '🌟 SpotlightQuickView';

  public get data(): ISpotlightQuickViewData {
    const { spotlights, currentIndex } = this.state;
    const spotlight = spotlights[currentIndex] ?? EMPTY_SPOTLIGHT;
    return {
      spotlight,
      hasPageUrl: !!spotlight.pageUrl && spotlight.pageUrl !== '#',
      hasPrev: currentIndex > 0,
      hasNext: currentIndex < spotlights.length - 1,
      currentPosition: `${currentIndex + 1} of ${spotlights.length}`,
      showNavigation: spotlights.length > 1
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }

  public async onAction(action: ISubmitActionArguments): Promise<void> {
    const { spotlights, currentIndex } = this.state;
    if (action.type === 'Submit') {
      if (action.data?.id === 'prev' && currentIndex > 0) {
        await this.setState({ currentIndex: currentIndex - 1 });
      } else if (action.data?.id === 'next' && currentIndex < spotlights.length - 1) {
        await this.setState({ currentIndex: currentIndex + 1 });
      }
    }
  }
}
