import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  PrimaryTextCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'CompanyAnnouncementsCardAdaptiveCardExtensionStrings';
import {
  ICompanyAnnouncementsCardProps,
  ICompanyAnnouncementsCardState,
  DETAIL_VIEW_REGISTRY_ID
} from '../CompanyAnnouncementsCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ICompanyAnnouncementsCardProps,
  ICompanyAnnouncementsCardState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { announcements, currentIndex, isLoading } = this.state;
    const current = announcements[currentIndex];

    if (isLoading || !current) {
      return PrimaryTextCardView({
        cardBar: { componentName: 'cardBar', title: this.properties.title },
        header: {
          componentName: 'text',
          text: isLoading ? strings.Loading : strings.NoAnnouncements
        },
        body: { componentName: 'text', text: '' },
        footer: undefined
      });
    }

    const bodyPreview = current.body.length > 100
      ? `${current.body.substring(0, 100)}…`
      : current.body;

    const counter = announcements.length > 1
      ? ` (${currentIndex + 1} of ${announcements.length})`
      : '';

    return PrimaryTextCardView({
      cardBar: { componentName: 'cardBar', title: this.properties.title },
      header: { componentName: 'text', text: current.title },
      body: { componentName: 'text', text: `${bodyPreview}${counter}` },
      footer: {
        componentName: 'cardButton',
        title: strings.ReadMore,
        action: {
          type: 'QuickView',
          parameters: { view: DETAIL_VIEW_REGISTRY_ID }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: DETAIL_VIEW_REGISTRY_ID }
    };
  }
}
