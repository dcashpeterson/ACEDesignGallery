import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  ImageCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import {
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  FACILITIES_MY_REQUESTS_VIEW_ID,
  FACILITIES_MAINTENANCE_QUEUE_VIEW_ID,
  FACILITIES_MANAGER_QUEUE_VIEW_ID
} from '../FacilitiesMaintenanceCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IFacilitiesMaintenanceCardAdaptiveCardExtensionProps,
  IFacilitiesMaintenanceCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { userRole, myRequests, maintenanceQueue, managerQueue, isLoading } = this.state;

    let headerText: string;
    let primaryViewId: string;
    let buttonLabel: string;

    if (isLoading) {
      headerText = 'Loading…';
      primaryViewId = FACILITIES_MY_REQUESTS_VIEW_ID;
      buttonLabel = 'My Requests';
    } else if (userRole === 'manager') {
      const count = managerQueue.length;
      headerText = `${count} item${count !== 1 ? 's' : ''} pending review`;
      primaryViewId = FACILITIES_MANAGER_QUEUE_VIEW_ID;
      buttonLabel = 'Manager Queue';
    } else if (userRole === 'maintenance') {
      const count = maintenanceQueue.length;
      headerText = `${count} item${count !== 1 ? 's' : ''} in your queue`;
      primaryViewId = FACILITIES_MAINTENANCE_QUEUE_VIEW_ID;
      buttonLabel = 'My Queue';
    } else {
      const count = myRequests.length;
      headerText = `${count} request${count !== 1 ? 's' : ''} submitted`;
      primaryViewId = FACILITIES_MY_REQUESTS_VIEW_ID;
      buttonLabel = 'My Requests';
    }

    return ImageCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: headerText
      },
      image: {
        url: this.properties.imageUrl || '',
        altText: 'Facilities Maintenance'
      },
      footer: {
        componentName: 'cardButton' as const,
        title: buttonLabel,
        style: 'positive',
        action: {
          type: 'QuickView' as const,
          parameters: { view: primaryViewId }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    const { userRole } = this.state;
    const view = userRole === 'manager'
      ? FACILITIES_MANAGER_QUEUE_VIEW_ID
      : userRole === 'maintenance'
        ? FACILITIES_MAINTENANCE_QUEUE_VIEW_ID
        : FACILITIES_MY_REQUESTS_VIEW_ID;

    return {
      type: 'QuickView',
      parameters: { view }
    };
  }
}
