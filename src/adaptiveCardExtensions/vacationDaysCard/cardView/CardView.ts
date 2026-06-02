import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  PieChartCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'VacationDaysCardAdaptiveCardExtensionStrings';
import {
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState,
  REQUESTS_VIEW_REGISTRY_ID
} from '../VacationDaysCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  IVacationDaysCardAdaptiveCardExtensionProps,
  IVacationDaysCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { summary, isLoading } = this.state;

    const usedY = isLoading ? 0 : summary.used;
    const availableY = isLoading ? 1 : (summary.available > 0 ? summary.available : 0);

    const headerText = isLoading
      ? 'Loading…'
      : `${summary.used} of ${summary.available} used`;

    return PieChartCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: headerText
      },
      body: {
        componentName: 'dataVisualization',
        dataVisualizationKind: 'pie',
        isDonut: false,
        series: [{
          data: [
            { x: 'Used', y: usedY, color: '#0078d4', showLabel: true },
            { x: 'Available', y: availableY, color: '#e0e0e0', showLabel: true }
          ]
        }]
      },
      footer: {
        componentName: 'cardButton',
        title: isLoading ? 'Loading…' : strings.QuickViewButton,
        style: 'positive',
        action: {
          type: 'QuickView',
          parameters: { view: REQUESTS_VIEW_REGISTRY_ID }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: REQUESTS_VIEW_REGISTRY_ID }
    };
  }
}
