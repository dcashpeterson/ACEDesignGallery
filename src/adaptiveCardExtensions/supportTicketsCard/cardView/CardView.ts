import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  PieChartCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'SupportTicketsCardAdaptiveCardExtensionStrings';
import {
  ISupportTicketsCardAdaptiveCardExtensionProps,
  ISupportTicketsCardAdaptiveCardExtensionState,
  CHART_VIEW_REGISTRY_ID
} from '../SupportTicketsCardAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ISupportTicketsCardAdaptiveCardExtensionProps,
  ISupportTicketsCardAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    const { counts, isLoading } = this.state;

    const headerText = isLoading
      ? 'Loading tickets…'
      : `${counts.open} open ticket${counts.open !== 1 ? 's' : ''}`;

    return PieChartCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      body: {
        componentName: 'dataVisualization',
        dataVisualizationKind: 'pie',
        isDonut: false,
        series: [{
          data: [
            { x: 'Open', y: counts.open, color: '#0078d4', showLabel:true },
            { x: 'In Progress', y: counts.inProgress, color: '#ffb900', showLabel:true },
            { x: 'Resolved', y: counts.resolved, color: '#107c10', showLabel:true },
            { x: 'Escalated', y: counts.escalated, color: '#d13438', showLabel:true }
          ]
        }]
      },
      footer: {
        componentName: 'cardButton',
        title: strings.QuickViewButton,
        style: 'positive',
        action: {
          type: 'QuickView',
          parameters: { view: CHART_VIEW_REGISTRY_ID }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: CHART_VIEW_REGISTRY_ID }
    };
  }
}
