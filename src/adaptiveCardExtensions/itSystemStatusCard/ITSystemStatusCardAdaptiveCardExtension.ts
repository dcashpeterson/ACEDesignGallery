import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { ITSystemStatusCardPropertyPane } from './ITSystemStatusCardPropertyPane';
import { IServiceHealth, IServiceHealthIssue } from '../../common/models/models';
import { ServiceHealthService } from '../../common/services/ServiceHealthService';
import { IssueView } from './quickView/IssueView';

export interface IITSystemStatusCardAdaptiveCardExtensionProps {
  title: string;
}

export interface IITSystemStatusCardAdaptiveCardExtensionState {
  services: IServiceHealth[];
  issues: IServiceHealthIssue[];
  selectedService: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'ITSystemStatusCard_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'ITSystemStatusCard_QUICK_VIEW';
export const ISSUE_VIEW_REGISTRY_ID: string = 'ITSystemStatusCard_ISSUE_VIEW';
export const NEW_VIEW_REGISTRY_ID: string = 'ITSystemStatusCard_NEW_VIEW';

export default class ITSystemStatusCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IITSystemStatusCardAdaptiveCardExtensionProps,
  IITSystemStatusCardAdaptiveCardExtensionState
> {
  private LOG_SOURCE = '🔴 ITSystemStatusCardAdaptiveCardExtension';

  private _deferredPropertyPane: ITSystemStatusCardPropertyPane | undefined;
  private _healthService = new ServiceHealthService(this.context);

  public async onInit(): Promise<void> {
    this.state = { services: [], issues: [], selectedService: '' };

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
    this.quickViewNavigator.register(ISSUE_VIEW_REGISTRY_ID, () => new IssueView());

    try {
      void await this._healthService.Init(this.context.pageContext.web.absoluteUrl, this.context.pageContext, this.context.aadTokenProviderFactory);
      const services = await this._healthService.getServices();
      const issues = await this._healthService.getIssuesByService();
      this.setState({ services, issues });
    } catch (err) {
      console.error(`${this.LOG_SOURCE} (render) - ${err}`);
    }

    return Promise.resolve();
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'ITSystemStatusCard-property-pane'*/
      './ITSystemStatusCardPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.ITSystemStatusCardPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration() ?? super.getPropertyPaneConfiguration();
  }
}
