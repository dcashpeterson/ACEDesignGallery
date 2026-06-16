import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseWebQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import { ISalesBarChartCardProps, ISalesBarChartCardState } from '../SalesBarChartCardAdaptiveCardExtension';
import { SalesBarChartComponent } from '../components/SalesBarChartComponent';

export class QuickView extends BaseWebQuickView<
  ISalesBarChartCardProps,
  ISalesBarChartCardState
> {
  public render(): void {
    if (this.domElement) {
      const element = React.createElement(SalesBarChartComponent, {
        series: this.state.series,
        totalSales: this.state.totalSales,
        isLoading: this.state.isLoading
      });
      ReactDOM.render(element, this.domElement);
    }
  }

  public onBeforeUnload(): boolean {
    if (this.domElement) {
      ReactDOM.unmountComponentAtNode(this.domElement);
    }
    return false;
  }
}