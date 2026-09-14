import type { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { LoopAccessRequestsPropertyPane } from './LoopAccessRequestsPropertyPane';
import { SPHttpClient } from '@microsoft/sp-http';

export interface ILoopRequest {
  id: number;
  title: string;
  requester: string;
  expiresOn: string;
}

export interface ILoopAccessRequestsAdaptiveCardExtensionState {
  loading: boolean;
  requests: ILoopRequest[];
}

export interface ILoopAccessRequestsAdaptiveCardExtensionProps {
  title: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'LoopAccessRequests_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'LoopAccessRequests_QUICK_VIEW';

export default class LoopAccessRequestsAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  ILoopAccessRequestsAdaptiveCardExtensionProps,
  ILoopAccessRequestsAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: LoopAccessRequestsPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = { loading: true, requests: [] };
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
    this._loadRequests().catch(console.error);   // deliberately not awaited
    return Promise.resolve();
  }

  private async _loadRequests(): Promise<void> {
    const web = this.context.pageContext.web.absoluteUrl;
    const url = `${web}/_api/web/lists/getbytitle('LoopAccessRequests')/items` +
      `?$select=Id,Title,ExpiresOn,Requester/Title&$expand=Requester` +
      `&$filter=Status eq 'Pending'&$orderby=ExpiresOn asc`;
    const res = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    const json = await res.json();
    this.setState({
      loading: false,
      requests: json.value.map((i: any) => ({
        id: i.Id, title: i.Title, requester: i.Requester?.Title ?? '', expiresOn: i.ExpiresOn
      }))
    });
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'LoopAccessRequests-property-pane'*/
      './LoopAccessRequestsPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.LoopAccessRequestsPropertyPane();
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
