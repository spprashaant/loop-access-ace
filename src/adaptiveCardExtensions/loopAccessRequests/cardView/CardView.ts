import {
  BaseComponentsCardView,
  ComponentsCardViewParameters,
  BasicCardView,
  IExternalLinkCardAction,
  IQuickViewCardAction
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'LoopAccessRequestsAdaptiveCardExtensionStrings';
import {
  ILoopAccessRequestsAdaptiveCardExtensionProps,
  ILoopAccessRequestsAdaptiveCardExtensionState,
  QUICK_VIEW_REGISTRY_ID
} from '../LoopAccessRequestsAdaptiveCardExtension';

export class CardView extends BaseComponentsCardView<
  ILoopAccessRequestsAdaptiveCardExtensionProps,
  ILoopAccessRequestsAdaptiveCardExtensionState,
  ComponentsCardViewParameters
> {
  public get cardViewParameters(): ComponentsCardViewParameters {
    return BasicCardView({
      cardBar: {
        componentName: 'cardBar',
        title: this.properties.title
      },
      header: {
        componentName: 'text',
        text: this.state.loading ? 'Loading…' : `${this.state.requests.length} pending ${this.state.requests.length === 1 ? 'request' : 'requests'}`
      },
      footer: {
        componentName: 'cardButton',
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    });
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'QuickView',
      parameters: { view: QUICK_VIEW_REGISTRY_ID }
    };
  }
}
