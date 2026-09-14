import { ISPFxAdaptiveCard, BaseAdaptiveCardQuickView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'LoopAccessRequestsAdaptiveCardExtensionStrings';
import template from './template/QuickViewTemplate.json';
import {
  ILoopAccessRequestsAdaptiveCardExtensionProps,
  ILoopAccessRequestsAdaptiveCardExtensionState,
  ILoopRequest
} from '../LoopAccessRequestsAdaptiveCardExtension';

export interface IQuickViewData {
  requests: ILoopRequest[]
}

export class QuickView extends BaseAdaptiveCardQuickView<
  ILoopAccessRequestsAdaptiveCardExtensionProps,
  ILoopAccessRequestsAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      requests: this.state.requests
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return template as ISPFxAdaptiveCard;
  }
}
