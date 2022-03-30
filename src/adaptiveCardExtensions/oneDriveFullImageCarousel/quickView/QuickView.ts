import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'OneDriveFullImageCarouselAdaptiveCardExtensionStrings';
import { IOneDriveFullImageCarouselAdaptiveCardExtensionProps, IOneDriveFullImageCarouselAdaptiveCardExtensionState } from '../OneDriveFullImageCarouselAdaptiveCardExtension';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IQuickViewData {
  detailsLabel: string;
  fileNameLabel: string;
  sizeLabel: string;
  modifiedLabel: string;
  currentItem: MicrosoftGraph.DriveItem;
}

export class QuickView extends BaseAdaptiveCardView<
  IOneDriveFullImageCarouselAdaptiveCardExtensionProps,
  IOneDriveFullImageCarouselAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    return {
      detailsLabel: "strings.DetailsLabel",
      fileNameLabel: "strings.FileNameLabel",
      sizeLabel: "strings.SizeLabel",
      modifiedLabel: "strings.ModifiedLabel",
      currentItem: (this.state.targetFolder && this.state.targetFolder.children) ? this.state.targetFolder.children[this.state.itemIndex] : undefined
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}