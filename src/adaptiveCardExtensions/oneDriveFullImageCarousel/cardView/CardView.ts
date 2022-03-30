import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton,
  ISPFxAdaptiveCard
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'OneDriveFullImageCarouselAdaptiveCardExtensionStrings';
import { IOneDriveFullImageCarouselAdaptiveCardExtensionProps, IOneDriveFullImageCarouselAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../OneDriveFullImageCarouselAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IOneDriveFullImageCarouselAdaptiveCardExtensionProps, IOneDriveFullImageCarouselAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    var buttons = [];
    
    if(!this.state.error && this.properties.fullBleed == false) {
      buttons = [
        {
          title: strings.QuickViewButton,
          action: {
            type: 'QuickView',
            parameters: {
              view: QUICK_VIEW_REGISTRY_ID
            }
          }
        }
      ];
    }

    return <[ICardButton] | [ICardButton, ICardButton] | undefined>buttons;
  }

  public get data(): IBasicCardParameters {
    return {
      primaryText: this.getImageUrl()
    };
  }

  private getImageUrl(): string {
    if (this.state.error) {
      return require('../assets/Error.png');
    }
    
    let imageUrl: string = require('../assets/MicrosoftLogo.png');
    
    if (this.state.targetFolder && this.state.targetFolder.children) {
      imageUrl = this.state.targetFolder.children[this.state.itemIndex].webUrl;
    }

    return imageUrl;
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/CardViewTemplate.json');
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: (this.state.targetFolder) ? this.state.targetFolder.webUrl : "https://onedrive.com/"
      }
    };
  }
}
