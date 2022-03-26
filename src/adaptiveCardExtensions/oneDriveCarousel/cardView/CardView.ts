import {
  BaseImageCardView,
  IImageCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'OneDriveCarouselAdaptiveCardExtensionStrings';
import { IOneDriveCarouselAdaptiveCardExtensionProps, IOneDriveCarouselAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../OneDriveCarouselAdaptiveCardExtension';

export class CardView extends BaseImageCardView<IOneDriveCarouselAdaptiveCardExtensionProps, IOneDriveCarouselAdaptiveCardExtensionState> {
  /**
   * Buttons will not be visible if card size is 'Medium' with Image Card View.
   * It will support up to two buttons for 'Large' card size.
   */
  public get cardButtons():[ICardButton] | [ICardButton, ICardButton] | undefined {
    var buttons = [];
    
    if(!this.state.error) {
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

  public get data(): IImageCardParameters {
    return {
      primaryText: (this.state.error) ? strings.ErrorMessage : ((this.properties.description) ? this.properties.description : strings.PrimaryText),
      imageUrl: (this.state.error) ? require('../assets/Error.png') : ((this.state.targetFolder && this.state.targetFolder.children) ? this.state.targetFolder.children[this.state.itemIndex].webUrl : require('../assets/MicrosoftLogo.png'))
    };
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
