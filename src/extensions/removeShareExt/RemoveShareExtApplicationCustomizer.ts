/**
 * Written By: Adi Makkar 
 * Objective: To remove the share button on all pages to meet
 * the criteria for Pro B 
 */

import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { override } from '@microsoft/decorators';
import * as strings from 'RemoveShareExtApplicationCustomizerStrings';

const LOG_SOURCE: string = 'RemoveShareExtApplicationCustomizer';

export interface IRemoveShareExtApplicationCustomizerProperties {
  // no properties required 
}

export default class RemoveShareExtApplicationCustomizer
  extends BaseApplicationCustomizer<IRemoveShareExtApplicationCustomizerProperties> {

  private observer: MutationObserver | undefined = undefined;
  private shareButton: Element | null = null;

@override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    // using polling to see if the button actually loaded
    this.waitForShareButton();

    window.addEventListener('beforeunload', () => {
      this.disconnectObserver();
    });

    return Promise.resolve();
  }

  private waitForShareButton(): void {
    this.shareButton = document.querySelector('[data-automation-id="shareButton"]');
    if (this.shareButton) {
      this.removeShareButton();
    } else {
      // I needed to add this because it is taking forever for the share to load sometimes
      setTimeout(() => this.waitForShareButton(), 250); // poll every 250ms, should rapidly detect if share is missing
    }
  }

  private removeShareButton(): void {
    if (this.shareButton) {
      this.shareButton.remove();
      Log.info(LOG_SOURCE, 'Share button is removed');

       // disconnecting observer to prevent intra and inter ext issues 
       if (this.observer) {
         this.disconnectObserver();
       }

      this.observer = new MutationObserver((mutations) => {
        mutations.forEach((mutation) => {
          mutation.addedNodes.forEach((node) => {
            if (node instanceof Element && node.matches('[data-automation-id="shareButton"]')) {
              node.remove();
                Log.info(LOG_SOURCE, 'Share button removed dynamically');
                  // if the button reappears and is removed, disconnect the observer; Will have to restart the polling if you need to handle future reappearances
                  this.disconnectObserver();
                  this.waitForShareButton(); // restarting polling for future reappearances 
              }
            });
          })
         });

      this.observer.observe(document.body, { childList: true, subtree: true });

    } 

    else {
      Log.info(LOG_SOURCE, 'Share button is not found');  
    }
   }

  private disconnectObserver(): void {
    if (this.observer) {
      this.observer.disconnect();
        Log.info(LOG_SOURCE, 'MutationObserver is disconnected');
      this.observer = undefined;
    }
  }
}