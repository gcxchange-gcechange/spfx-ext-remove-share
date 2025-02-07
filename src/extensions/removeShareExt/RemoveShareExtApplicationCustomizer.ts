import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { override } from '@microsoft/decorators';

import * as strings from 'RemoveShareExtApplicationCustomizerStrings';

const LOG_SOURCE: string = 'RemoveShareExtApplicationCustomizer';

export interface IRemoveShareExtApplicationCustomizerProperties {
  // No properties needed for this example
}

export default class RemoveShareExtApplicationCustomizer
  extends BaseApplicationCustomizer<IRemoveShareExtApplicationCustomizerProperties> {

  private observer: MutationObserver | undefined = undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    window.addEventListener('DOMContentLoaded', () => {
      this.removeShareButton();
    });

    setTimeout(() => {
      this.removeShareButton();
    }, 1000);

    // *** Corrected Disconnection Logic - ONLY beforeunload ***
    window.addEventListener('beforeunload', () => {
      this.disconnectObserver();
    });

    return Promise.resolve();
  }

  private removeShareButton(): void {
    // Check if the document body exists before trying to query it.
    if (document && document.body) {
      const shareButton = document.querySelector('[data-automation-id="shareButton"]');

      if (shareButton) {
        shareButton.remove();
        Log.info(LOG_SOURCE, 'Share button removed.');

        this.observer = new MutationObserver((mutations) => {
          mutations.forEach((mutation) => {
            mutation.addedNodes.forEach((node) => {
              if (node instanceof Element && node.matches('[data-automation-id="shareButton"]')) {
                node.remove();
                Log.info(LOG_SOURCE, 'Share button removed (dynamic).');
              }
            });
          });
        });

        this.observer.observe(document.body, { childList: true, subtree: true });

      } else {
        Log.info(LOG_SOURCE, 'Share button not found. Possibly loaded later.');
      }
    } else {
      Log.warn(LOG_SOURCE, "Document body not available. Cannot proceed.");
    }
  }

  private disconnectObserver(): void {
    if (this.observer) {
      this.observer.disconnect();
      Log.info(LOG_SOURCE, 'MutationObserver disconnected.');
      this.observer = undefined;
    }
  }
}