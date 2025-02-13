import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { override } from '@microsoft/decorators';
import * as strings from 'RemoveShareExtApplicationCustomizerStrings';

const LOG_SOURCE: string = 'RemoveShareExtApplicationCustomizer';

export interface IRemoveShareExtApplicationCustomizerProperties { }

export default class RemoveShareExtApplicationCustomizer extends BaseApplicationCustomizer<IRemoveShareExtApplicationCustomizerProperties> {

  private pollingInterval: ReturnType<typeof setInterval> | null = null;
  private mutationObserver: MutationObserver | null = null;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.startPollingForShareButton();  // need to start polling immediately
    this.observeDomChanges(); // observes DOM changes, using this for Mutation Observer

    window.addEventListener('hashchange', () => { this.restartPolling(); });
    window.addEventListener('popstate', () => { this.restartPolling(); });
    window.addEventListener('beforeunload', () => { this.stopPolling(); this.disconnectObserver(); }); // will disconnect the observer 
    return Promise.resolve();
  }

  private startPollingForShareButton(): void {
    this.pollingInterval = setInterval(() => {
      this.checkForShareButtonAndRemove(document.body); // check the entire page periodically because share has a tendency to keep coming up
    }, 250); // setting the interval to 250ms for efficient removal
  }

  private stopPolling(): void {
    if (this.pollingInterval) {
      clearInterval(this.pollingInterval);
      this.pollingInterval = null;
      Log.info(LOG_SOURCE, 'Polling has been stopped');
    }
  }

  private restartPolling(): void {
    this.stopPolling();
    this.startPollingForShareButton();
  }

  private observeDomChanges(): void {
    this.mutationObserver = new MutationObserver((mutationsList) => {
      for (const mutation of mutationsList) {
        if (mutation.type === 'childList') {
          mutation.addedNodes.forEach(node => {
            if (node instanceof Element) {
              this.checkForShareButtonAndRemove(node);
            }
          });
        }
      }
    });

    this.mutationObserver.observe(document.body, { childList: true, subtree: true });
  }

  private disconnectObserver(): void {
    if (this.mutationObserver) {
      this.mutationObserver.disconnect();
      this.mutationObserver = null;
      Log.info(LOG_SOURCE, 'MutationObserver disconnected');
    }
  }

  private checkForShareButtonAndRemove(element: Element): void {
    const shareButtons = element.querySelectorAll('[data-automation-id="shareButton"]');
    shareButtons.forEach(button => {
      button.remove();
      Log.info(LOG_SOURCE, 'Share button is removed');
    });
  }
}