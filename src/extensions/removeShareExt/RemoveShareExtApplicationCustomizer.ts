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

    this.startPollingForShareButton();
    this.observeDomChanges();

    window.addEventListener('hashchange', () => { this.restartPolling(); });
    window.addEventListener('popstate', () => { this.restartPolling(); });
    window.addEventListener('beforeunload', () => { this.stopPolling(); this.disconnectObserver(); });

    return Promise.resolve();
  }

  private startPollingForShareButton(): void {
    this.pollingInterval = setInterval(() => {
      this.checkForShareButtonAndHide(document.body);
    }, 250);
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
              this.checkForShareButtonAndHide(node);
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

  private checkForShareButtonAndHide(element: Element): void {
    const shareButtons = element.querySelectorAll('[data-automation-id="shareButton"]');
    shareButtons.forEach(button => {
      if (button instanceof HTMLElement) {
        button.style.display = 'none';
        Log.info(LOG_SOURCE, 'Share button is hidden');
      } else {
        Log.warn(LOG_SOURCE, 'Element is not an HTMLElement. Cannot hide.');
      }
    });
  }
}