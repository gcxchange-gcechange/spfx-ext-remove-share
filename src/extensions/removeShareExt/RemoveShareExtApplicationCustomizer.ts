import { Log } from '@microsoft/sp-core-library';
import { BaseApplicationCustomizer } from '@microsoft/sp-application-base';
import { override } from '@microsoft/decorators';
import * as strings from 'RemoveShareExtApplicationCustomizerStrings';

const LOG_SOURCE: string = 'RemoveShareExtApplicationCustomizer';

export interface IRemoveShareExtApplicationCustomizerProperties { }

export default class RemoveShareExtApplicationCustomizer extends BaseApplicationCustomizer<IRemoveShareExtApplicationCustomizerProperties> {

  private pollingInterval: ReturnType<typeof setInterval> | null = null;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.startPollingForShareButton();  // need to start polling immediately

    // need to watch the pages for dynamically loaded contents 
    document.addEventListener('DOMNodeInserted', (event: Event) => {
      if (event.target instanceof Element) {
        this.checkForShareButtonAndRemove(event.target); // check for the inserted element
      }
    });

    window.addEventListener('hashchange', () => { this.restartPolling(); });
    window.addEventListener('popstate', () => { this.restartPolling(); });
    window.addEventListener('beforeunload', () => { this.stopPolling(); });

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


  private checkForShareButtonAndRemove(element: Element): void {
    const shareButtons = element.querySelectorAll('[data-automation-id="shareButton"]');
    shareButtons.forEach(button => {
      button.remove();
      Log.info(LOG_SOURCE, 'Share button is removed');
    });
  }
}