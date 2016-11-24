import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import styles from './SharepointListsSubscriptionsExplorer.module.scss';
import * as strings from 'sharepointListsSubscriptionsExplorerStrings';
import { ISharepointListsSubscriptionsExplorerWebPartProps } from './ISharepointListsSubscriptionsExplorerWebPartProps';

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export interface ISPListSubscriptions {
  value: ISPListSubscription[];
}

export interface ISPListSubscription {
  id: string;
  resource: string;
  notificationUrl: string;
  expirationDateTime: string;
  clientState: string;
}

export default class SharepointListsSubscriptionsExplorerWebPart extends BaseClientSideWebPart<ISharepointListsSubscriptionsExplorerWebPartProps> {
  private _intervalHandler : number = -1;

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.sharepointListsSubscriptionsExplorer}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg12">
              <span class="ms-font-xl ms-fontColor-white">SharePoint Lists WebHook Subscriptions Explorer</span>
              <p class="ms-font-l ms-fontColor-white">Select List:</p>
              <div id="lists-selector-container">
                <select id='lists'>
                </select>
              </div>

              <div id="subscriptions-list-container">
                <ul id='${styles.subscriptions}'>
                </ul>
              </div>

              <p class="ms-font-l ms-fontColor-white">Add new subscription to the current list:</p>
              <div id="${styles.subscriptionAddContainer}">
                <div>
                  WebHook URL:
                </div>
                <div>
                  <textarea id="subscription-url"></textarea>
                </div>
                <div>
                  <a href="javascript:void(0)" class="ms-Button" id="add-subscription">
                      <span class="ms-Button-label">Add</span>
                  </a>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>`;

    this._fetchLists();
    this._addEventListeners();
  }

  private _addEventListeners(): void {
    var lists = this.domElement.querySelector("#lists") as HTMLSelectElement;
    lists.addEventListener("change", (evnt) => {
      this._fetchSubscriptions(lists.value);
    });

    var addSubscription = this.domElement.querySelector("#add-subscription") as HTMLLinkElement;
    addSubscription.addEventListener("click", evnt => {
      var subsUrlInput = this.domElement.querySelector("#subscription-url") as HTMLInputElement;
      this._addSubscription(lists.value, subsUrlInput.value);
    });
  }

  private _fetchLists(): void {
    this._getLists()
      .then((response) => {
        this._renderList(response.value);
        var listSelectorContainer = this.domElement.querySelector("#lists") as HTMLSelectElement;
        this._fetchSubscriptions(listSelectorContainer.value);
      });
  }

  private _getLists(): Promise<ISPLists> {
    return this.context.httpClient
      .get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=Hidden eq false`)
      .then((response: Response) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    let html: string = "";
    items.forEach((item: ISPList) => {
      html += `<option value=${item.Id}>${item.Title}</option>`;
    });

    const listContainer: Element = this.domElement.querySelector('#lists');
    listContainer.innerHTML = html;
  }

  private _fetchSubscriptions(listId: string): void {
    this._getSubscriptions(listId)
      .then((response) => {
        this._renderSubscriptions(response.value);      
        
        window.clearInterval(this._intervalHandler);
        this._intervalHandler = window.setInterval(() => {
          this._fetchSubscriptions(listId);
        }, 5000);
      });
  }

  private _getSubscriptions(listId: string): Promise<ISPListSubscriptions> {
    return this.context.httpClient
      .get(this.context.pageContext.web.absoluteUrl + `/_api/Web/Lists(guid'${listId}')/subscriptions`)
      .then((response: Response) => {
        return response.json();
      });
  }

  private _renderSubscriptions(items: ISPListSubscription[]): void {
    let html: string = "<li>There are no web hooks attached to this list.</li>";
    if (items.length !== 0) {
      html = "";
      items.forEach((item: ISPListSubscription) => {
        html += `
          <li>
            <div>Subscribtion ID: ${item.id}</div>
            <div>Notification URL:</div>
            <textarea>${item.notificationUrl}</textarea>
            <div>
              <a href="javascript:void(0)" class="ms-Button remove-subscription" data-id="${item.id}">
                  <span class="ms-Button-label" data-id="${item.id}">Remove</span>
              </a>
            </div>
          </li>`;
      });
    }

    const listContainer: Element = this.domElement.querySelector("#" + styles.subscriptions);
    listContainer.innerHTML = html;

    let buttons = this.domElement.querySelectorAll(".remove-subscription");
    for (let i = 0; i < buttons.length; i++) {
      let element = buttons[i];
      element.addEventListener("click", (evnt) => {
        let targetElement = (event.target || event.srcElement) as HTMLElement;
        var lists = this.domElement.querySelector("#lists") as HTMLSelectElement;
        this._removeSubscription(lists.value, targetElement.getAttribute('data-id')).then((response: Response) => {
          this._fetchSubscriptions(lists.value);
        });
      });
    }
  }

  private _removeSubscription(listId: string, id: string): Promise<any> {
    return this.context.httpClient.fetch(this.context.pageContext.web.absoluteUrl + `/_api/Web/Lists(guid'${listId}')/subscriptions(guid'${id}')`, {
      method: "DELETE",
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json;odata=nometadata",
        "X-RequestDigest": document.getElementById("__REQUESTDIGEST").getAttribute('value')
      }
    }).then((response: Response) => {
      return null;
    });
  }

  private _addSubscription(listId, url): Promise<any> {
    return this.context.httpClient.fetch(this.context.pageContext.web.absoluteUrl + `/_api/Web/Lists(guid'${listId}')/subscriptions`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Accept": "application/json;odata=verbose",
        "X-RequestDigest": document.getElementById("__REQUESTDIGEST").getAttribute('value')
      },
      body: JSON.stringify({
          "resource": this.context.pageContext.web.absoluteUrl + `/_api/Web/Lists(guid'${listId}')`, 
          "notificationUrl": url, 
          "expirationDateTime": "2017-01-01T16:00:00+00:00",   
          "clientState": "A0A354EC-97D4-4D83-9DDB-144077ADB440" 
      })
    }).then((response: Response) => {
      return response.json();
    });
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

