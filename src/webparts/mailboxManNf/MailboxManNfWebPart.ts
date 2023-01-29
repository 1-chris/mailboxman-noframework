import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { findIndex } from '@microsoft/sp-lodash-subset';

import styles from './MailboxManNfWebPart.module.scss';
import * as strings from 'MailboxManNfWebPartStrings';

import { allComponents, provideFluentDesignSystem } from '@fluentui/web-components';
provideFluentDesignSystem().register(
  allComponents
);

import {
  AadHttpClient,
  HttpClientResponse,
} from '@microsoft/sp-http';

export interface IMailboxManNfWebPartProps {
  description: string;
}

export interface MailboxPermission {
  DisplayName: string;
  PrimaryEmailAddress: string;
  PermissionVia: string;
  AccessLevel: string;
}

export interface CalendarPermission {
  DisplayName: string;
  PrimaryEmailAddress: string;
  PermissionVia: string;
  AccessLevel: string;
}

export default class MailboxManNfWebPart extends BaseClientSideWebPart<IMailboxManNfWebPartProps> {

  private _mailboxPermissions: MailboxPermission[] = [];
  private _calendarPermissions: CalendarPermission[] = [];

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.mailboxManNf} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div>
        <div id="loading">
          <center>
            <fluent-skeleton style="height: 400px; width: 700px; background-color: #ffffff" shape="rect" shimmer="true">
              <svg
                style="position: absolute; left: 0; top: 0;"
                id="pattern"
                width="100%"
                height="100%">
              <defs>
                  <mask id="mask" x="0" y="0" width="100%" height="100%">
                      <rect x="0" y="0" width="100%" height="100%" fill="#ffffff" />
                      <rect x="0" y="0" width="100%" height="45%" rx="4" />
                      <rect x="25" y="55%" width="90%" height="15px" rx="4" />
                      <rect x="25" y="65%" width="70%" height="15px" rx="4" />
                      <rect x="25" y="80%" width="90px" height="30px" rx="4" />
                  </mask>
              </defs>
              <rect
                  x="0"
                  y="0"
                  width="100%"
                  height="100%"
                  mask="url(#mask)"
                  fill="#ffffff"
              />
            </svg>
          </fluent-skeleton>
          </center>
        </div>

      </div>
      <fluent-tabs id="permissionTabs" activeid="mailboxaccess">
        <fluent-tab id="mailboxaccess" style="margin: 0 auto">My shared mailbox access</fluent-tab>
        <fluent-tab id="calendaraccess" style="margin: 0 auto">My shared calendar access</fluent-tab>
        <fluent-tab-panel id="mailboxaccess">
          <div id="mailboxPermissionContainer"></div>
        </fluent-tab-panel>
        <fluent-tab-panel id="calendaraccess">
          <div id="calendarPermissionContainer"></div>
        </fluent-tab-panel>
      </fluent-tabs>
    </section>`;

    // set button event handlers after render
    this._renderMailboxPermissionsAsync();
    this._renderCalendarPermissionsAsync();
    this._setButtonEventHandlers();
  }


  private _getMailboxPermissions(): Promise<MailboxPermission[]> {
    return this.context.aadHttpClientFactory.getClient('api://ca6525e8-f700-4f8b-95ab-053387edf950')
      .then((client: AadHttpClient): Promise<HttpClientResponse> => {
        return client.get(`https://chjb-mailmanage.azurewebsites.net/api/GetMyMailboxAccess`, AadHttpClient.configurations.v1);
      })
      .then((response: HttpClientResponse): Promise<MailboxPermission[]> => {
        return response.json();
      })
      .then((response: MailboxPermission[]): MailboxPermission[] => {
        return response;
      });
  }

  // code to set button event handlers
  private _setButtonEventHandlers(): void {
    console.log(this.domElement)
    const buttons: NodeListOf<Element> = this.domElement.querySelectorAll('[data-action="remove"]');
    console.log(buttons)
    buttons.forEach((button: Element) => {
      console.log(button);
      button.addEventListener('click', (event: Event) => {
        console.log(event);
        const id: string = (event.target as Element).getAttribute('data-id');
        this._removeMailboxPermission(id);
      });
    });
  }
  
  // code to remove the mailbox element from button onclick
  public _removeMailboxPermission(id: string): void {
    const index: number = parseInt(id.replace('listitem', ''), 10);
    const item: MailboxPermission = this._mailboxPermissions[index];
    const removalType: string = item.PermissionVia === 'Direct' ? 'user' : 'group';
    console.log(item);

    // show loading icon
    document.getElementById('listitemdeleting' + index.toString()).innerHTML = '<fluent-progress-ring></fluent-progress-ring>';

    // use aadhttpclient to remove the mailbox permission
    this.context.aadHttpClientFactory.getClient('api://ca6525e8-f700-4f8b-95ab-053387edf950')
      .then((client: AadHttpClient): Promise<any> => {
        return client.get(`https://chjb-mailmanage.azurewebsites.net/api/RemoveMyMailboxAccess?mailboxToRemove=${item.PrimaryEmailAddress}&removalType=${removalType}`, AadHttpClient.configurations.v1);
      })
      .then((response: HttpClientResponse): Promise<any> => {
        console.log(response.json)
        return response.json();
      })


    const listContainer: Element = this.domElement.querySelector('#mailboxPermissionContainer');
    const element: Element = this.domElement.querySelector('#' + id);
    listContainer.removeChild(element);
    console.log();
    document.getElementById('listitemdeleting' + index).innerHTML = '';
  }

  private _renderMailboxPermissions(items: MailboxPermission[]): void {
    let html: string = '';
    items.forEach((item: MailboxPermission) => {
      html += `
      <div id="listitem${findIndex(items, item)}">
        <ul class="${styles.list}">
          <li class="${styles.listItem}">
            <span class="ms-font-l">Name: <strong>${item.DisplayName}</strong></span><br>
            <span class="ms-font-l">Address: <strong>${item.PrimaryEmailAddress}</strong></span><br>Access level: <strong>${item.AccessLevel}</strong><br>Permission via: <strong>${item.PermissionVia}</strong><br>
            <fluent-button class="deleteButton" appearance="accent" data-action="remove" data-id="listitem${findIndex(items, item)}">Remove my access</fluent-button>
            <div class="" id="listitemdeleting${findIndex(items, item)}"></div>
          </li>
        </ul>
      </div>`;
    });

    const listContainer: Element = this.domElement.querySelector('#mailboxPermissionContainer');
    listContainer.innerHTML = html;
    document.getElementById('loading').innerHTML = '';
    // set button event handlers
    this._setButtonEventHandlers();
  }

  private _renderMailboxPermissionsAsync(): Promise<void> {
    this._getMailboxPermissions()
      .then((response) => {
        this._mailboxPermissions = response;
        this._renderMailboxPermissions(response);
      })
      .catch((error) => { console.log(error) });
    return Promise.resolve();
  }

  private _getCalendarPermissions(): Promise<CalendarPermission[]> {
    return this.context.aadHttpClientFactory.getClient('api://ca6525e8-f700-4f8b-95ab-053387edf950')
      .then((client: AadHttpClient): Promise<HttpClientResponse> => {
        return client.get(`https://chjb-mailmanage.azurewebsites.net/api/GetMyCalendarAccess`, AadHttpClient.configurations.v1);
      })
      .then((response: HttpClientResponse): Promise<CalendarPermission[]> => {
        return response.json();
      })
      .then((response: CalendarPermission[]): CalendarPermission[] => {
        return response;
      });
  }

  private _renderCalendarPermissions(items: CalendarPermission[]): void {
    let html: string = '';
    items.forEach((item: CalendarPermission) => {
      html += `
      <div id="listitem${findIndex(items, item)}">
        <ul class="${styles.list}">
          <li class="${styles.listItem}">
            <span class="ms-font-l">Name: <strong>${item.DisplayName}</strong></span><br>
            <span class="ms-font-l">Address: <strong>${item.PrimaryEmailAddress}</strong></span><br>Access level: <strong>${item.AccessLevel}</strong><br>Permission via: <strong>${item.PermissionVia}</strong><br>
            <fluent-button class="deleteButton" appearance="accent" data-action="remove" data-id="listitem${findIndex(items, item)}">Remove my access</fluent-button>
            <div class="" id="listitemdeleting${findIndex(items, item)}"></div>
          </li>
        </ul>
      </div>`;
    });

    const listContainer: Element = this.domElement.querySelector('#calendarPermissionContainer');
    listContainer.innerHTML = html;
    document.getElementById('loading').innerHTML = '';
    // set button event handlers
    this._setButtonEventHandlers();
  }

  private _renderCalendarPermissionsAsync(): Promise<void> {
    this._getCalendarPermissions()
      .then((response) => {
        this._calendarPermissions = response;
        this._renderCalendarPermissions(response);
      })
      .catch((error) => { console.log(error) });
    return Promise.resolve();
  }


  public _removeCalendarPermission(id: string): void {
    const index: number = parseInt(id.replace('listitem', ''), 10);
    const item: CalendarPermission = this._calendarPermissions[index];
    const removalType: string = item.PermissionVia === 'Direct' ? 'user' : 'group';
    console.log(item);

    // show loading icon
    document.getElementById('listitemdeleting' + index).innerHTML = '<fluent-progress-ring>Removing permissions...</fluent-progress-ring>';

    // use aadhttpclient to remove the mailbox permission
    this.context.aadHttpClientFactory.getClient('api://ca6525e8-f700-4f8b-95ab-053387edf950')
      .then((client: AadHttpClient): Promise<any> => {
        return client.get(`https://chjb-mailmanage.azurewebsites.net/api/RemoveMyCalendarAccess?mailboxToRemove=${item.PrimaryEmailAddress}&removalType=${removalType}`, AadHttpClient.configurations.v1);
      })
      .then((response: HttpClientResponse): Promise<any> => {
        console.log(response.json)
        return response.json();
      })


    const listContainer: Element = this.domElement.querySelector('#calendarPermissionContainer');
    const element: Element = this.domElement.querySelector('#' + id);
    listContainer.removeChild(element);
    console.log();
    document.getElementById('listitemdeleting' + index).innerHTML = '';
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
