import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
//import { findIndex } from '@microsoft/sp-lodash-subset';

import styles from './MailboxManPermListNfWebPart.module.scss';
import * as strings from 'MailboxManPermListNfWebPartStrings';

import {
  AadHttpClient,
  HttpClientResponse,
} from '@microsoft/sp-http';

import { allComponents, provideFluentDesignSystem } from '@fluentui/web-components';
provideFluentDesignSystem().register(allComponents);

export interface IMailboxManPermListNfWebPartProps {
  description: string;
}

interface MailboxPermissionRequest {
  Identity: string;
  User: string;
  AccessRights: string;
  IsInherited: string;
  Deny: string;
  InheritanceType: string;
}

export default class MailboxManPermListNfWebPart extends BaseClientSideWebPart<IMailboxManPermListNfWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <section>
    <center><h3>Mailbox Access Lookup</h3></center>
    <div id="mailboxPermissionLookupForm" class="${styles.mailboxSearchForm}">
          <fluent-text-field id="mailboxPermissionSearch" label="Search" placeholder="testmailbox@${this._getUserDomain()}"></fluent-text-field>
          <fluent-button data-action="startSearch" id="mailboxPermissionSearchButton" appearance="accent">Search</fluent-button>
    </div>
    <div id="mailboxPermissionGridContainer" class="${styles.mailboxPermissionContainer}"></div>
    </section>`;

    this._setButtonEventHandlers();
  }

  private _renderMailboxPermissions(items: MailboxPermissionRequest[]): void {
    let html: string = `<fluent-tree-view>`;
    items.forEach((item: MailboxPermissionRequest) => {
      html += `
      <fluent-tree-item expanded appearance="accent">
        <strong>${item.User}</strong>
        <fluent-tree-item>
          Access rights: <b>${item.AccessRights}</b>
        </fluent-tree-item>
        <fluent-tree-item>
          Is inherited: <b>${item.IsInherited}</b>
        </fluent-tree-item>
        <fluent-tree-item>
          Inheritance type: <b>${item.InheritanceType}</b>
        </fluent-tree-item>
      </fluent-tree-item>
        `;
    });

    html += `</fluent-tree-view>`;
    const listContainer: Element = this.domElement.querySelector('#mailboxPermissionGridContainer');
    listContainer.innerHTML = html;
    document.getElementById('loading').innerHTML = '';
  }

  private _getMailboxPermissions(mailAddress: string): Promise<MailboxPermissionRequest[]> {
    return this.context.aadHttpClientFactory.getClient('api://ca6525e8-f700-4f8b-95ab-053387edf950')
      .then((client: AadHttpClient): Promise<HttpClientResponse> => {
        return client.get(`https://chjb-mailmanage.azurewebsites.net/api/RequestMailboxAccess?requestAddress=${mailAddress}`, AadHttpClient.configurations.v1);
      })
      .then((response: HttpClientResponse): Promise<MailboxPermissionRequest[]> => {
        return response.json();
      })
      .then((response: MailboxPermissionRequest[]): MailboxPermissionRequest[] => {
        return response;
      });
  }

  private _setButtonEventHandlers(): void {
    const searchButton: Element = this.domElement.querySelector('#mailboxPermissionSearchButton');
    searchButton.addEventListener('click', () => {
      const searchBox: HTMLInputElement = this.domElement.querySelector('#mailboxPermissionSearch');
      this._getMailboxPermissions(searchBox.value).then((items: MailboxPermissionRequest[]) => {
        this._renderMailboxPermissions(items);
      });
    });
  }

  private _getUserDomain(): string {
    let userDomain: string = '';
    if (this.context.pageContext.user.email) {
      userDomain = this.context.pageContext.user.email.split('@')[1];
    }
    return userDomain;
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
