import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { MSGraphClient } from "@microsoft/sp-http";

import styles from './QiangTestWebPart.module.scss';
import * as strings from 'QiangTestWebPartStrings';

export interface IQiangTestWebPartProps {
  description: string;
}

export default class QiangTestWebPart extends BaseClientSideWebPart<IQiangTestWebPartProps> {

  private needConsent: boolean | undefined = undefined;

  public onInit(): Promise<void> {
    // This will be in the web part lifecycle management. Here is a demo of concept.

    return this.context.aadTokenProviderFactory.getTokenProvider().then((tokenProvider) => {
      tokenProvider.getToken('https://graph.microsoft.com', false)
        .then(token => {
          this.needConsent = false;
          this.render();
        })
        .catch(e => {
          if (e.toString().indexOf("AADSTS65001") != -1) {
            this.needConsent = true;
            this.render();
          }
        })
    }).then(super.onInit);
  }

  public render(): void {
    if (this.needConsent == undefined) {
      this.domElement.innerHTML =`<div></div>`;
      return;
    }

    if (this.needConsent) {
      this.domElement.innerHTML = `
      <div>You haven't grant permission for this low trust app, please click the button below to grant your consent.</div>
      <a target="_blank" href="https://login.microsoftonline.com/${this.context.pageContext.aadInfo.tenantId.toString()}/oauth2/v2.0/authorize?response_type=id_token%20token&scope=https://graph.microsoft.com/Sites.Read.All openid profile&client_id=${'c199d865-f514-4171-81c0-10cb0f1ce923'}">Consent in the AAD</a>
      `
      return;
    }

    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("sites/root")
          .version("v1.0")
          .get((err, res) => {
            if (err) {
              console.error(err);
              return;
            }

            console.log(res);
          });
      });

    this.domElement.innerHTML = `
      <div class="${styles.qiangTest}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to SharePoint!</span>
              <p class="${styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${styles.description}">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
