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

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import PermissionRequest from './components/permissionRequest/PemissionRequest';
import GraphInfoComponent from './components/graphInfo/GraphInfoComponent';

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
          //if (e.toString().indexOf("AADSTS65001") != -1) {
            this.needConsent = true;
            this.render();
          //}
        })
    }).then(super.onInit);
  }

  public render(): void {
    if (this.needConsent == undefined) {
      this.domElement.innerHTML =`<div></div>`;
      return;
    }

    if (this.needConsent) {
      ReactDOM.render(
        React.createElement(PermissionRequest, { tenantId: this.context.pageContext.aadInfo.tenantId.toString() }),
        this.domElement
      );
      return;
    }

    this.context.msGraphClientFactory
      .getClient()
      .then((client: MSGraphClient) => {
        client
          .api("/sites/root")
          .version("v1.0")
          .get((err, res) => {
            if (err) {
              console.error(err);
              return;
            }

            ReactDOM.render(
              React.createElement(GraphInfoComponent, {
                displayName: res.displayName,
                description: res.description,
                createdDateTime: res.createdDateTime,
                lastModifiedDateTime: res.lastModifiedDateTime
              }),
              this.domElement
            );
            console.log(res);
            return;
          });
      });
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
