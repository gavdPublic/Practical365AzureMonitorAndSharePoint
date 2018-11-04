import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './Prac365AzureMonitorWpWebPart.module.scss';
import * as strings from 'Prac365AzureMonitorWpWebPartStrings';
import {AppInsights} from "applicationinsights-js";

export interface IPrac365AzureMonitorWpWebPartProps {
  description: string;
}

export default class Prac365AzureMonitorWpWebPart extends BaseClientSideWebPart<IPrac365AzureMonitorWpWebPartProps> {

  public render(): void {

    let appInsightsKey: string = "2561a768-e7e3-4de2-a398-2f181b0d9e41";
    AppInsights.downloadAndSetup({ instrumentationKey: appInsightsKey });
    AppInsights.trackPageView();

    this.domElement.innerHTML = `
      <div class="${ styles.prac365AzureMonitorWp }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected myJavaScript() {
    alert("hola gustavito");
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
