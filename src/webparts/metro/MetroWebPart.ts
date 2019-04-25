import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';
import Metro from '../../CSS/Metro.module.scss';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './MetroWebPart.module.scss';
import * as strings from 'MetroWebPartStrings';
import {Environment,EnvironmentType} from '@microsoft/sp-core-library';
import {SPHttpClient,SPHttpClientResponse   } from '@microsoft/sp-http';

export interface IMetroWebPartProps {
  description: string;
  list: string;
}

export interface ISPLists {
  value: ISPList[];
 }
 
 
 export interface ISPList {
  Size: string;
 }

export default class MetroWebPart extends BaseClientSideWebPart<IMetroWebPartProps> {


  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${this.properties.list}')/Items?$select=Size`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
   }



   private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
      <div class="${item.Size}">From String</div>
      <div class ="${Metro.Long}">From Div</div>
      `;
    });

    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }


  private _renderListAsync(): void {
    // Local environment
 if (Environment.type == EnvironmentType.SharePoint || 
              Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getListData()
        .then((response) => {
          this._renderList(response.value);
        });
    }
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${Metro.Metro}">
              <div id="spListContainer"/></div>
                </div>


      `;this._renderListAsync();
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
                PropertyPaneTextField('list', {
                  label: strings.DescriptionFieldLabel
                  
                }
                
              )

              ]
            }
          ]
        }
      ]
    };
  }
}
