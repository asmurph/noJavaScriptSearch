import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NoJavascriptWebpartWebPartWebPart.module.scss';
import * as strings from 'NoJavascriptWebpartWebPartWebPartStrings';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
export interface INoJavascriptWebpartWebPartWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Id: string;
  Title: string;
  Urgent: string;
}
export default class NoJavascriptWebpartWebPartWebPart extends BaseClientSideWebPart<INoJavascriptWebpartWebPartWebPartProps> {
  private _getListData(): Promise<ISPLists> {
    let queryString: string = '';
    let queryStringforPR: string = '';
    let searchboxVal: string=(this.domElement.querySelector('#searchbox') as  HTMLInputElement).value;
    if(searchboxVal!=""){
      queryString="$filter=((substringof('"+searchboxVal+"',PR_Number)))";
      queryStringforPR= searchboxVal;
      console.log("qurery string value is " + queryString);
    }
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('site pages')/Items?${queryString}`, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
         // debugger;
          return response.json();
        });
  }
  private _renderListAsync(): void {
    this._getListData()
    .then((response) => {
      console.log("respone is " + this._renderList(response.value));
      this._renderList(response.value);
    });
  }
  
  private _renderList(items: ISPList[]): void {
    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';
    html += `<th>PR_Number</th><th>Description</th><th>Request_Date</th>`;
    items.forEach((item: ISPList) => {
      html += `
          <tr>
          <td>${item.Id}</td>
          <td>${item.Title}</td>
          <td>${item.Urgent}</td>
          </tr>
          `;
    });
    html += `</table>`;
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }
  private _setSearchBtnEventHandlers(): void {
    this.domElement.querySelector('#searchBtn').addEventListener('click', () => {
        this._renderListAsync();
    });
 }
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.noJavascriptWebpartWebPart }">
        <div class="${ styles.container }">
        <div class= "${styles.row}">
        <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <p style="text-align: center">Demo : Retrieve Data from SharePoint List</p>
        </div>
      </div>
        <br>
        <br>
        <input id="searchbox" type="textbox"/><input id="searchBtn" type="button" value="Search"/>
        <br>
        <br>
        <div id="spListContainer" />
    </div>
  </div>`;
  this._renderListAsync();
  this._setSearchBtnEventHandlers();
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
