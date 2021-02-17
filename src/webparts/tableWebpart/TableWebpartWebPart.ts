import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import {  
  Environment,  
  EnvironmentType  
} from '@microsoft/sp-core-library';  
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TableWebpartWebPart.module.scss';
import * as strings from 'TableWebpartWebPartStrings';
import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';
export interface ITableWebpartWebPartProps {
  description: string;
}
export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {   
  ID: string;  
  SubName: string;   
}  
export default class TableWebpartWebPart extends BaseClientSideWebPart<ITableWebpartWebPartProps> {
  private _getListData(): Promise<ISPLists> {  
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/GetByTitle('Subjects')/Items`, SPHttpClient.configurations.v1)  
      .then((response: SPHttpClientResponse) => {
        return response.json();
    });  
  }

  private _renderListAsync(): void {  
      
    if (Environment.type === EnvironmentType.Local) {   
    }  
    else {  
        this._getListData()  
      .then((response) => {  
        this._renderList(response.value);  
      });  
    }    
  }

  private _renderList(items: ISPList[]): void {  
    let html: string = '<table class="TFtable" border=1 width=100% style="border-collapse: collapse;">';  
    html += `<th>ID</th><th>Name</th>`;  
    items.forEach((item: ISPList) => {  
      html += `  
           <tr>  
          <td>${item.ID}</td>  
          <td>${item.SubName}</td>  
          </tr>  
          `;  
    });  
    html += `</table>`;  
    const listContainer: Element = this.domElement.querySelector('#spListContainer');  
    listContainer.innerHTML = html;  
  }
  public render(): void {
    this.domElement.innerHTML = `  
    <div class="${styles.tableWebpart}">  
  <div class="${styles.container}">  
   <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
     <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
       <span class="ms-font-xl ms-fontColor-white" style="font-size:28px">Welcome to SharePoint Framework Development</span>  
         
       <p class="ms-font-l ms-fontColor-white" style="text-align: center">Demo : Retrieve Subjects Data from SharePoint List</p>  
     </div>  
   </div>  
   <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
   <div style="background-color:Black;color:white;text-align: center;font-weight: bold;font-size:18px;">Subject Details</div>  
   <br>  
  <div id="spListContainer" />  
   </div>  
  </div>  
  </div>`; 
  this._renderListAsync();  
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
