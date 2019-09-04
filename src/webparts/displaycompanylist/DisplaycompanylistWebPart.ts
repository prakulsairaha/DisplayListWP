import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DisplaycompanylistWebPart.module.scss';
import * as strings from 'DisplaycompanylistWebPartStrings';

// 1. Import 2 classes 
import {SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration} from "@microsoft/sp-http"

// 12. Import 2 classes for deploying
import {Environment, EnvironmentType} from "@microsoft/sp-core-library";

export interface IDisplaycompanylistWebPartProps {
  description: string;
}

// 21. create an interface to replace any
export interface ICompany{
  Title : string;
  Location : string;
  HeadCount : number;
}

export default class DisplaycompanylistWebPart extends BaseClientSideWebPart<IDisplaycompanylistWebPartProps> {
  // 5. Declare Variable to store data
  private data : ICompany[] = [ // 22. replaced ANY with array of Icompany interface
    //15. entering sample data for local
    { Title : "Sample 1", Location : "Austin", HeadCount : 50},
    { Title : "Sample 2", Location : "Dallas", HeadCount : 51},
    { Title : "Sample 3", Location : "Houston", HeadCount : 52}
  ];

  // 2. Write function to get Data
  private getCompanyData() : ICompany[] {
    // 3. REST URL for getting All items from list
    let listURL = this.context.pageContext.web.absoluteUrl + "/_api/Lists/GetByTitle('Company')/Items?$select=Title,Location,HeadCount"; // 19. Query updated with select to selective data
    console.log("Calling REST ENDPOINT : " + listURL);

    // 4. here is gets data using promise, this is a async call
    this.context.spHttpClient.get(listURL,SPHttpClient.configurations.v1)
      .then((res : SPHttpClientResponse) => {
        console.log("Rest call successful returning data ..");
        return res.json();
      }).then((d : ICompany[]) =>{
        console.log("Data received : " + JSON.stringify(d));
          this.data = d;
          // 9. Call renderlist method, here HTML and data combines
          this.renderList(this.data);
      }).catch((err) => { // 11. Error handling
        console.log("Error in REST API CALL : " + err);
      });
      return this.data;
  }

  // 6. Write function to render data in webpart by creatin HTML
  private renderList(items : ICompany[]) : void { // 23. Replaces ANY
  let html ="<div>";

  items.forEach((item : ICompany) => { // 24. Replaces ANY
    html += `
    <div>
      Company ID : ${item.Title} <br\>
      Location : ${item.Location} <br\>
      HaedCount : ${item.HeadCount}
    </div>`;
  });

  html += '</div>';

  // 7. HTML Tag for html
  this.domElement.querySelector("#companydiv").innerHTML = html;
  }

// 8. Change Render
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.displaycompanylist }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Company Data</span>
              <p class="${ styles.subTitle }">List of Local Companies</p>              
             <div id="companydiv">
              ... Loading ...
              </div>
            </div>
          </div>
        </div>
      </div>`;
// 10. Do the Rest call and render wheneve data is returned
// 13. check for environment added
if(Environment.type === EnvironmentType.SharePoint){
      this.getCompanyData();
    } 
    // 14. Read from local data, updated line 24, this will show sample data else it will show loading on workbench
    else if (Environment.type === EnvironmentType.Local) {
 this.renderList(this.data);
    }
  }

  // 16. to package this solution use gulp package-solution
  // 17. Deploy the sppkg file on APP Catalog
  // 18. Add the webpart to the page, it would show error because it reads resourses from local so do gulp serve on local so its running and then web part will display data
  // 20. Replacing ANY type with proper datatypes

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
