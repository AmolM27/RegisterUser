import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './RegisterUserWpWebPart.module.scss';
import * as strings from 'RegisterUserWpWebPartStrings';

import { sp, Web, List, ItemAddResult } from "@pnp/sp";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as jQuery from 'jquery';
import 'jqueryui';

export interface IRegisterUserWpWebPartProps {
  description: string;
}

export default class RegisterUserWpWebPart extends BaseClientSideWebPart<IRegisterUserWpWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div id="registrationitems"></div>
      <div id="sample"></div>
      <button  id="registerUserBtn">Register</button>`

    //Show Dropdown registration options
    this.showDropdownOptions();
    //Get User Profile property for user
    this.getUserProfilePropertyForUser("AccountName");
    //Register events
    this.registerEvents();
  }

  //Add user details to SP List Registration
  private registerUser() {

    /*sp.profiles.myProperties.get().then(function(result) {
      var props = result.UserProfileProperties;
      var propValue = "";
      props.forEach(function(prop) {
      propValue += prop.Key + " - " + prop.Value + "<br/>";
      });
      
      document.getElementById("sample").innerHTML = propValue;
      }).catch(function(err) {
      console.log("Error: " + err);
      });*/

    const web: Web = new Web(this.context.pageContext.web.absoluteUrl);
    var curruser = web.currentUser.get().then(function (res) {
      return res.Title;
    })
    alert("curruser: " + this.getUserProfilePropertyForUser("AccountName"));
    web.lists.getByTitle("MyList").items.add({
      Title: this.getUserProfilePropertyForUser("AccountName")
    });

  }

  //Register Events
  private registerEvents() {
    let btn = document.getElementById("registerUserBtn");
    btn.addEventListener("click", (e: Event) => this.registerUser());
  }

  //Show dropdown registration options by querying registrationoptions list
  private showDropdownOptions() {
    var html = "<select>";
    this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl + '/_api/web/lists/getbytitle(\'registrationoptions\')/items',
      SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((items: any) => {
          items.value.forEach(item => {
            html += `
              <option value="${item.Title}">${item.Title}</option>    
            `
          });
          html += "</select>"
          //this.domElement.querySelector('#registrationitems').innerHTML = html;
          document.getElementById("registrationitems").innerHTML = html;
        })
      });
  }

  //Get User profile property for propertyName
  private getUserProfilePropertyForUser(propertyName: string): string {
    var propertyValue = "";
    sp.profiles.myProperties.get().then(function (result) {
      var properties = "";
      var profileProperties = result.UserProfileProperties;
      profileProperties.forEach(function (val) {
        if (val.Key == propertyName) {
          propertyValue = val.Value;
          alert(propertyValue);
        }
      });
    }
    );
    return propertyValue;
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
