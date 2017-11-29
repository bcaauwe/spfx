import * as pnp from 'sp-pnp-js';

import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './UserProfileInformationWebPart.module.scss';
import * as strings from 'UserProfileInformationWebPartStrings';

export interface IUserProfileInformationWebPartProps {
  description: string;
}

export default class UserProfileInformationWebPartWebPart extends BaseClientSideWebPart<IUserProfileInformationWebPartProps> {

private GetUserPropertiesByName(email): void {
  alert("Processing Properties for '" + email + "'");

  pnp.sp.profiles.getUserProfilePropertyFor(email, "PreferredName").then(result => {
    console.log("PreferredName '" + result + "'");
  });

  pnp.sp.profiles.getUserProfilePropertyFor(email, "Title").then(result => {
    console.log("Title '" + result + "'");
  });

  pnp.sp.profiles.getUserProfilePropertyFor(email, "Office").then(result => {
    console.log("Office '" + result + "'");
  });
}

private GetUserProperties(): void {
  pnp.sp.profiles.myProperties.get().then(function(result) {
    var userProperties = result.UserProfileProperties;
    var userPropertyValues = "";

    userProperties.forEach(function(property) {
      userPropertyValues += property.Key + " - " + property.Value + "<br/>";
    });

    document.getElementById("spUserProfileProperties").innerHTML = userPropertyValues;
  }).catch(function(error) {
    console.log("Error: '" + error + "'");
  });
}

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.userProfileInformation}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-lg10 ms-xl8 ms-xlPush2 ms-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Your Office 365 User Information</span>
            </div>
          </div>
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div id="spUserProfileProperties" />
          </div>          
        </div>
      </div>`;

      pnp.sp.web.currentUser.get().then(result => {
        this.GetUserPropertiesByName(result.LoginName);
      });
      this.GetUserProperties();
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
