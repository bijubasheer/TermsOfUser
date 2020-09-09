import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'TermOfUseWpWebPartStrings';
import TermOfUseWp from './components/TermOfUseWp';
import { ITermOfUseWpProps } from './components/ITermOfUseWpProps';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";

export interface ITermOfUseWpWebPartProps {
  description: string;
}

export default class TermOfUseWpWebPart extends BaseClientSideWebPart<ITermOfUseWpWebPartProps> {

  protected onInit(): Promise<void> {

    return super.onInit().then(_ => {
        sp.setup({
        spfxContext: this.context
      });
    });
  }
  public render(): void {
    let email:string = '';
    sp.web.lists.getByTitle("Terms of Use").items.select("Title", "Content", "AcceptedBy").top(1).orderBy("Modified", true).get().then(items =>
      {
        //console.log(items);
        let content = items[0]["Content"];
        let title = items[0]["Title"];
        let acceptedBy:string = items[0]["AcceptedBy"];
        let user = sp.web.currentUser.get().then((u: any) => { 
          email = u.Email;

          if(acceptedBy.indexOf(email) < 0 ) 
          {
            const element: React.ReactElement<ITermOfUseWpProps > = React.createElement(
              TermOfUseWp,
              {
                description: this.properties.description,
                email:u.Email,
                content:content,
                title :title,
              }
            );
          ReactDom.render(element, this.domElement);
        }
      }
    );
   });   
  }

  // private async LoadTermsContent()
  // {
  //   await sp.web.lists.getByTitle("Terms of Use").items.select("Title", "Content").top(1).orderBy("Modified", true).get().then(items =>
  //     {
  //       //console.log(items);
  //       content = items[0]["Content"];
  //       title = items[0]["Title"];
  //     }
  //   );
  // }
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
