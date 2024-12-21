/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @microsoft/spfx/pair-react-dom-render-unmount */
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'FormWebPartStrings';
import Form from './components/uiComponents/Form';
import ViewForm from './components/uiComponents/view';
import { IFormProps } from './components/IFormProps';
import { IViewFormProps } from './components/IViewFormProps';

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/files";  
import "@pnp/sp/site-users/web";
import "@pnp/sp/webs"; 
import "@pnp/sp/lists"; 
import "@pnp/sp/items"; 
import "@pnp/sp/files/web";
import '@pnp/sp/profiles';



export interface IFormWebPartProps {
  FormType: string;
  description: string;
  listId:any;
  libraryId:any;
  homePageUrl:any;
  passCodeUrl:any;
  existPageUrl:any;
}

export {};
declare global {
  interface Window {
      AdobeDC: any;
  }
}


export default class FormWebPart extends BaseClientSideWebPart<IFormWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private sp: ReturnType<typeof spfi>;

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.sp = spfi().using(SPFx(this.context));

    this._environmentMessage = await this._getEnvironmentMessage();
  }





  public render(): void {
    let element: React.ReactElement<IFormProps> | React.ReactElement<IViewFormProps> | null = null;

    const newFormProps =  {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      sp: this.sp, 
      context: this.context, 
      listId:this.properties.listId,
      libraryId:this.properties.libraryId,
      formType:this.properties.FormType,
      homePageUrl:this.properties.homePageUrl,
      passCodeUrl:this.properties.passCodeUrl,
      existPageUrl:this.properties.existPageUrl
    }

    const viewFormProps =  {
      description: this.properties.description,
      isDarkTheme: this._isDarkTheme,
      environmentMessage: this._environmentMessage,
      hasTeamsContext: !!this.context.sdks.microsoftTeams,
      userDisplayName: this.context.pageContext.user.displayName,
      sp: this.sp, 
      context: this.context, 
      listId:this.properties.listId,
      libraryId:this.properties.libraryId,
      formType:this.properties.FormType,
      homePageUrl:this.properties.homePageUrl,
      passCodeUrl:this.properties.passCodeUrl,
      existPageUrl:this.properties.existPageUrl
    }

    if (this.properties.FormType === "New") {
      element = React.createElement(
        Form,
        newFormProps
      );
     
    }
    else if (this.properties.FormType === "View") {
      element = React.createElement(
        ViewForm,viewFormProps
       
      );
     
    }
    else if (this.properties.FormType === "BoardNoteView") {
      element = React.createElement(
        ViewForm,
        viewFormProps
      );
     
    }
    else if (this.properties.FormType === "BoardNoteNew") {
      element = React.createElement(
        Form,
        newFormProps
      );
     
    }
 
      

    if (element !== null) {
      ReactDom.render(element, this.domElement);
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { 
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': 
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': 
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': 
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

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
               
               
                PropertyFieldListPicker('listId', {
                  label: 'Select a list',
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  includeListTitleAndUrl: true,
                  orderBy: PropertyFieldListPickerOrderBy.Id,
                  disabled: false,
                  baseTemplate: 100,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                 
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  multiSelect: false,
                }),
                PropertyFieldListPicker('libraryId', {
                  label: 'Select a Library',
                  selectedList: this.properties.libraryId,
                  includeHidden: false,
                  includeListTitleAndUrl: true,
                  orderBy: PropertyFieldListPickerOrderBy.Id,
                  disabled: false,
                  baseTemplate: 101,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                 
                  deferredValidationTime: 0,
                  key: 'listPickerFieldId',
                  multiSelect: false,
                }),
                 
                PropertyPaneDropdown('FormType', {
                  label: "Form Type",
                  selectedKey: 'New',
                  options: [
                    { key: 'New', text: 'eCommittee New' },
                    { key: 'View', text: 'eCommittee View' },
                   
                    { key: 'BoardNoteNew', text: 'BoardNote New' },
                    { key: 'BoardNoteView', text: 'BoardNote View' }
                    


                  ]
                }),
                PropertyPaneTextField('homePageUrl', {
                  label: "Home Page URL",
                 
                  value: this.properties.homePageUrl,
                  resizable: true,
          
                }),
                
                PropertyPaneTextField('passCodeUrl', {
                  label: "Create Passcode URL",
                
                  value: this.properties.passCodeUrl,
                  resizable: true,
                 
                }),
                PropertyPaneTextField('existPageUrl', {
                  label: "Exist Page URL",
                 
                  value: this.properties.existPageUrl,
                  resizable: true,
                  
                }),
              ]
            },
            
          ]
        }
      ]
    };
  }
}