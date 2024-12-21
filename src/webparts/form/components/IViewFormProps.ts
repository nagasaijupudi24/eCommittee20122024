/* eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";



export interface IViewFormProps {
    description: string;
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
    sp: SPFI; // Pass the configured sp object
    context: WebPartContext; // Pass the WebPartContext
    listId:any;
    libraryId:any;
    formType:any;
    homePageUrl:any;
    passCodeUrl:any;
    existPageUrl:any;
  }


  