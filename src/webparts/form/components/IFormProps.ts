/* eslint-disable @typescript-eslint/no-explicit-any */
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";

export interface IFormProps {
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

// pdfLink: 'https://xencia1.sharepoint.com/:b:/s/XenciaDemoApps/uco/EcFS2u_tQFhMmEy0LV6wx5wBEf8gycMjKYn0RIHHvCVzRw?e=de5FmB' // Link to the PDF
