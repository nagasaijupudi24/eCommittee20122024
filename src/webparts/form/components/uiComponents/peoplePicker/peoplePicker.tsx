/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */

import * as React from "react";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";


export interface IPnPPeoplePickerProps {
  disabled:any;
  context: WebPartContext;
  spProp: any;
  getDetails:any;
  typeOFButton:any;
  clearPeoplePicker:any;

}

export interface IPnPPeoplePickerState {
  selectedPeople: any[];
  key: any;
  peoplePickerData: any[];
}

export default class PnPPeoplePicker extends React.Component<
  IPnPPeoplePickerProps,
  IPnPPeoplePickerState
> {
  constructor(props: IPnPPeoplePickerProps) {
    super(props);
    this.state = {
      selectedPeople: [],
      peoplePickerData: [],
      key: 0, 
    };
  }

  

  private _clearPeoplePicker = () => {
   
    this.setState((prev)=>({
      selectedPeople: [], key: prev.key + 1

    }))
  
  };

  private _getUserProperties = async (loginName: any): Promise<any> => {
   
    let designation = "NA";
    let email = "NA";
   
    const profile = await this.props.spProp.profiles.getPropertiesFor(
      loginName
    );
    
    designation = profile.Title;
    email = profile.Email;
   
    const props: any = {};
    profile.UserProfileProperties.forEach(
      (prop: { Key: string | number; Value: any }) => {
        props[prop.Key] = prop.Value;
      }
    );

    profile.userProperties = props;
    
    return [designation, email];
  };

 

  private _getPeoplePickerItems = async (items: any[]) => {
    

    const dataRec = await this._getUserProperties(items[0].loginName);
    

    if (typeof dataRec[0]?.toString() === "undefined") {
      const newItemsDataNA = items.map(
        (obj: { [x: string]: any; loginName: any }) => {
          
          return {
            ...obj,
            optionalText: "N/A",
            
            email: obj.secondaryText,
          };
        }
      );
     
      this.setState({ selectedPeople: newItemsDataNA });
    } else {
      const newItemsData = items.map((obj: { loginName: any }) => {
        return {
          ...obj,
          optionalText: dataRec[0],
          
          email: dataRec[1],
          srNo: dataRec[1].split("@")[0],
        };
      });
      
      this.props.getDetails(newItemsData,this.props.typeOFButton)
      // eslint-disable-next-line no-unused-expressions
      newItemsData.length > 0 && this.props.clearPeoplePicker(this._clearPeoplePicker,"clearFuntion")
      this.setState({ selectedPeople: newItemsData });
      
    }
  };

  public render(): React.ReactElement<IPnPPeoplePickerProps> {
   
    const peoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };

    

    return (
      <div style={{ minWidth: '180px!important' }}>
        <PeoplePicker
          key={this.state.key}
          context={peoplePickerContext}
        
          personSelectionLimit={1}
          groupName={""}
          showtooltip={true}
          disabled={this.props.typeOFButton ==='atr' && this.props.disabled}
          ensureUser={true}
          onChange={this._getPeoplePickerItems.bind(this)}
          principalTypes={[PrincipalType.User]}
          resolveDelay={1000}
          styles={{ root: {minWidth: '180px!important' } }}
        />
       
      </div>
    );
  }
}
