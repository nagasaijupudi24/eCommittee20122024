/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import {
  DetailsList,
  IColumn,
 
  IconButton,
 
  Modal,

  PrimaryButton,
  SelectionMode,
} from "@fluentui/react";
import { mergeStyleSets } from "@fluentui/react/lib/Styling";
import PnPPeoplePicker from "../peoplePicker/peoplePicker";


interface ITableItem {
  id: any;
  comments: any;
  assignedTo: any;
  status: any;
 
}


interface IATRAssigneeProps {
  updategirdData: any;
  sp: any;
  context: any; 
  artCommnetsGridData: any;
  submitFunctionForMarkInfo: any;
  deletedGridData: any;
  homePageUrl:any
}


interface IATRAssigneeState {
  submitBtnVisable:any;
  tableData: any;
  selectedUsers: any;
  currentRowKey: any;
  selectedStatus: any;
  selectedValue: any;
  isModalOpen: boolean;
  modalMessage: string;
  clearPeoplePicker:any;
  warnType:any;
}

export class MarkInfo extends React.Component<
  IATRAssigneeProps,
  IATRAssigneeState
> {
  constructor(props: IATRAssigneeProps) {
    super(props);

    
    this.state = {
      submitBtnVisable:false,
      tableData: this.props.artCommnetsGridData,
      selectedUsers: {},
      currentRowKey: null,
      selectedStatus: undefined,
      selectedValue:{},
      isModalOpen: false,
      modalMessage: "",
      clearPeoplePicker:"",
      warnType:''
    };
  }

  
  private columns: IColumn[] = [
    {
      key: "serialNo",
      name: "S.No",
      minWidth: 50,
      maxWidth: 75,
      isResizable: false,
      onRender: (_item: any, _index?: number) => (
        <span>{(_index !== undefined ? _index : 0) + 1}</span>
      ),
    },
    {
      key: "text",
      name: "User Info",
      fieldName: "text",
      minWidth: 100,
      maxWidth: 290,
      isResizable: true,
    },
    {
      key: "delete",
      name: "Action",
      fieldName: "delete",
      minWidth: 100,
      maxWidth: 150,
      onRender: (item: ITableItem) => (
        <IconButton
          iconProps={{ iconName: "Delete" }}
          title="Delete"
          ariaLabel="Delete"
          styles={{ root: { paddingBottom: '16px' } }}
          onClick={() => this.handleDeleteRow(item.id)} 
        />
      ),
    },
  ];


  private handleDeleteRow = (rowKey: number): void => {
    const updatedTableData = this.state.tableData.filter(
      (item: { id: number }) => item.id !== rowKey
    );
    this.setState({ tableData: updatedTableData,selectedValue:[], submitBtnVisable:true});
    this.props.deletedGridData(updatedTableData);
  };

  public _handleAdd = (): any => {
    const { tableData, selectedValue } = this.state;
  
    
    if (Object.keys(this.state.selectedValue).length === 0) {
      this.state.clearPeoplePicker()
      this.setState({
        isModalOpen: true,
        modalMessage: "Please select the user then click on Add User.",
        warnType:'no'
      });
      return;
    }
  
    if (tableData.length >= 10) {
      this.state.clearPeoplePicker()
      this.setState({
        isModalOpen: true,
        modalMessage: "You cannot add more than 10 items.",
         warnType:'no'
      });
      return;
    }
  
    const itemExists = tableData.some(
      (item: ITableItem) => item.id === selectedValue.id
    );
  
    if (itemExists) {
      this.state.clearPeoplePicker()
      this.setState({
        isModalOpen: true,
        modalMessage: "The selected user already exist. Kindly choose another user.",
         warnType:'no'
      });
      return;
    }
  
    this.props.updategirdData({
      markInfoassigneeDetails: selectedValue,
    });

    this.state.clearPeoplePicker()
    this.setState({selectedValue:{}})
  
    if (Object.keys(this.state.selectedValue).length > 0) {
      

      this.setState(
        (prev)=>({
          tableData:[...prev.tableData,selectedValue],submitBtnVisable:true
        })
      )
    }
  };
  

  public _getDetailsFromPeoplePickerData = (data: any, type: any): any => {
    
    this.setState({ selectedValue: data[0] });
  };

  private _closeModal = (): void => {
    this.setState({ isModalOpen: false });
  };

  private _handleSubmit =async (): Promise<void> => {
    if (this.state.tableData.length === 0) {
      this.setState({
        
        isModalOpen: true,
        modalMessage: "Please select a user and click Add.",
         warnType:'no'
      });
      return;
    }

    await this.props.submitFunctionForMarkInfo();
    this.setState({
      isModalOpen: true,
      modalMessage: "The mark for information has been updated successfully.",
       warnType:'yes',
       submitBtnVisable:false
    });
  };

  public render(): React.ReactElement<IATRAssigneeProps> {
    const { tableData, isModalOpen, modalMessage } = this.state;
 
    console.log(this.state)

    const styles = mergeStyleSets({
      modal: {
        padding: "10px",
        minWidth: "300px",
        maxWidth: "80vw",
        width: "100%",
        "@media (min-width: 768px)": {
          maxWidth: "580px", 
        },
        "@media (max-width: 767px)": {
          maxWidth: "290px",
        },
        margin: "auto",
        backgroundColor: "white",
        borderRadius: "4px",
        boxShadow: "0 2px 8px rgba(0, 0, 0, 0.26)",
      },
      header: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        borderBottom: "1px solid #ddd",
        minHeight: "50px",
      },
      headerTitle: {
        margin: "5px",
        marginLeft: "5px",
        fontSize: "16px",
        fontWeight: "400",
        
      },
      peoplePickerAndAddCombo:{
        display:'flex',
        gap:'5px',
        width:'80%',
        flexWrap:'wrap',


      },
      body: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        textAlign: "center",
        padding: "20px 0",
      },
      footer: {
        display: "flex",
        justifyContent: "flex-end",
        marginTop: "20px",
        borderTop: "1px solid #ddd", 
        paddingTop: "10px",
      },
    });

    return (
      <>
        
        <div className={styles.peoplePickerAndAddCombo}>
          <PnPPeoplePicker
            context={this.props.context}
            spProp={this.props.sp}
            getDetails={this._getDetailsFromPeoplePickerData}
            typeOFButton="markInfo"
            clearPeoplePicker={(data: any, funtionName: any) => {
            
              this.setState({ clearPeoplePicker: data });
            } } disabled={true}         />

          <PrimaryButton
            iconProps={{ iconName: "Add" }}
            onClick={this._handleAdd}
          >
            Add User
          </PrimaryButton>
        </div>

       
        <DetailsList
          items={tableData}
          columns={this.columns}
          setKey="set"
          layoutMode={0} 
          selectionMode={SelectionMode.none}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        />

        <div style={{  marginTop: "10px" }}>
          {this.state.submitBtnVisable?
            <PrimaryButton
            iconProps={{ iconName: "Save" }}
            onClick={this._handleSubmit}
          >
            Submit
          </PrimaryButton>:''}
        
        </div>

        
        <Modal
          isOpen={isModalOpen}
          onDismiss={this._closeModal}
          isBlocking={true}
          containerClassName={styles.modal}
        >
          <div className={styles.header}>
            <div style={{ display: "flex", alignItems: "center" }}>
              
              <IconButton iconProps={{ iconName: 'Info' }}/>
              <h4 className={styles.headerTitle}>Alert</h4>
            </div>
            <IconButton
          
              iconProps={{ iconName: 'Cancel' }}
            
              
              onClick={this._closeModal}
            />
          </div>
          <div className={styles.body}>
            <p>{modalMessage}</p>
          </div>
          <div className={styles.footer}>
            <PrimaryButton
              iconProps={{ iconName: "ReplyMirrored" }}
              
              
              onClick={() => {
                if (this.state.warnType !=="no"){
                  const pageURL: string = this.props.homePageUrl;
                  window.location.href = `${pageURL}`;
                

                }
                this._closeModal()
              }}
              text="OK"
            />
          </div>
        </Modal>
      </>
    );
  }
}



