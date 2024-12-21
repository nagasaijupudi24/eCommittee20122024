/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @rushstack/no-new-null */
/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import {
  DetailsList,
  IColumn,
  IconButton,
  SelectionMode,
  Modal,
  PrimaryButton,
  mergeStyleSets,
  ChoiceGroup,
  IChoiceGroupOption,
  Dropdown,
} from "@fluentui/react";
import {  IComboBoxOption } from "@fluentui/react/lib/ComboBox";

import PnPPeoplePicker from "../peoplePicker/peoplePicker";
import { v4 } from "uuid";


interface ITableItem {
  id: any;
  key: any;
  comments: any;
  assignedTo: any;
  status: any;
}

interface IDropdownOption {
  id: any;
  key: any;
  text: any;
  email: any;
}


interface IATRAssigneeProps {

  atrType:any;
  getATRTypeOnChange:any;
  clearAtrGridDataOnSelectionOFATRType:any;
  checkingCurrentATRCreatorisCurrentApproverOrNot: any;
  getATRJoinedComments: any;
  
  updategirdData: any;
  commentsData: any;
  sp: any;
  context: any; 

  artCommnetsGridData: any;

  deletedGridData: any;
  approverDetails: any;
  currentATRCreatorDetails: any;
}


interface IATRAssigneeState {
  tableData: any;
  selectedUsers: any;
  currentRowKey: any;
  selectedStatus: any;
  selectedValue: any;
  commentsData: any;
  isModalOpen: boolean;
  modalMessage: string;
  clearPeoplePicker: any;
  atrJoinedComments: any;
  selectedChoice: any;
  selectedDropDownValue:any;

  isDisabled: boolean;
  statusOptions: IDropdownOption[];
}





const choiceOptions: IChoiceGroupOption[] = [
  { key: "Internal", text: "Internal" },
  { key: "External", text: "External" },
];

export class ATRAssignee extends React.Component<
  IATRAssigneeProps,
  IATRAssigneeState
> {
  constructor(props: IATRAssigneeProps) {
    super(props);

    
    this.state = {
      tableData: this.props.artCommnetsGridData,
      selectedUsers: {},
      currentRowKey: null,
      selectedStatus: undefined,
      selectedValue: {},
      commentsData: this.props.commentsData,
      isModalOpen: false,
      modalMessage: "",
      clearPeoplePicker: "",
      atrJoinedComments: "",
      selectedChoice: this.props.atrType,
      statusOptions:[],
      selectedDropDownValue:'',

      isDisabled: !this.props.checkingCurrentATRCreatorisCurrentApproverOrNot,
    };

    
  }


  public componentDidMount(): void {
    this._updateStatusOptions();
  }


  private _updateStatusOptions = () => {
    
    const indexOF = this.props.approverDetails.findIndex(
      (each: any) => each.approverEmail === this.props.currentATRCreatorDetails
    );
    
    const optionUptoATRCreator = this.props.approverDetails.slice(0, indexOF);
    

    const newOptions = optionUptoATRCreator.map((each: any) => {
      return ({
        key:  each.approverEmailName,
        text: each.approverEmailName,
        id: each.id,
        email: each.email,
      })
      
    });

    this.setState({statusOptions:newOptions})


  };


  private columns: IColumn[] = [
    {
      key: "comments",
      name: "Comments",
      fieldName: "comments",
      minWidth: 100,
      maxWidth: 200,
      isResizable: true,
    },
    {
      key: "assignedTo",
      name: "Assigned To",
      fieldName: "assignedTo",
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "status",
      name: "Status",
      fieldName: "status",
      minWidth: 80,
      maxWidth: 100,
      isResizable: true,
    },
    {
      key: "delete",
      name: "Action",
      fieldName: "delete",
      minWidth: 50,
      maxWidth: 75,
      onRender: (item: ITableItem) => (
        <IconButton
          iconProps={{ iconName: "Delete" }}
          disabled={this.state.isDisabled}
          title="Delete"
          ariaLabel="Delete"
          styles={{ root: { paddingBottom: '16px',background:'transparent' } }}
          onClick={() => this.handleDeleteRow(item.key)} 
        />
      ),
    },
  ];

  
  private handleStatusChange = (option: IComboBoxOption | undefined): void => {
    const newStatus = option?.text || "";
    
    this.setState({ selectedValue: option,selectedUsers: option,selectedDropDownValue:newStatus });
   
  };

  
  private handleRowClick = (rowKey: number): void => {
    this.setState({ currentRowKey: rowKey });

  };


  private handleDeleteRow = (rowKey: number): void => {
    const updatedTableData = this.state.tableData.filter(
      (item: { key: number }) => item.key !== rowKey
    );
    this.setState({ tableData: updatedTableData });
    this.props.deletedGridData(updatedTableData);
  };

  public _getDetailsFromPeoplePicker = (): any => {
    console.log("add btn triggered in ATR Assignee")
    
    if (Object.keys(this.state.selectedValue).length === 0) {
     
      this.setState({
        isModalOpen: true,
        modalMessage: "Please select the Assignee then click on Add.",
      });
    } else {
      const itemExists = this.state.tableData.some(
        (item: ITableItem) => item.id === this.state.selectedUsers.id
      );

      if (itemExists) {
        this.setState({
          isModalOpen: true,
          modalMessage:
            "The selected assignee already exist. Kindly choose another assignee.",
        });
        return;
      }
    

      if (this.props.currentATRCreatorDetails === this.state.selectedValue.email) {
        this.setState({
          isModalOpen: true,
          modalMessage:
            "Current Approver cannot be  assignee.",
        });
        return;
      }

   
      const joinedCommentsData = this.state.commentsData
        .filter((each: any) => !!each)
        .map((each: any) => `${each?.pageNum} ${each?.page} ${each?.comment}`);

      

      const newTableData = {
        key: v4(),
        comments: joinedCommentsData.join(", "),
        assignedTo: this.state.selectedValue.text,
        status: "Submitted",

        ...this.state.selectedValue,
      };
      this.setState(
        { atrJoinedComments: joinedCommentsData.join(",") },
        this.props.getATRJoinedComments(joinedCommentsData.join(", "))
      );
      

      this.props.updategirdData({
        assigneeDetails: this.state.selectedValue,
        comments: [...this.state.tableData, newTableData],
        atrType:this.state.selectedChoice
      });
      // });
      this.props.getATRJoinedComments(joinedCommentsData.join(", "));
     

    

      this.setState((prev) => ({
        tableData: [...prev.tableData, newTableData], // Add newData to the array
        selectedValue: {},
        selectedUsers: {},
        selectedDropDownValue: ''
      }));
      // eslint-disable-next-line no-unused-expressions
      (this.state.selectedChoice === "External" && Object.keys(this.state.selectedValue).length > 0) && this.state.clearPeoplePicker();
    }
  };

  public _getDetailsFromPeoplePickerData = (data: any, type: any): any => {
 

    this.setState({ selectedValue: data[0], selectedUsers: data[0] });

   
  };

  private _closeModal = (): void => {
    this.setState({ isModalOpen: false });
    this.setState({
      
      selectedValue: {},selectedUsers:{},selectedDropDownValue:''
    });
    // eslint-disable-next-line no-unused-expressions
    (this.state.selectedChoice === "External" && Object.keys(this.state.selectedValue).length > 0) &&this.state.clearPeoplePicker();
  };

 
  private onChoiceChange = (
    ev: React.FormEvent<HTMLElement>,
    option?: IChoiceGroupOption
  ): void => {
    if (option) {
      this.setState({ selectedChoice: option.key,selectedValue:{},tableData:[],selectedDropDownValue:'' ,selectedUsers:{}});
      // console.log("Selected choice:", option.key);
      this.props.getATRTypeOnChange(option.key)
      this.props.clearAtrGridDataOnSelectionOFATRType()

    }
  };

  public render(): React.ReactElement<IATRAssigneeProps> {
    console.log(this.state)
    const { tableData } = this.state;
  

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
      peoplePickerAndAddCombo: {
        display: "flex",
        gap: "5px",
        width: "80%",
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
      dropdownFullWidth: {
        minWidth: '180px'
      }
    });

    const { isModalOpen, modalMessage, selectedChoice, isDisabled,selectedDropDownValue } =
      this.state;
     

    return (
      <div>
        <div>
          <ChoiceGroup
          disabled={isDisabled}
            selectedKey={selectedChoice}
            options={choiceOptions}
            onChange={this.onChoiceChange}
            required={true}
            styles={{
              flexContainer: {
                marginBottom: "5px",
                display: "flex",
                flexDirection: "row", 
              },
              root: {
                selectors: {
                  ".ms-ChoiceField-input": {
                    transform: "scale(0.8)", 
                  },
                  ".ms-ChoiceField-field": {
                    padding: "0 8px", 
                  },
                  ".ms-Label": {
                    fontSize: "0.85rem",
                  },
                },
              },
            }}
          />
        </div>

       
        <div className={styles.peoplePickerAndAddCombo}>
          
            {this.state.selectedChoice === "Internal" ? (
              <Dropdown
              className={styles.dropdownFullWidth}
                disabled={isDisabled}
                options={this.state.statusOptions}
                selectedKey={selectedDropDownValue}
                onChange={(event, option) => this.handleStatusChange(option)}
                
              />
            ) : (
              <PnPPeoplePicker
                disabled={isDisabled}
                context={this.props.context}
                spProp={this.props.sp}
                getDetails={this._getDetailsFromPeoplePickerData}
                typeOFButton="atr"
                clearPeoplePicker={(data: any, funtionName: any) => {
                  
                  this.setState({ clearPeoplePicker: data });
                }}
              />
            )}
        
          
            <PrimaryButton
              disabled={isDisabled}
              iconProps={{ iconName: "Add" }}
              onClick={this._getDetailsFromPeoplePicker}
            >
              Add
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
          onItemInvoked={(item: ITableItem) => this.handleRowClick(item.key)}
        />

       
        <Modal
          isOpen={isModalOpen}
          onDismiss={this._closeModal}
          isBlocking={true}
          containerClassName={styles.modal}
        >
          <div className={styles.header}>
            <div style={{ display: "flex", alignItems: "center" }}>
              <IconButton iconProps={{ iconName: "Info" }} />
              <h4 className={styles.headerTitle}>Alert</h4>
            </div>
            <IconButton
              iconProps={{ iconName: "Cancel" }}
              onClick={this._closeModal}
            />
          </div>
          <div className={styles.body}>
            <p>{modalMessage}</p>
          </div>
          <div className={styles.footer}>
            <PrimaryButton
              iconProps={{ iconName: "ReplyMirrored" }}
              onClick={this._closeModal}
              text="OK"
            />
          </div>
        </Modal>
      </div>
    );
  }
}
