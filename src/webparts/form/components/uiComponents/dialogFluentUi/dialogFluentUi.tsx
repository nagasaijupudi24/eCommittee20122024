/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { Modal } from "@fluentui/react/lib/Modal";
import { PrimaryButton, DefaultButton } from "@fluentui/react/lib/Button";
import {  IIconProps, mergeStyleSets, Stack, TextField,IconButton, Text } from "@fluentui/react";
import PnPPeoplePicker from "../peoplePicker/peoplePicker";

import { v4 } from "uuid";
import SpanComponent from "../spanComponent/spanComponent";


interface IDialogProps {
  changeApproverDataMandatory:any;
  referCommentsAndDataMandatory:any;
  statusNumberForChangeApprover:any;
  referDto:any;
  requesterEmail:any;
  dialogUserCheck:any;
  hiddenProp: any;
  dialogDetails: any;
  sp: any;
  context: any;
 
  fetchAnydata: any;
  fetchReferData:any;
  isUserExistingDialog:any;
  
}

const Header = (props: any) => (
  <Stack
    horizontal
    horizontalAlign="space-between"
    verticalAlign="center"
    styles={{ root: { padding: "10px", borderBottom: "1px solid #ccc" } }}
  >
    <Stack horizontal verticalAlign="center">
    
        <IconButton iconProps={{ iconName: "Info" }} />
     
      <Text variant="large" styles={{ root: { marginLeft: "3px",fontSize:'16px' } }}>
        {props.heading}
      </Text>
    </Stack>
    <IconButton iconProps={{ iconName: "Cancel" }} onClick={props.onClose} />
  </Stack>
);




export const DialogBlockingExample: React.FunctionComponent<IDialogProps> = (props,) => {
  const {
    dialogUserCheck,
    hiddenProp,
    dialogDetails,
    context,
    sp,
    fetchAnydata,
    isUserExistingDialog,
    requesterEmail
    
  } = props
 
  const styles = mergeStyleSets({
    modal: {
      minWidth: '300px',
      maxWidth: '80vw',
      width: '100%',
      '@media (min-width: 768px)': {
        maxWidth: '580px', 
      },
      '@media (max-width: 767px)': {
        maxWidth: '290px', 
      },
      margin: 'auto',
      padding: '10px',
      backgroundColor: 'white',
      borderRadius: '4px',
      boxShadow: '0 2px 8px rgba(0, 0, 0, 0.26)',
    },
    header: {
      display: 'flex',
      justifyContent: 'space-between',
      alignItems: 'center',
      borderBottom: '1px solid #ddd',
        height:'50px'
    },
    headerTitle: {
     margin:'5px',
     marginLeft:'5px',
     fontSize:'16px',
   fontWeight:'400'
    },
    headerIcon: {
     paddingRight: '0px', 
     
    },
    body: {
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      textAlign: 'center',
      padding: '20px 0',
      height:'100%',
      '@media (min-width: 768px)': {
        marginLeft: '20px', 
        marginRight: '20px', 
      },
      '@media (max-width: 767px)': {
        marginLeft: '20px',
        marginRight: '20px',
      } 
    },
    contentContainer:{
      width:'70%',
      display:'flex',
      flexDirection:'column',
      justifyContent: 'flex-start',
      alignItems:'flex-start'

    },
    footer: {
      display: 'flex',
      justifyContent: 'space-between', 
      
      borderTop: '1px solid #ddd',
      paddingTop: '10px',
      minHeight:'50px'
    },
    button: {
      maxHeight:'32px',
      flex: '1 1 50%', 
      margin: '0 5px', 
    },
    buttonContent: {
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
    },
    buttonIcon: {
      marginRight: '4px',
    },
  
    removeTopMargin:{
      marginTop: '4px',
      marginBottom: '4px'
    }
    
  });
  const [data, setData] = React.useState<any>('');
  const [isUserExistsModalVisible, setIsUserExistsModalVisible] = React.useState(false);
    React.useState<any>('');
    
  const [referredCommentTextBoxValue, setReferredCommentTextBoxValue] =
    React.useState<any>('');

   

  const handleConfirmBtn = () => {
  
    dialogDetails.functionType(
      dialogDetails.status === "Noted"?"Approved":dialogDetails.status,
      dialogDetails.statusNumber
    );
  };

  const closeIcon: IIconProps = { iconName: "Cancel" };

  const getGeneralDialogJSX = (): any => {
   
    return (
      <Modal
        isOpen={!hiddenProp}
        onDismiss={dialogDetails.closeFunction}
        isBlocking={true}
        containerClassName={styles.modal}
       
      >
        <div  className={styles.header}>
          <div style={{ display: 'flex', alignItems: 'center' }}>
           
            <IconButton iconProps={{ iconName: "WaitlistConfirm" }} />
            <h4 className={styles.headerTitle}>Confirmation</h4>
          </div>
          <IconButton iconProps={closeIcon} onClick={dialogDetails.closeFunction} />
        </div>
        <div
        className={styles.body}
      
        >
          <p >{dialogDetails.subText}</p>
          <p style={{textAlign:'center'}}>{dialogDetails.message}</p>
        </div>

        
        <div className={styles.footer}>
          <PrimaryButton
          styles={{ root: styles.buttonContent }}
          className={styles.button}
          
          iconProps={{ iconName: "SkypeCircleCheck" }} onClick={handleConfirmBtn} text="Confirm"  />
          <DefaultButton
          styles={{ root: styles.buttonContent }}
          className={styles.button}
          iconProps={{ iconName: "ErrorBadge" }} onClick={dialogDetails.closeFunction} text="Cancel" />
        </div>
      </Modal>
    );
  };


  const checkReviewer = (data:any): boolean => {
    const approverTitles = dialogUserCheck.peoplePickerApproverData.map(
      (each: any) => each.text
    );
 
    const reviewerTitles = dialogUserCheck.peoplePickerData.map(
      (each: any) => each.text
    );
    
  
    const reviewerInfo = data[0];
  
    const reviewerEmail = reviewerInfo.email || reviewerInfo.secondaryText;
    
    const reviewerName = reviewerInfo.text;
  
  
    const isReviewerOrApprover =
      reviewerTitles.includes(reviewerName) ||
      approverTitles.includes(reviewerName);

     
    
    const isCurrentUserReviewer = context.pageContext.user.email === reviewerEmail;
  
    const isRequester = reviewerInfo.email === requesterEmail

    
    console.log(props.referDto)

   

    if (props.dialogDetails.type === 'Refer'){
      return isReviewerOrApprover || isCurrentUserReviewer ||isRequester 

    }else{
      if (props.statusNumberForChangeApprover === '4000'){
        const isSelectedUserIsAnReferee =(Object.keys(props.referDto).length > 0) ? props.referDto.referrerEmail ===  reviewerInfo.email :false
        console.log(isSelectedUserIsAnReferee)
        console.log(props.dialogDetails)
        return isReviewerOrApprover || isCurrentUserReviewer ||isRequester ||isSelectedUserIsAnReferee;

      }
      return isReviewerOrApprover || isCurrentUserReviewer ||isRequester 
      
      

    }
    
  
   
    
  };
  
  
  

  const _getDetails = (data: any, typeOFButtonTriggererd: any): any => {
  
    
    setData(data);
  
   
  
    
    fetchAnydata(data, typeOFButtonTriggererd, dialogDetails.status);
  };

  const handleChangeApporver = () => {

    if (dialogDetails.referPassFuntion !==''){
      dialogDetails.referPassFuntion()

    }

    


    if (dialogDetails.functionType !==''){
    
      dialogDetails.functionType(
        dialogDetails.status,
        dialogDetails.statusNumber
      );
    }
   
  };

  const handleReferData = () => {
   

    if (dialogDetails.referPassFuntion !==''){
      dialogDetails.referPassFuntion()

    }

    


    if (dialogDetails.functionType !==''){
      dialogDetails.functionType(
        dialogDetails.status,
        dialogDetails.statusNumber,
        referredCommentTextBoxValue
      );

    }


   

    props.fetchReferData(referredCommentTextBoxValue)
   
  };

  const closeUserExistsModal = () => {
    setIsUserExistsModalVisible(false);
  };

  const getUserExistsModalJSX = (): any => {
  
    return (
      <Modal
        isOpen={isUserExistsModalVisible}
        onDismiss={closeUserExistsModal}
        isBlocking={true}
        styles={{
          main: {
            width: "100%",
            maxWidth: "290px",
            "@media (min-width: 768px)": {
              maxWidth: "580px",
            },
          },
        }}
      >
       
        <div style={{
          display: "flex",
          flexDirection: "row",
          justifyContent: "space-between",
          alignItems: "center",
          padding: "8px 12px",
          borderBottom: "1px solid #ddd",
        }}>
        
          <div style={{
            display: "flex",
            alignItems: "center",
            gap: "8px",
          }}>
            <IconButton iconProps={{ iconName: "Info" }} />
           
            <h4 className={styles.headerTitle}>Alert</h4>
          </div>
  
         
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Close modal"
            onClick={closeUserExistsModal}
          />
        </div>
  
      
        <Stack tokens={{ padding: "16px" }} horizontalAlign="center" verticalAlign="center">
          <Text style={{ margin: "16px 0", fontSize: "14px", textAlign: "center" }}>
          The selected approver cannont be same as existing Reviewers/Requester/referee/CurrentActioner
          </Text>
        </Stack>
  
      
        <div style={{
          display: "flex",
          justifyContent: "flex-end",
          padding: "12px 16px",
          borderTop: "1px solid #ddd",
        }}>
          <PrimaryButton
          iconProps={{ iconName: 'ReturnToSession', styles: { root: styles.buttonIcon } }}
           
            text="ok"
            onClick={closeUserExistsModal}
            ariaLabel="Close modal"
          />
        </div>
      </Modal>
    );
  };
  

  const getChangeApproverJsx = (): any => {
   
  
  
  
    return (
      <Modal
        isOpen={!hiddenProp}
        onDismiss={dialogDetails.closeFunction}
        isBlocking={true}
        containerClassName={styles.modal}
      >
       
        <Header heading={'Change Approver'} onClose={dialogDetails.closeFunction} />
        <div className={styles.body} style={{paddingTop:'10px'}}>
         

          <div style={{ width: "90%" }}>
            <div style={{width:'100%'}}>
            <p style={{textAlign:'left'}}>{dialogDetails.message}<SpanComponent/></p>
            <PnPPeoplePicker
              context={context}
              spProp={sp}
              getDetails={_getDetails}
              typeOFButton="Change Approver" clearPeoplePicker={undefined} disabled={true}   />

            </div>
             
            </div>
          
         
        </div>
        <div className={styles.footer}>
          <PrimaryButton  styles={{ root: styles.buttonContent }} iconProps={{ iconName: "SkypeCircleCheck" }} className={styles.button} onClick={
            
           
            ()=>{
              console.log(data)
              if (data ===''){

               
                props.changeApproverDataMandatory()
                return
              }
              if (checkReviewer(data)) {
                dialogDetails.closeFunction()
                isUserExistingDialog()
              
              return;
            }
            
            
            handleChangeApporver()
            } }
            text="Submit" />
          <DefaultButton  styles={{ root: styles.buttonContent }} iconProps={{ iconName: "ErrorBadge" }} className={styles.button} onClick={dialogDetails.closeFunction} text="Cancel" />
          </div>
      </Modal>
    );
  };
  const getReferJSX = (): any => {
 
    return (
      <Modal
        isOpen={!hiddenProp}
        onDismiss={dialogDetails.closeFunction}
        isBlocking={true}
        containerClassName={styles.modal}
      >
       
        <div>
          <Header heading={'Add Refree'} onClose={dialogDetails.closeFunction} />
          <div
            style={{
            
              display: "flex",
              flexDirection: "column",
              justifyContent: "center",
              alignItems: "center",
              width: "100%",
              padding: "20px",
              paddingTop:'5px'
            }}
          >
            <div style={{ width: "90%" }}>
            <h4 className={styles.headerTitle}>{dialogDetails.message[0]}</h4>
             
              <PnPPeoplePicker
                context={context}
                spProp={sp}
                getDetails={_getDetails}
                typeOFButton="Refer" clearPeoplePicker={undefined} 
                disabled={true}              />
            </div>
            <div style={{width:'90%'}}>
            <h4 className={styles.headerTitle} style={{alignSelf:'flex-start'}}>{dialogDetails.message[1]}</h4>
            <TextField
             
              multiline
              rows={3}
              onChange={(
                _: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
                newText: string
              ): void => {
               
                setReferredCommentTextBoxValue(() => {
                 
                  const commentsObj = {
                    id: v4(),
                    pageNum: "N/A",
                    page: "N/A",
                    comment: newText,
                    commentedBy: context.pageContext.user.displayName,
                    commentsFrom: dialogDetails.status,
                    commentedEmail: context.pageContext.user.email,
                  };
                
                  return commentsObj;
                });
              }}
              styles={{ root: { width: "100%" } }}
            />

            </div>
           
          </div>
          <div className={styles.footer}
          >
            <PrimaryButton
              onClick={()=>{
                if (data ===''){

                
                  props.referCommentsAndDataMandatory()
                }else if(referredCommentTextBoxValue===''){
                 
                  props.referCommentsAndDataMandatory()
                  
                }else{
               
                  if (checkReviewer(data)) {
                    dialogDetails.closeFunction()
                    isUserExistingDialog()
                  
                  return; 
                }

                // }

                 


                  
                  handleReferData()
                }


              }}
              className={styles.button}
              text="Confirm"
              iconProps={{ iconName: "SkypeCircleCheck" }}
              styles={{ root: styles.buttonContent }}
              
            />
            <DefaultButton
            className={styles.button}
              onClick={dialogDetails.closeFunction}
              text="Cancel"
              iconProps={{ iconName: "ErrorBadge" }}
              styles={{ root: styles.buttonContent }}
             
            />
          </div>
        </div>
      </Modal>
    );
  };

  switch (props.dialogDetails.type) {
   
    case "Change Approver":
      return  <>
      {getChangeApproverJsx()}
      {getUserExistsModalJSX()}
    </>
    case "Refer":
      return  <>
      {getReferJSX()}
      {getUserExistsModalJSX()} 
    </>
      default:
        return (
          <>
            {getGeneralDialogJSX()}
            {getUserExistsModalJSX()} 
          </>
        );
  }
};
