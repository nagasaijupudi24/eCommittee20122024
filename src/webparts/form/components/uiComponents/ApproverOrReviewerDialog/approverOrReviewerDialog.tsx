/* eslint-disable @typescript-eslint/no-unused-vars */

import * as React from "react";
import {
  Modal,
  PrimaryButton,
  IconButton,
  IIconProps,
  IModalStyles,
  Stack,
  Text,
  
} from "@fluentui/react";


interface MyModalProps {
  hidden: boolean;
  handleDialogBox: () => void;
}


const closeIcon: IIconProps = { iconName: 'Cancel' };
const okIcon: IIconProps = { iconName: 'ReturnToSession' }; 

const ApproverOrReviewerModal: React.FC<MyModalProps> = ({
  hidden,
  handleDialogBox,
}) => {
 
  const headerStyles: React.CSSProperties = {
    display: "flex",
    flexDirection: "row",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "8px 12px",
    borderBottom: "1px solid #ddd", 
  };

 
  const alertStyles: React.CSSProperties = {
    display: "flex",
    alignItems: "center",
    gap: "8px", 
  };

  
  const footerStyles: React.CSSProperties = {
    display: "flex",
    justifyContent: "flex-end", 
    padding: "12px 16px",
    borderTop: "1px solid #ddd", 
  };

 
  const modalStyles: IModalStyles = {
    main: {
      width: "100%",
      maxWidth: "290px", 
      "@media (min-width: 768px)": {
        maxWidth: "580px", 
      },
    },
    root: "",
    scrollableContent: "",
    layer: "",
    keyboardMoveIconContainer: "",
    keyboardMoveIcon: ""
  };

  return (
    <Modal
      isOpen={!hidden}
      isBlocking={true}
      onDismiss={handleDialogBox}
      styles={modalStyles} 
    >
      
      <div style={headerStyles}>
       
        <div style={alertStyles}>
        <IconButton iconProps={{ iconName: 'info' }} />
          <Text variant="large">Alert</Text>
        </div>

       
        <IconButton
          iconProps={closeIcon}
          ariaLabel="Close modal"
          onClick={handleDialogBox}
        />
      </div>

     
      <Stack tokens={{ padding: "16px" }} horizontalAlign="center" verticalAlign="center">
        <Text style={{ margin: "16px 0", fontSize: "14px", textAlign: "center" }}>
          The selected approver cannot be the same as existing Reviewers/ Approver/ Requester/ Current Actioner.
        </Text>
      </Stack>

     
      <div style={footerStyles}>
        <PrimaryButton
          text="OK"
          iconProps={okIcon} 
          onClick={handleDialogBox}
          ariaLabel="Confirm action"
        />
      </div>
    </Modal>
  );
};

export default ApproverOrReviewerModal;
