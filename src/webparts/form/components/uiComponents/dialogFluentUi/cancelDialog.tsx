import * as React from 'react';
import { Modal, PrimaryButton, DefaultButton, IconButton } from '@fluentui/react';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';


interface IConfirmationDialogProps {
  hidden: boolean;
  onConfirm: () => void; 
  onCancel: () => void; 

}

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
   minHeigth:'50px',
   padding:'5px'
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
    marginBottom: '14px',
    fontWeight:'400'
  },
  
});

const CancelConfirmationDialog: React.FC<IConfirmationDialogProps> = ({ hidden, onConfirm, onCancel, }) => {
  return (
    <Modal
      isOpen={hidden}
      onDismiss={onCancel}
      isBlocking={true}
      containerClassName={styles.modal}
    >
      <div className={styles.header}>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <IconButton iconProps={{ iconName: 'WaitlistConfirm' }} className={styles.headerIcon} />
          <h4 className={styles.headerTitle}>Confirmation</h4>
        </div>
        <IconButton iconProps={{ iconName: 'Cancel' }} onClick={onCancel} />
      </div>
      <div className={styles.body}>
        <p className={`${styles.removeTopMargin}`}>Are you sure you want to cancel this request?</p>
        <p className={`${styles.removeTopMargin}`}>Please click on the Confirm button to cancel the request.</p>
      </div>
      <div className={styles.footer}>
        <PrimaryButton
          onClick={onConfirm}
          text="Confirm"
          iconProps={{ iconName: 'SkypeCircleCheck', styles: { root: styles.buttonIcon } }}
          styles={{ root: styles.buttonContent }}
          className={styles.button}
        />
        <DefaultButton
          onClick={onCancel}
          text="Cancel"
          iconProps={{ iconName: 'ErrorBadge', styles: { root: styles.buttonIcon } }}
          styles={{ root: styles.buttonContent }}
          className={styles.button}
        />
      </div>
    </Modal>

    
  );
};

export default CancelConfirmationDialog;
