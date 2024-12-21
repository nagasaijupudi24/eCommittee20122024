/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { Modal, PrimaryButton, IconButton } from '@fluentui/react';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';

const RejectBtnCommentCheckDialog: React.FC<{ isVisibleAlter: boolean; onCloseAlter: () => void; statusOfReq: any }> = ({ isVisibleAlter, onCloseAlter, statusOfReq }) => {
  const styles = mergeStyleSets({
    modal: {
      padding: '10px',
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
      justifyContent: 'flex-end',
     
      borderTop: '1px solid #ddd', 
      paddingTop: '10px',
       minHeight:'50px'
    },
    button: {
     
      maxHeight:'32px',
     
    },
  });

  return (
    <Modal
      isOpen={isVisibleAlter}
      onDismiss={onCloseAlter}
      isBlocking={true}
      containerClassName={styles.modal}
    >
      <div className={styles.header}>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <IconButton iconProps={{ iconName: 'Info' }} />
          <h4 className={styles.headerTitle}>Alert</h4>
        </div>
        <IconButton iconProps={{ iconName: 'Cancel' }} onClick={onCloseAlter} />
      </div>
      <div className={styles.body}>
        <p>Please fill in comments then click on Reject.</p>
      </div>
      <div className={styles.footer}>
        <PrimaryButton className={styles.button}  iconProps={{ iconName: 'ReplyMirrored' }} onClick={onCloseAlter} text="OK" />
      </div>
    </Modal>
  );
};

export default RejectBtnCommentCheckDialog;
