/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from 'react';
import { Modal, PrimaryButton, IconButton } from '@fluentui/react';
import { mergeStyleSets } from '@fluentui/react/lib/Styling';

const GistDocSubmitted: React.FC<{ isVisibleAlter: boolean; onCloseAlter: () => void; statusOfReq: any ,existUrl:any}> = ({ isVisibleAlter, onCloseAlter, statusOfReq ,existUrl}) => {
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
      minHeight: "50px",
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
    },
    footer: {
      display: 'flex',
      justifyContent: 'flex-end',
      marginTop: '20px',
      borderTop: '1px solid #ddd', 
      paddingTop: '10px',
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
        <p>The request has been updated successfully</p>
      </div>
      <div className={styles.footer}>
        <PrimaryButton iconProps={{ iconName: 'ReplyMirrored' }} onClick={()=>{
          onCloseAlter()
          const pageURL: string =existUrl;
          window.location.href = `${pageURL}`;
          
        }} text="OK" />
      </div>
    </Modal>
  );
};

export default GistDocSubmitted;