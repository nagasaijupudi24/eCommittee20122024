/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import React from 'react';
import { Dialog, DialogType, DialogFooter, PrimaryButton, DefaultButton } from '@fluentui/react';

const dialogStyles = { 
  main: { 
    maxWidth: '800px' 
  } 
};


interface AutoSaveDialogProps {
    hidden: boolean;
    onDismiss: () => void;
  }

const AutoSaveDialog:React.FC<AutoSaveDialogProps> = ({ hidden, onDismiss }) => {
  
  return (
    <Dialog
      hidden={hidden}
      onDismiss={onDismiss}
      dialogContentProps={{
        type: DialogType.largeHeader,
        title: 'Dialog Title',
        subText: 'This is the dialog content.'
      }}
      modalProps={{
        isBlocking: false,
        styles: dialogStyles
      }}
    >
      <DialogFooter>
        <PrimaryButton onClick={onDismiss} text="OK" />
        <DefaultButton onClick={onDismiss} text="Cancel" />
      </DialogFooter>
    </Dialog>
  );
};

export default AutoSaveDialog;
