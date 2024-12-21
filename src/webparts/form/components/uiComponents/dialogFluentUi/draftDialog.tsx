import * as React from 'react';
import { Dialog, DialogFooter, DialogType, DefaultButton, Icon } from '@fluentui/react';

interface IDraftSuccessDialogProps {
  hidden: boolean;
  onClose: () => void;
}

const DraftSuccessDialog: React.FC<IDraftSuccessDialogProps> = ({ hidden, onClose }) => {
  return (
    <Dialog
      hidden={hidden}
      onDismiss={onClose}
      dialogContentProps={{
        type: DialogType.normal,
        title: 'Alert',
        closeButtonAriaLabel: 'Close',
        subText: 'The request for eCommittee note has been drafted successfully.',
      }}
      modalProps={{
        isBlocking: false,
      }}
    >
      <Icon iconName="Warning" styles={{ root: { fontSize: '24px', color: 'red' } }} />

      <DialogFooter>
        <DefaultButton
          text="OK"
          onClick={onClose}
          iconProps={{ iconName: "ReplyMirrored" }}
        />
      </DialogFooter>
    </Dialog>
  );
};

export default DraftSuccessDialog;
