import * as React from 'react';
import {
    Dialog,
    DialogType,
    DialogFooter,
    PrimaryButton,
    DefaultButton,
    MessageBar,
    MessageBarType
} from '@fluentui/react';

interface ICustomPopupProps {
    isOpen: boolean;
    type: 'confirmation' | 'validation' | 'success' | 'error';
    title: string;
    message: string;
    onConfirm?: () => void;
    onCancel?: () => void;
    onClose?: () => void;
    onSuccessOk?: () => void;
}

const CustomPopup: React.FC<ICustomPopupProps> = (props) => {
    const {
        isOpen,
        type,
        title,
        message,
        onConfirm,
        onCancel,
        onClose,
        onSuccessOk
    } = props;

    const getDialogType = () => {
        switch (type) {
            case 'success':
                return DialogType.normal;
            case 'error':
                return DialogType.normal;
            case 'validation':
                return DialogType.normal;
            default:
                return DialogType.normal;
        }
    };

    const getMessageBarType = () => {
        switch (type) {
            case 'success':
                return MessageBarType.success;
            case 'error':
                return MessageBarType.error;
            case 'validation':
                return MessageBarType.warning;
            default:
                return MessageBarType.info;
        }
    };

    return (
        <Dialog
            hidden={!isOpen}
            onDismiss={onClose}
            dialogContentProps={{
                type: getDialogType(),
                title: title,
            }}
            modalProps={{
                isBlocking: true
            }}
        >
            <MessageBar messageBarType={getMessageBarType()}>
                {message}
            </MessageBar>
            <DialogFooter>
                {type === 'confirmation' && (
                    <>
                        <PrimaryButton onClick={onConfirm} text="Yes" />
                        <DefaultButton onClick={onCancel} text="No" />
                    </>
                )}
                {type === 'success' && (
                    <PrimaryButton onClick={onSuccessOk} text="OK" />
                )}
                {(type === 'validation' || type === 'error') && (
                    <PrimaryButton onClick={onClose} text="OK" />
                )}
            </DialogFooter>
        </Dialog>
    );
};

export default CustomPopup;