import React from 'react';
import { MessageBar, MessageBarType, PrimaryButton, DefaultButton } from '@fluentui/react';

interface ConfirmationDialogProps {
  isVisible: boolean;
  onSendAgain: () => void;
  onCancel: () => void;
}

const ConfirmationDialogComponent: React.FC<ConfirmationDialogProps> = ({
  isVisible,
  onSendAgain,
  onCancel
}) => {
  if (!isVisible) return null;

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(0, 0, 0, 0.5)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 1000
    }}>
      <div style={{
        backgroundColor: 'white',
        padding: '20px',
        borderRadius: '0px',
        maxWidth: '400px',
        width: '80%'
      }}>
        <h3 style={{ margin: '0 0 16px 0' }}>This email has already been sent for ingestion</h3>
        <p style={{ margin: '0 0 20px 0' }}>Are you sure you want to send this email again?</p>
        <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end' }}>
          <DefaultButton text="Cancel" onClick={onCancel} />
          <PrimaryButton text="Send Again" onClick={onSendAgain} />
        </div>
      </div>
    </div>
  );
};

export const ConfirmationDialog = React.memo(ConfirmationDialogComponent);
ConfirmationDialog.displayName = 'ConfirmationDialog';

interface RetryButtonProps {
  isVisible: boolean;
  onRetry: () => void;
  reason?: string;
  hasValidData?: boolean;
}

const RetryButtonComponent: React.FC<RetryButtonProps> = ({ isVisible, onRetry, reason, hasValidData = true }) => {
  if (!isVisible) return null;

  const getButtonText = () => {
    if (reason === 'DRAFT_EMAIL_NO_ITEM_ID') {
      return "Email forwarding not available for draft emails. Please save or send the email first.";
    }
    if (!hasValidData) {
      return "Retry not available - missing email data";
    }
    return "Retry Send to Shared Mailbox";
  };

  const isDisabled = reason === 'DRAFT_EMAIL_NO_ITEM_ID' || !hasValidData;

  return (
    <div>
      <MessageBar
        messageBarType={MessageBarType.warning}
        isMultiline={true}
        styles={{
          root: {
            backgroundColor: "#FFF3CD",
            border: "1px solid #FFEAA7",
            color: "#856404",
            marginBottom: 16,
          },
        }}
      >
        {reason === 'DRAFT_EMAIL_NO_ITEM_ID' 
          ? "Email forwarding requires the email to be saved or sent first. The placement was successful, but forwarding to the shared mailbox is not available for draft emails."
          : !hasValidData
          ? "Email forwarding failed, but placement was successful. Retry is not available because the email data is not accessible."
          : "Email forwarding failed, but placement was successful. You can retry forwarding to the shared mailbox."
        }
      </MessageBar>
      <PrimaryButton
        text={getButtonText()}
        onClick={onRetry}
        disabled={isDisabled}
        styles={{ 
          root: { 
            marginTop: 16,
            backgroundColor: isDisabled ? "#C8C6C4" : "#0F1E32",
          },
          rootHovered: { 
            backgroundColor: isDisabled ? "#C8C6C4" : "#0F1E32" 
          },
          rootPressed: { 
            backgroundColor: isDisabled ? "#C8C6C4" : "#0F1E32" 
          },
        }}
      />
    </div>
  );
};

export const RetryButton = React.memo(RetryButtonComponent);
RetryButton.displayName = 'RetryButton';

interface SuccessMessageProps {
  isVisible: boolean;
  onSuccess?: () => void;
}

const SuccessMessageComponent: React.FC<SuccessMessageProps> = ({ isVisible, onSuccess }) => {
  if (!isVisible) return null;

  return (
    <div>
      <header style={{
        display: "flex",
        alignItems: "center",
        padding: "0px 0px",
        backgroundColor: "transparent",
      }}>
        <h3 style={{ margin: 0, padding: "20px 4px", display: "none" }}>
          Underwriting Workbench
        </h3>
      </header>
      <MessageBar
        messageBarType={MessageBarType.success}
        isMultiline={true}
        styles={{
          root: {
            backgroundColor: "#E6F4EA",
            border: "1px solid #A6D8A8",
            color: "#0F5132",
            fontWeight: "normal",
            marginTop: 20,
            height: 200,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            textAlign: "left",
          },
          icon: {
            color: "#198754",
            fontSize: 28,
            height: 28,
            width: 28,
            alignSelf: "center",
          },
        }}
      >
        <div>
          <strong>Action Completed</strong>
          <div>Your action was completed successfully.</div>
        </div>
      </MessageBar>
    </div>
  );
};

export const SuccessMessage = React.memo(SuccessMessageComponent);
SuccessMessage.displayName = 'SuccessMessage';

interface ErrorMessageProps {
  isVisible: boolean;
  message?: string;
  onSubmit?: () => void;
}

const ErrorMessageComponent: React.FC<ErrorMessageProps> = ({ isVisible, message, onSubmit }) => {
  if (!isVisible) return null;

  return (
    <div>
      <div>
        <header style={{
          display: "flex",
          alignItems: "center",
          padding: "0px 0px",
          backgroundColor: "transparent",
        }}>
          <h3 style={{ margin: 0, padding: "20px 4px", display: "none" }}>
            Underwriting Workbench
          </h3>
        </header>
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={true}
          styles={{
            root: {
              backgroundColor: "#FBE9E7",
              border: "1px solid #F4F4F4",
              color: "#0F5132",
              fontWeight: "normal",
              marginTop: 20,
              height: 200,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              textAlign: "left",
            },
            icon: {
              color: "#D7422F",
              fontSize: 28,
              height: 28,
              width: 28,
              alignSelf: "center",
            },
          }}
        >
          <div>
            <strong>Something went wrong...</strong>
            <div>
              {message || "An error occurred while processing your request. Please try again."}
            </div>
          </div>
        </MessageBar>
      </div>
      <div className="ms-Grid-row attachmentdiv">
        <div className="ms-Grid-col ms-sm3 ms-md2 ms-lg2 savebuttonmargin">
          <PrimaryButton
            className="bottomLeftButton"
            text="Retry"
            type="submit"
            onClick={onSubmit}
            styles={{
              root: {
                width: 100,
                height: 40,
                fontWeight: "bold",
                backgroundColor: "#0F1E32",
                borderRadius: 4,
                opacity: 1,
                cursor: "pointer",
              },
              rootHovered: {
                backgroundColor: "#0F1E32",
                opacity: 1,
              },
              rootPressed: {
                backgroundColor: "#0F1E32",
                opacity: 1,
              },
            }}
          />
        </div>
      </div>
    </div>
  );
};

export const ErrorMessage = React.memo(ErrorMessageComponent);
ErrorMessage.displayName = 'ErrorMessage';
