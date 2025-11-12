import React, { useState, useMemo, useCallback } from "react";
import { useLocalStorage } from "../hooks/useLocalStorage";
import {
  DatePicker,
  DefaultButton,
  IconButton,
  defaultDatePickerStrings,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  IIconProps,
  ResponsiveMode,
  Toggle,
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";

import SpinnerOverlay from "./SpinnerOverlay";
import { EmailConverterService } from "../service/EmailConverterService";
import { WorkbenchService } from "../service/WorkbenchService";
import {
  ConfirmationDialog,
  RetryButton,
  SuccessMessage,
  ErrorMessage,
} from "./WorkbenchDialogs";
import LoggingService from "../service/LoggingService";
import DebugService from "../service/DebugService";
import OfficeModeService from "../service/OfficeModeService";
import { checkDuplicateSubmission } from "../utils/duplicateDetection";
// Removed placementSubmission wrapper - calling service directly
import FileValidationService, { FileValidationResult } from "../service/FileValidationService";
import LandingSection from "./LandingSection";
import BUProductsSection from "./BUProductsSection";
import WorkbenchHeader from "./WorkbenchHeader";
import { BU_OPTIONS, PRODUCT_OPTIONS, DEFAULTS, STORAGE_KEYS } from "../constants";

export interface WorkbenchLandingProps {
  apiToken: string | null;
  graphToken: string | null;
}

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: {
    width: "100%",
    boxShadow: "0 0px 4px rgba(0, 0, 0, 0.2)",
    borderRadius: 4,
    backgroundColor: "transparent",
    borderBottom: "3px #242424",
  },
};

const optionsBU: IDropdownOption[] = BU_OPTIONS;
const optionsProducts: IDropdownOption[] = PRODUCT_OPTIONS;

const addIcon: IIconProps = { iconName: "Add" };
const backIcon: IIconProps = { iconName: "Back" };

const WorkbenchLanding: React.FC<WorkbenchLandingProps> = ({
  apiToken,
  graphToken,
}) => {
  const workbenchService = WorkbenchService.getInstance();

  // State
  const [showLanding, setShowLanding] = useState(true);
  const [showBUProducts, setShowBUProducts] = useState(false);
  const [selectedProduct, setSelectedProduct] = useLocalStorage(STORAGE_KEYS.SELECTED_PRODUCT, DEFAULTS.PRODUCT);
  const [selectedBU, setSelectedBU] = useLocalStorage(STORAGE_KEYS.SELECTED_BU, DEFAULTS.BU);
  const [sendCopyToCyberAdmin, setSendCopyToCyberAdmin] = useLocalStorage<boolean>(STORAGE_KEYS.SEND_COPY_TO_CYBER_ADMIN, DEFAULTS.SEND_COPY_TO_CYBER_ADMIN);
  const [showSuccessMessage, setShowSuccessMessage] = useState(false);
  const [showFailureMessage, setShowFailureMessage] = useState(false);
  const [showLoadingMessage, setShowLoadingMessage] = useState(false);
  const [forwardingFailed, setForwardingFailed] = useState(false);
  const [forwardingFailedReason, setForwardingFailedReason] = useState<
    string | undefined
  >(undefined);
  const [lastPlacementId, setLastPlacementId] = useState<string | undefined>(
    undefined
  );
  const [lastGraphItemId, setLastGraphItemId] = useState<string | undefined>(
    undefined
  );
  const [lastSharedMailbox, setLastSharedMailbox] = useState<
    string | undefined
  >(undefined);
  const [showConfirmationDialog, setShowConfirmationDialog] = useState(false);
  const [emailSubject, setEmailSubject] = useState<string>("");
  const [isDraftEmail, setIsDraftEmail] = useState<boolean>(false);
  const [isDuplicate, setIsDuplicate] = useState<boolean>(false);
  const [fileValidationResult, setFileValidationResult] = useState<FileValidationResult>({ isValid: true, errors: [] });
  const [isValidatingFiles, setIsValidatingFiles] = useState<boolean>(false);

  const handleDownloadEmail = async () => {
    try {
      await Office.onReady();
    } catch (err) {
      DebugService.error("Office.js failed to initialize:", err);
      setShowFailureMessage(true);
      return;
    }
    let item;
    try {
      item = Office.context.mailbox.item;
    } catch (err) {
      DebugService.error("Unable to access mailbox item:", err);
      setShowFailureMessage(true);
      return;
    }
    if (!item) {
      DebugService.error("No email item available");
      return;
    }

    // Parallelize independent operations: duplicate check, subject extraction, and file validation
    const isDraft = OfficeModeService.isComposeMode();
    setIsDraftEmail(isDraft);

    try {
      const [isDuplicateN, subject, filesValid] = await Promise.all([
        checkDuplicateSubmission(item),
        new Promise<string>((resolve) => {
          (item as any).subject.getAsync((result: any) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value || "");
            } else {
              resolve("");
            }
          });
        }).catch(() => ""),
        validateEmailFiles(item)
      ]);
      
      setIsDuplicate(isDuplicateN);
      setEmailSubject(subject);
      DebugService.debug("Is duplicate Check Done");
      
      if (!filesValid) {
        DebugService.warn("File validation failed - showing error message");
        // Don't return here - let the UI show the error message
      }
    } catch (err) {
      DebugService.error("Error in parallel operations:", err);
      // Set defaults on error
      setIsDuplicate(false);
      setEmailSubject("");
    }

    // Trigger early save when user first interacts with UI
    try {
      await workbenchService.attemptEarlySave(item);
    } catch (error) {
      DebugService.warn("Early save failed, continuing with workflow:", error);
    }

    setShowLanding(false);
    setShowBUProducts(true);
  };

  // Adds a banner to the top of the email body
  const addBannerToEmail = async (): Promise<void> => {
    await Office.onReady();

    const item = Office.context.mailbox.item;
    if (!item) {
      DebugService.error("No email item available for banner");
      return;
    }

    // Check if item is a message
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      const messageCompose = item as Office.MessageCompose;
      
      // Check if we're in compose mode by checking if prependAsync is available
      if (messageCompose.body && typeof messageCompose.body.prependAsync === 'function') {
        const bannerHtml = `<div style="padding:10px;font-weight:bold;color:#00796b;margin-bottom:10px;">Sent for Ingestion</div>`;

        return new Promise<void>((resolve, reject) => {
          messageCompose.body.prependAsync(
            bannerHtml,
            { coercionType: Office.CoercionType.Html },
            (result: Office.AsyncResult<void>) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                DebugService.debug("Banner added successfully.");
                resolve();
              } else {
                DebugService.error("Error adding banner:", result.error.message);
                reject(result.error);
              }
            }
          );
        });
      } else {
        // We're in read mode, can't add banner
        DebugService.debug("Cannot add banner in read mode - item is not in compose mode.");
        return Promise.resolve(); // Don't reject, just skip adding the banner
      }
    } else {
      DebugService.error("Not a message item.");
      return Promise.reject("Invalid context - not a message item.");
    }
  };

  // Cast to MessageCompose to access body
  //   const message = item as Office.MessageCompose;

  //   if (message.body && typeof message.body.prependAsync === "function") {
  //     return new Promise<void>((resolve, reject) => {
  //       const bannerHtml = `<div style="background:#e0f7fa;padding:10px;font-weight:bold;color:#00796b;border:1px solid #00796b;margin-bottom:10px;">Sent for Ingestion</div>`;
  //       try {
  //         message.body.prependAsync(
  //           bannerHtml,
  //           { coercionType: Office.CoercionType.Html },
  //           (result: Office.AsyncResult<void>) => {
  //             if (result.status === Office.AsyncResultStatus.Succeeded) {
  //               resolve();
  //               console.log("Banner added successfully")
  //             } else {
  //               reject(result.error);
  //             }
  //           }
  //         );
  //       } catch (err) {
  //         reject(err as any);
  //       }
  //     });
  //   } else {
  //     console.error("Message body is not available in this context.");
  //     return Promise.reject("Message body is not available.");
  //   }
  // };

  const handleSendCopyToggle = (
    _event: React.MouseEvent<HTMLElement>,
    checked?: boolean
  ) => {
    setSendCopyToCyberAdmin(!!checked);
  };

  const validateEmailFiles = async (item: any): Promise<boolean> => {
    try {
      setIsValidatingFiles(true);
      const validationResult = await FileValidationService.validateEmailAttachments(item);
      setFileValidationResult(validationResult);
      return validationResult.isValid;
    } catch (error) {
      DebugService.error('File validation failed:', error);
      setFileValidationResult({
        isValid: false,
        errors: [{
          type: 'unsupported',
          message: 'Unable to validate attachments. Please check file types and try again.',
          files: []
        }]
      });
      return false;
    } finally {
      setIsValidatingFiles(false);
    }
  };

  const handleSubmit = async () => {
    setShowLoadingMessage(true);
    try {
      await Office.onReady();
      const item = Office.context.mailbox.item;
      if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message) {
        setShowLoadingMessage(false);
        setShowFailureMessage(true);
        return;
      }

      // For draft emails, check if subject is empty before proceeding
      if (OfficeModeService.isComposeMode()) {
        try {
          const currentSubject = await new Promise<string>((resolve) => {
            (item as any).subject.getAsync((result: any) => {
              if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value || "");
              } else {
                resolve("");
              }
            });
          });

          if (!currentSubject || currentSubject.trim() === "") {
            setShowLoadingMessage(false);
            setShowFailureMessage(true);
            return;
          }
        } catch (error) {
          DebugService.warn("Failed to get current email subject:", error);
          setShowLoadingMessage(false);
          setShowFailureMessage(true);
          return;
        }
      }

      // Validate files and check for duplicates in parallel (both are independent operations)
      const [filesValid, isDuplicate] = await Promise.all([
        validateEmailFiles(item),
        checkDuplicateSubmission(item as any)
      ]);
      
      DebugService.debug("Duplicate detection result:", isDuplicate);
      
      if (!filesValid) {
        setShowLoadingMessage(false);
        return;
      }
      
      if (isDuplicate) {
        DebugService.debug("Duplicate detected - showing confirmation dialog");
        setShowLoadingMessage(false);
        setShowConfirmationDialog(true);
        return;
      }
      const result = await workbenchService.submitPlacement(
        apiToken!,
        graphToken!,
        item,
        selectedProduct,
        sendCopyToCyberAdmin
      );
      if (result.success) {
        if (result.forwardingFailed) {
          setShowLoadingMessage(false);
          setShowFailureMessage(true);
          setForwardingFailed(true);
          setLastPlacementId(result.lastPlacementId);
          setLastGraphItemId(result.lastGraphItemId);
          setLastSharedMailbox(result.lastSharedMailbox);
        } else {
          setShowLoadingMessage(false);
          setShowSuccessMessage(true);
        }
      } else {
        setShowLoadingMessage(false);
        setShowFailureMessage(true);
      }
    } catch (error) {
      setShowLoadingMessage(false);
      setShowFailureMessage(true);
    }
  };

  const handleSendAgain = async () => {
    setShowConfirmationDialog(false);
    setShowLoadingMessage(true);
    try {
      const item = Office.context.mailbox.item;
      if (item) {
        const result = await workbenchService.submitPlacement(
          apiToken!,
          graphToken!,
          item,
          selectedProduct,
          sendCopyToCyberAdmin
        );
        if (result.success) {
          if (result.forwardingFailed) {
            setShowLoadingMessage(false);
            setShowFailureMessage(true);
            setForwardingFailed(true);
            setForwardingFailedReason(result.forwardingFailedReason);
            setLastPlacementId(result.lastPlacementId);
            setLastGraphItemId(result.lastGraphItemId);
            setLastSharedMailbox(result.lastSharedMailbox);
          } else {
            setShowLoadingMessage(false);
            setShowSuccessMessage(true);
          }
        } else {
          setShowLoadingMessage(false);
          setShowFailureMessage(true);
        }
      }
    } catch (error) {
      setShowLoadingMessage(false);
      setShowFailureMessage(true);
    }
  };

  const handleCancel = () => {
    setShowConfirmationDialog(false);
    setShowLanding(true);
    setShowBUProducts(false);
  };

  const handleRetryForward = async () => {
    DebugService.debug("=== Retry Forward Debug ===");
    DebugService.debug("lastPlacementId:", lastPlacementId);
    DebugService.debug("lastGraphItemId:", lastGraphItemId);
    DebugService.debug("lastSharedMailbox:", lastSharedMailbox);
    DebugService.debug("forwardingFailedReason:", forwardingFailedReason);
    DebugService.debug("=== End Retry Forward Debug ===");
    if (!lastPlacementId || !lastGraphItemId || !lastSharedMailbox) {
      DebugService.error("Missing required data for retry:", {
        lastPlacementId,
        lastGraphItemId,
        lastSharedMailbox,
      });
      return;
    }
    if (forwardingFailedReason === "DRAFT_EMAIL_NO_ITEM_ID") {
      DebugService.debug(
        "Cannot retry forwarding for draft email - itemId not available"
      );
      return;
    }
    setShowLoadingMessage(true);
    try {
      const result = await workbenchService.retryForward(
        graphToken!,
        lastPlacementId,
        lastGraphItemId,
        lastSharedMailbox
      );
      if (result.success) {
        setShowLoadingMessage(false);
        setShowSuccessMessage(true);
        setShowFailureMessage(false);
        setForwardingFailed(false);
      } else {
        setShowLoadingMessage(false);
        setShowFailureMessage(true);
      }
    } catch (error) {
      setShowLoadingMessage(false);
      setShowFailureMessage(true);
    }
  };

  const handleLandingSave = async () => {
    DebugService.debug('handleLandingSave called - starting file validation');
    
    // Validate email attachments when user clicks "New Placement"
    try {
      const item = Office.context.mailbox.item;
      if (item) {
        DebugService.debug('Office item found, calling validateEmailFiles');
        await validateEmailFiles(item);
        DebugService.debug('File validation completed in handleLandingSave');
      } else {
        DebugService.warn('No Office item found in handleLandingSave');
      }
    } catch (error) {
      DebugService.error('File validation failed in handleLandingSave:', error);
    }
    
    setShowLanding(false);
    setShowBUProducts(true);
  };

  const handleBack = () => {
    // Always clear overlays
    setShowSuccessMessage(false);
    setShowFailureMessage(false);
    setShowLoadingMessage(false);
    setShowConfirmationDialog(false);
    // Go to previous logical screen
    setShowLanding(false);
    setShowBUProducts(true);
  };

  const handleHome = () => {
    // Always clear overlays
    addBannerToEmail();
    setShowSuccessMessage(false);
    setShowFailureMessage(false);
    setShowLoadingMessage(false);
    setShowConfirmationDialog(false);
    // Go to landing screen
    setShowLanding(true);
    setShowBUProducts(false);
    // Optionally reset other state as needed
  };

  // Memoize expensive computations
  const fileValidationErrorMessage = useMemo(() => 
    FileValidationService.getAllErrorMessages(fileValidationResult.errors),
    [fileValidationResult.errors]
  );

  const isSubmitDisabled = useMemo(() => 
    !fileValidationResult.isValid || isValidatingFiles,
    [fileValidationResult.isValid, isValidatingFiles]
  );

  const handleProductChange = useCallback((
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) => {
    const selectedProduct = option?.key as string;
    let selectedBU = "MRSNA";
    if (selectedProduct === "20001") {
      selectedBU = "MRSGM";
    } else if (selectedProduct === "NA_LPL" || selectedProduct === "NA_MPL") {
      selectedBU = "MRSNA";
    }
    if (selectedProduct) setSelectedProduct(selectedProduct as any);
    if (selectedBU) setSelectedBU(selectedBU as any);
  }, [setSelectedProduct, setSelectedBU]);

  const handleBUChange = useCallback((
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ) => {
    if (option?.key) setSelectedBU(option.key as any);
  }, [setSelectedBU]);

  // Render
  // Security check - ensure we have valid tokens
  if (!apiToken || !graphToken) {
    return (
      <div style={{ padding: "16px" }}>
        <MessageBar messageBarType={MessageBarType.error} isMultiline={true}>
          Authentication required. Please reload the add-in and sign in again.
        </MessageBar>
      </div>
    );
  }

  return (
    <div>
      <WorkbenchHeader
        onBack={handleBack}
        onHome={handleHome}
        title="Underwriting Workbench"
        isDuplicate={isDuplicate}
      />
      {showLoadingMessage ? (
        <SpinnerOverlay label="Saving to workbench..." />
      ) : showSuccessMessage ? (
        <SuccessMessage isVisible={true} onSuccess={addBannerToEmail} />
      ) : showFailureMessage ? (
        <div>
          <ErrorMessage isVisible={true} onSubmit={handleSubmit} />
          <RetryButton
            isVisible={forwardingFailed}
            onRetry={handleRetryForward}
            reason={forwardingFailedReason}
            hasValidData={!!lastGraphItemId}
          />
        </div>
      ) : (
        <>
          {showLanding && <LandingSection onNewPlacement={handleLandingSave} />}
          {showBUProducts && (
            <BUProductsSection
              selectedProduct={selectedProduct}
              selectedBU={selectedBU}
              optionsProducts={optionsProducts}
              optionsBU={optionsBU}
              onProductChange={handleProductChange}
              onBUChange={handleBUChange}
              sendCopyToCyberAdmin={sendCopyToCyberAdmin}
              onSendCopyToggle={handleSendCopyToggle}
              onSubmit={handleSubmit}
              emailSubject={emailSubject}
              isDraftEmail={isDraftEmail}
                    fileValidationError={fileValidationErrorMessage}
                    isSubmitDisabled={isSubmitDisabled}
            />
          )}
          <ConfirmationDialog
            isVisible={showConfirmationDialog}
            onSendAgain={handleSendAgain}
            onCancel={handleCancel}
          />
        </>
      )}
    </div>
  );
};

export default WorkbenchLanding;
