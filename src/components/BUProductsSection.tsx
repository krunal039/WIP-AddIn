import React from "react";
import {
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  ResponsiveMode,
  Toggle,
  PrimaryButton,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import './SharedGrid.css';

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: {
    width: "100%",
    boxShadow: "0 0px 4px rgba(0, 0, 0, 0.2)",
    borderRadius: 4,
    backgroundColor: "transparent",
    borderBottom: "3px #242424",
  },
};

interface BUProductsSectionProps {
  selectedProduct: string;
  selectedBU: string;
  optionsProducts: IDropdownOption[];
  optionsBU: IDropdownOption[];
  onProductChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  onBUChange: (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => void;
  sendCopyToCyberAdmin: boolean;
  onSendCopyToggle: (event: React.MouseEvent<HTMLElement>, checked?: boolean) => void;
  onSubmit: () => void;
  emailSubject: string;
  isDraftEmail: boolean;
  fileValidationError?: string | null;
  isSubmitDisabled?: boolean;
}

const BUProductsSection: React.FC<BUProductsSectionProps> = ({
  selectedProduct,
  selectedBU,
  optionsProducts,
  optionsBU,
  onProductChange,
  onBUChange,
  sendCopyToCyberAdmin,
  onSendCopyToggle,
  onSubmit,
  emailSubject,
  isDraftEmail,
  fileValidationError,
  isSubmitDisabled: externalIsSubmitDisabled,
}) => {
  // Check if submit button should be disabled
  const isSubmitDisabled = externalIsSubmitDisabled || (isDraftEmail && (!emailSubject || emailSubject.trim() === ''));
  
  return (
    <div className="ms-Grid" dir="ltr" id="maindiv">
      {/* File Validation Error Message */}
      {fileValidationError && (
        <div className="ms-Grid-row" style={{ marginBottom: '16px' }}>
          <div className="ms-Grid-col ms-sm12">
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={true}
              styles={{
                root: {
                  backgroundColor: '#fdf2f2',
                  border: '1px solid #d13438',
                  borderRadius: '4px',
                },
                content: {
                  display: 'flex',
                  alignItems: 'flex-start',
                },
                icon: {
                  color: '#d13438',
                },
                text: {
                  color: '#d13438',
                  fontWeight: '500',
                }
              }}
            >
              {fileValidationError?.split('\n').map((line, index) => (
                <div key={index} style={{ marginBottom: index > 0 ? '4px' : '0' }}>
                  {line}
                </div>
              ))}
            </MessageBar>
          </div>
        </div>
      )}
      
      <div className="ms-Grid savesection">
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6 paddingtopalldivs">
            <Dropdown
              label="Products"
              selectedKey={selectedProduct}
              defaultSelectedKey="20001"
              options={optionsProducts}
              styles={dropdownStyles}
              responsiveMode={ResponsiveMode.large}
              onChange={onProductChange}
              calloutProps={{
                directionalHint: 4,
                isBeakVisible: false,
                doNotLayer: true,
              }}
            />
          </div>
        </div>
        <div className="ms-Grid-row">
          <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6 paddingtopalldivs">
            <Dropdown
              label="Business Unit"
              selectedKey={selectedBU}
              options={optionsBU}
              disabled={true}
              styles={{
                ...dropdownStyles,
                dropdown: {
                  backgroundColor: "transparent",
                  cursor: "default",
                },
                label: { color: "#000000" },
                title: {
                  backgroundColor: "transparent",
                  color: "#323130",
                  border: "1px solid #8A8886",
                },
                caretDown: { color: "#605E5C" },
              }}
              calloutProps={{
                directionalHint: 4,
                isBeakVisible: false,
                doNotLayer: true,
              }}
              onChange={onBUChange}
            />
          </div>
        </div>
        {selectedProduct === "20001" && (
          <>
            <h3 style={{ marginBottom: 10, padding: "5px 4px" }}>
              Shared Mailbox
            </h3>
            <div style={{ marginTop: 0, marginBottom: 20 }}>
              <Toggle
                inlineLabel
                label="Send a copy to MRSL457-CyberAdmin-(Pool)"
                checked={sendCopyToCyberAdmin}
                onChange={onSendCopyToggle}
                styles={{
                  root: { marginLeft: 2 },
                  label: { fontWeight: "semibold", color: "#242424" },
                }}
              />
            </div>
          </>
        )}
        {isDraftEmail && isSubmitDisabled && (
          <div style={{ 
            marginBottom: 10, 
            padding: "8px 12px", 
            color: "#d13438", 
            fontSize: "13px",
            backgroundColor: "#fde7e9",
            border: "1px solid #d13438",
            borderRadius: "4px",
            fontWeight: "500"
          }}>
            ⚠️ Please add a subject to your email before submitting
          </div>
        )}
        <div className="ms-Grid-row attachmentdiv">
          <div className="ms-Grid-col ms-sm3 ms-md2 ms-lg2 savebuttonmargin">
            <PrimaryButton
              className="bottomLeftButton"
              text="Submit"
              type="submit"
              onClick={onSubmit}
              disabled={isSubmitDisabled}
              styles={{
                root: {
                  width: 100,
                  height: 40,
                  fontWeight: "bold",
                  backgroundColor: isSubmitDisabled ? "#8A8886" : "#0F1E32",
                  borderRadius: 4,
                  opacity: isSubmitDisabled ? 0.6 : 1,
                  cursor: isSubmitDisabled ? "not-allowed" : "pointer",
                },
                rootHovered: { 
                  backgroundColor: isSubmitDisabled ? "#8A8886" : "#0F1E32",
                  opacity: isSubmitDisabled ? 0.6 : 1,
                },
                rootPressed: { 
                  backgroundColor: isSubmitDisabled ? "#8A8886" : "#0F1E32",
                  opacity: isSubmitDisabled ? 0.6 : 1,
                },
              }}
            />
          </div>
        </div>
      </div>
    </div>
  );
};

export default BUProductsSection; 