import React, { useState, useEffect, useRef } from "react";
import {
  Dropdown,
  IDropdownOption,
  IDropdownStyles,
  ResponsiveMode,
  Toggle,
  PrimaryButton,
  MessageBar,
  MessageBarType,
  ComboBox,
  IComboBox,
  IComboBoxOption
} from "@fluentui/react";
import './SharedGrid.css';
import PlacementApiSearchService from "../service/PlacementApiSearchService";
import AuthService from "../service/AuthService";

interface ExtendedComboBoxOption extends IComboBoxOption {
  insurerName?: string;
  brokerName?: string;
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
  isUpdateMode?: boolean;
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
  isUpdateMode,
}) => {
  // Check if submit button should be disabled
  const isSubmitDisabled = externalIsSubmitDisabled || (isDraftEmail && (!emailSubject || emailSubject.trim() === ''));
  const [placements, setPlacements] = useState<ExtendedComboBoxOption[]>([]);
  const [selectedPID, setSelectedPID] = useState<string | undefined>("");
  const comboRef = useRef<IComboBox>(null);

  const fetchPlacements = async (inputValue: string) => {
    if (inputValue.trim() === "") {
      setPlacements([]);
      return;
    }
    try {
      let apiTokenSearch = "";
      await AuthService.getApiToken().then(value => {
        apiTokenSearch = value?.accessToken || "";
      });
      const response = await PlacementApiSearchService.searchPlacementID(apiTokenSearch, {
        productCode: selectedProduct,
        searchString: inputValue,
      });
      
      const options = response.map((placement) => ({
        key: placement.placementId,
        text: placement.placementId,
        insurerName: placement.insuredName,
        brokerName: placement.broker,
      }));
      setPlacements(options);
    } catch (error) {
      console.error("Error fetching placements:", error);
    }
  };

 const handleInputValueChange = (text: string): void => {
    // Clear placements if the input is empty
    if (text.trim() === "") {
      setPlacements([]);
      return;
    }
 
    // Trigger API call only when the input length is 2 or more characters
    if (text.length >= 2) {
      fetchPlacements(text); // Fetch placements based on the input
    } else {
      setPlacements([]); // Clear placements when less than 2 characters
    }
    comboRef.current?.focus(true);
  };

  const onPIDChange = (_: any, option?: IComboBoxOption) => {
    setSelectedPID(option?.key?.toString());
  };

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
              {fileValidationError?.split('\n').map((line, index) => {
                // Style the main message differently
                if (index === 0) {
                  return (
                    <div key={index} style={{
                      marginBottom: '8px',
                      fontWeight: '600'
                    }}>
                      {line}
                    </div>
                  );
                }
                // Style bullet points
                return (
                  <div key={index} style={{
                    marginLeft: '16px',
                    marginBottom: '4px'
                  }}>
                    {line}
                  </div>
                );
              })}
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
              disabled={isUpdateMode}
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

        {/* Placement ID ComboBox */}
        {isUpdateMode && (
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-sm12 ms-md6 ms-lg6 paddingtopalldivs">
              <ComboBox
                label="Placement ID"
                componentRef={comboRef}
                placeholder="Search for Placement ID / Insured"
                selectedKey={selectedPID}
                onInputValueChange={handleInputValueChange}
                allowFreeform
                autoComplete="on"
                options={placements}
                styles={{
                  root: { width: "100%", position: "relative", backgroundColor: "transparent" },
                  input: { paddingRight: 60 },
                  callout: { minWidth: "100%" },
                }}
                onChange={onPIDChange}
                onRenderOption={(option) => {
                  const opt = option as ExtendedComboBoxOption;
                  return (
                    <div style={{ lineHeight: "1.4", padding: "4px 8px" }}>
                      <div style={{ fontWeight: 600, color: "#323130" }}>{opt.text}</div>
                      {opt.insurerName && (
                        <div style={{ fontSize: 12, color: "#605E5C" }}>
                          {opt.insurerName}
                        </div>
                      )}
                      {opt.brokerName && (
                        <div style={{ fontSize: 12, color: "#605E5C" }}>
                          {opt.brokerName}
                        </div>
                      )}
                    </div>
                  );
                }}
              />
            </div>
          </div>
        )}

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