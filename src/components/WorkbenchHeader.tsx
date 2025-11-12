import React, { memo } from "react";
import { IconButton, IIconProps } from "@fluentui/react";

const backIcon: IIconProps = { iconName: "Back" };
const homeIcon: IIconProps = { iconName: "Home" };

interface WorkbenchHeaderProps {
  onBack: () => void;
  onHome: () => void;
  title: string;
  isDuplicate: boolean;
}

const WorkbenchHeaderComponent: React.FC<WorkbenchHeaderProps> = ({ onBack, onHome, title, isDuplicate }) => (
  <div>
    <header style={{ display: "flex", alignItems: "center", padding: "0px" }}>
      <IconButton
        iconProps={backIcon}
        title="Back"
        ariaLabel="Back"
        onClick={onBack}
        styles={{ root: { marginRight: 4 } }}
      />
      <IconButton
        iconProps={homeIcon}
        title="Home"
        ariaLabel="Home"
        onClick={onHome}
        styles={{ root: { marginRight: 8 } }}
      />
      <h3 style={{ margin: 0, padding: "20px 4px" }}>{title}</h3>
    </header>
    {isDuplicate && (      
        <div style={{ 
          paddingLeft: "20px", 
          fontSize: "14px", 
          fontWeight: "bold", 
          marginLeft: "20px", 
          marginBottom: "10px"
        }}>
          Sent for ingestion
        </div>
    )}
  </div>
);

const WorkbenchHeader = memo(WorkbenchHeaderComponent);
WorkbenchHeader.displayName = 'WorkbenchHeader';

export default WorkbenchHeader; 