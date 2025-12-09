import React from "react";
import { DefaultButton, IIconProps } from "@fluentui/react";
import './SharedGrid.css';

const addIcon: IIconProps = { iconName: "Add" };
const editIcon: IIconProps = { iconName: "Edit" };

interface LandingSectionProps {
  onNewPlacement: () => void;
  onUpdatePlacement: () => void;
}

const LandingSection: React.FC<LandingSectionProps> = ({ onNewPlacement, onUpdatePlacement }) => (
  <div className="ms-Grid" dir="ltr" id="maindiv">
    <div className="ms-Grid savesection">
      <div className="ms-Grid-row attachmentdiv">
        <div className="ms-Grid-col ms-sm3 ms-md2 ms-lg2 savebuttonmargin-placement">
          <DefaultButton
            iconProps={addIcon}
            text="New Placement"
            type="submit"
            onClick={onNewPlacement}
            styles={{
              root: {
                width: "100%",
                height: 40,
                justifyContent: "flex-start",
                paddingLeft: 4,
              },
              label: {
                textAlign: "left",
                width: "100%",
                fontWeight: "bold",
                fontSize: 12,
                color: "#242424",
              },
              icon: {
                backgroundColor: "#EBF3FC",
                padding: 7,
                borderRadius: 1,
                color: "#242424",
                fontWeight: "bold",
                fontSize: 10,
                width: 15,
                height: 15,
              },
            }}
          />
        </div>
        <br/>
        <div className="ms-Grid-col ms-sm3 ms-md2 ms-lg2 savebuttonmargin-placement">
          <DefaultButton
            iconProps={editIcon}
            text="Update Placement / Submission"
            type="submit"
            onClick={onUpdatePlacement}
            styles={{
              root: {
                width: "100%",
                height: 40,
                justifyContent: "flex-start",
                paddingLeft: 4,
              },
              label: {
                textAlign: "left",
                width: "100%",
                fontWeight: "bold",
                fontSize: 12,
                color: "#242424",
              },
              icon: {
                backgroundColor: "#EBF3FC",
                padding: 7,
                borderRadius: 1,
                color: "#242424",
                fontWeight: "bold",
                fontSize: 10,
                width: 15,
                height: 15,
              },
            }}
          />
        </div>
      </div>
    </div>
  </div>
);

export default LandingSection; 