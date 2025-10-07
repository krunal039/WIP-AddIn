import React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react";

interface SpinnerOverlayProps {
  label?: string;
  isBlocking?: boolean;
  backgroundColor?: string;
}

const SpinnerOverlay: React.FC<SpinnerOverlayProps> = ({
  label = "Loading...",
  isBlocking = true,
  backgroundColor = "rgba(255, 255, 255, 0.6)",
}) => {
  const style: React.CSSProperties = {
    position: "fixed",
    top: 0,
    left: 0,
    width: "100vw",
    height: "100vh",
    backgroundColor,
    zIndex: 9999,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    pointerEvents: isBlocking ? "auto" : "none",
  };

  return (
    <div style={style}>
      <Spinner label={label} size={SpinnerSize.large} />
    </div>
  );
};

export default SpinnerOverlay;
