import * as React from "react";
import { Panel, PanelType } from "@fluentui/react";
import styles from "./FloorMapPanel.module.scss";

interface Props {
  isOpen: boolean;
  onDismiss: () => void;
  imageUrl: string; // Floor map image URL
}

const FloorMapPanel: React.FC<Props> = ({ isOpen, onDismiss, imageUrl }) => {
  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.large}
      headerText="Office Floor Map"
      closeButtonAriaLabel="Close"
    >
      <div className={styles.wrapper}>
        <img
          src={imageUrl}
          alt="Floor Map"
          className={styles.mapImage}
        />
      </div>
    </Panel>
  );
};

export default FloorMapPanel;
