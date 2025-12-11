import * as React from "react";
import { Panel, PanelType } from "@fluentui/react";
import OrgChart from "@dabeng/react-orgchart";
import styles from "./OrgChartPanel.module.scss";

interface Props {
  isOpen: boolean;
  onDismiss: () => void;
  data: any; // Hierarchy tree
}

const OrgChartPanel: React.FC<Props> = ({ isOpen, onDismiss, data }) => {
  return (
    <Panel
      isOpen={isOpen}
      onDismiss={onDismiss}
      type={PanelType.large}
      headerText="Organizational Chart"
      closeButtonAriaLabel="Close"
    >
      <div className={styles.wrapper}>
        {data ? (
          <OrgChart datasource={data} collapsible={true} />
        ) : (
          <div className={styles.loading}>Building org chart…</div>
        )}
      </div>
    </Panel>
  );
};

export default OrgChartPanel;
