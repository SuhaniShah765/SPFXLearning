import * as React from "react";
import styles from "./EmployeeCard.module.scss";
import { Employee } from "../models/Employee";
import { Icon } from "@fluentui/react";

interface Props {
  employee: Employee;
}

const presenceClasses: Record<string, string> = {
  available: styles.presenceAvailable,
  busy: styles.presenceBusy,
  away: styles.presenceAway,
  offline: styles.presenceOffline
};

const EmployeeCard: React.FC<Props> = ({ employee }): JSX.Element => {
  const initials = React.useMemo(() => {
    return (employee.Title || "")
      .split(" ")
      .map((p) => p.charAt(0))
      .slice(0, 2)
      .join("")
      .toUpperCase();
  }, [employee.Title]);

  const goTo = (url: string): void => {
  void window.open(url, "_blank");
};


  return (
    <div className={styles.card}>
      <div className={styles.banner} />

      <div className={styles.photoWrapper}>
        {employee.Photo ? (
          <img src={employee.Photo} className={styles.photo} alt={employee.Title} />
        ) : (
          <div className={styles.initials}>{initials}</div>
        )}

        <span
          className={`${styles.presenceDot} ${
            presenceClasses[employee.presence || "offline"]
          }`}
          title={`Presence: ${employee.presence || "offline"}`}
        />
      </div>

      <div className={styles.content}>
        <div className={styles.name}>{employee.Title}</div>
        <div className={styles.job}>{employee.JobTitle}</div>
        <div className={styles.department}>{employee.Department}</div>
      </div>

      <div className={styles.sep} />

      <div className={styles.icons}>
        <button
          className={styles.iconBtn}
          title="Email"
          onClick={() => employee.Email && goTo(`mailto:${employee.Email}`)}
        >
          <Icon iconName="Mail" />
        </button>

        <button
          className={styles.iconBtn}
          title="Chat on Teams"
          onClick={() =>
            employee.Email &&
            goTo(
              `https://teams.microsoft.com/l/chat/0/0?users=${encodeURIComponent(
                employee.Email
              )}`
            )
          }
        >
          <Icon iconName="TeamsLogo" />
        </button>

        <button
          className={styles.iconBtn}
          title="OneDrive"
          onClick={() =>
            employee.Email &&
            goTo(
              `https://${window.location.hostname.split(".")[0]}-my.sharepoint.com/personal/${
                employee.Email.replace("@", "_")
              }/`
            )
          }
        >
          <Icon iconName="OneDriveLogo" />
        </button>
      </div>
    </div>
  );
};

export default EmployeeCard;
