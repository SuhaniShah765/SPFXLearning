export type PresenceState = "available" | "busy" | "away" | "offline";

export interface Employee {
  Id: number;
  Title: string;
  EmployeeCode?: number;
  Department?: string;
  JobTitle?: string;
  Email: string;
  Photo?: string;
  PhoneNumber?: string;
  Location?: string;
  Manager?: string;
  JoiningDate?: string;
  WorkAnniversary?: string;
  Skills?: string;
  Status?: string;
  AboutMe?: string;

  presence?: PresenceState;  // ⭐ Use the type
}
