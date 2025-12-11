import * as React from "react";
import { useEffect, useState } from "react";
import { AadHttpClient, SPHttpClient } from "@microsoft/sp-http";
import { IEmployeeDirectoryProps } from "./IEmployeeDirectoryProps";
import { Employee, PresenceState } from "../models/Employee";
import EmployeeCard from "./EmployeeCard";
import styles from "./EmployeeDirectory.module.scss";

// NEW: Org Chart + Floor map imports
import OrgChartPanel from "./OrgChartPanel";
import FloorMapPanel from "./FloorMapPanel";

const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("");

const EmployeeDirectory: React.FC<IEmployeeDirectoryProps> = ({ context, listName }) => {
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [filteredEmployees, setFilteredEmployees] = useState<Employee[]>([]);
  const [loading, setLoading] = useState<boolean>(true);

  // UI states
  const [search, setSearch] = useState<string>("");
  const [letter, setLetter] = useState<string | null>(null);
  const [department, setDepartment] = useState<string>("");
  const [jobTitle, setJobTitle] = useState<string>("");
  const [zoom, setZoom] = useState<number>(100);

  // NEW: Org Chart + Floor Map state
  const [showOrgChart, setShowOrgChart] = useState<boolean>(false);
  const [showFloorMap, setShowFloorMap] = useState<boolean>(false);
  const [orgChartData, setOrgChartData] = useState<any>(null);

  /* --------------------------------------------
     LOAD EMPLOYEES FROM SHAREPOINT LIST
  -------------------------------------------- */
  const loadEmployees = async (): Promise<void> => {
    setLoading(true);
    try {
      const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(
        listName
      )}')/items?$top=5000`;

      const spRes = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const json = await spRes.json();

      const listEmployees: Employee[] = (json.value || []).map((item: any) => ({
  Id: Number(item.Id),
  Title: item.Title || "",
  JobTitle: item.JobTitle || "",
  Department: item.Department || "",
  Email: item.Email || "",
  Photo: item.ProfilePhoto || "",
  
  // FIXED manager mapping (Person field → email)
  Manager: item.Manager ? item.Manager.Email : "",

  presence: "offline"
}));


      setEmployees(listEmployees);
      setFilteredEmployees(listEmployees);

      await loadPresence(listEmployees);

    } catch (err) {
      console.error("List load error:", err);
    } finally {
      setLoading(false);
    }
  };

  /* --------------------------------------------
     LOAD MICROSOFT GRAPH PRESENCE
  -------------------------------------------- */
  const loadPresence = async (users: Employee[]): Promise<void> => {
    try {
      const client: AadHttpClient = await context.aadHttpClientFactory.getClient("https://graph.microsoft.com");
      const updated = users.map((u) => ({ ...u }));

      for (const emp of updated) {
        if (!emp.Email) {
          emp.presence = "offline";
          continue;
        }

        try {
          const presenceRes = await client.get(
            `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(emp.Email)}/presence`,
            AadHttpClient.configurations.v1
          );

          if (presenceRes.ok) {
            const pJson = await presenceRes.json();
            const avail = (pJson.availability || "offline").toLowerCase() as PresenceState;
            emp.presence = ["available", "busy", "away"].includes(avail) ? avail : "offline";
          } else {
            emp.presence = "offline";
          }
        } catch {
          emp.presence = "offline";
        }
      }

      setEmployees(updated);
      applyFilters(updated);

    } catch (e) {
      console.error("Presence error:", e);
    }
  };

  /* --------------------------------------------
     FILTER EMPLOYEE LIST
  -------------------------------------------- */
  const applyFilters = (source: Employee[] = employees): void => {
    let out = [...source];

    if (search.trim()) {
      const s = search.trim().toLowerCase();
      out = out.filter(
        (e) =>
          e.Title.toLowerCase().includes(s) ||
          (e.JobTitle || "").toLowerCase().includes(s)
      );
    }

    if (letter) {
      out = out.filter((e) => e.Title.charAt(0).toUpperCase() === letter);
    }

    if (department) {
      out = out.filter((e) => (e.Department || "").toLowerCase() === department.toLowerCase());
    }

    if (jobTitle) {
      out = out.filter((e) => (e.JobTitle || "").toLowerCase() === jobTitle.toLowerCase());
    }

    setFilteredEmployees(out);
  };

  /* --------------------------------------------
     INITIAL LOAD + REFRESH PRESENCE
  -------------------------------------------- */
  useEffect(() => {
    void loadEmployees();

    const presenceTimer = setInterval(() => {
      void loadPresence(employees);
    }, 60000);

    return () => clearInterval(presenceTimer);
  }, []);

  useEffect(() => {
    applyFilters();
  }, [search, letter, department, jobTitle, employees]);

  /* --------------------------------------------
     BUILD ORG CHART HIERARCHY
  -------------------------------------------- */
  const buildOrgHierarchy = (): any => {
    if (!employees.length) return null;

    const root = employees.find((e) => !e.Manager);
    if (!root) return null;

    const buildNode = (emp: Employee): any => ({
      name: emp.Title,
      title: emp.JobTitle,
      department: emp.Department,
      email: emp.Email,
      children: employees
        .filter((x) => x.Manager === emp.Email)
        .map((child) => buildNode(child)),
    });

    return buildNode(root);
  };

  /* --------------------------------------------
     DROPDOWN OPTIONS
  -------------------------------------------- */
  const departments = Array.from(new Set(employees.map((e) => e.Department).filter(Boolean)));
  const jobTitles = Array.from(new Set(employees.map((e) => e.JobTitle).filter(Boolean)));

  /* --------------------------------------------
     RENDER UI
  -------------------------------------------- */
  return (
    <div className={styles.directoryWrapper}>

      {/* TOP BAR */}
      <div className={styles.topBar}>
        <div className={styles.leftGroup}>
          <input
            className={styles.searchBox}
            placeholder="Search employees"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>

        <div className={styles.rightGroup}>
          <button className={styles.actionBtn} onClick={() => setShowFloorMap(true)}>
            Floor Map
          </button>

          <button
            className={styles.actionBtn}
            onClick={() => {
              setOrgChartData(buildOrgHierarchy());
              setShowOrgChart(true);
            }}
          >
            Org Chart
          </button>

          <div className={styles.zoomControl}>
            <label>Zoom</label>
            <input
              type="range"
              min={70}
              max={150}
              value={zoom}
              onChange={(e) => setZoom(Number(e.target.value))}
            />
          </div>
        </div>
      </div>

      {/* ALPHABET FILTER */}
      <div className={styles.lettersRow}>
        <button
          className={`${styles.letterButton} ${letter === null ? styles.active : ""}`}
          onClick={() => setLetter(null)}
        >
          👥
        </button>

        {letters.map((l) => (
          <button
            key={l}
            className={`${styles.letterButton} ${letter === l ? styles.active : ""}`}
            onClick={() => setLetter(l)}
          >
            {l}
          </button>
        ))}
      </div>

      {/* DROPDOWN FILTERS */}
      <div className={styles.filterRow}>
        <select className={styles.select} value={department} onChange={(e) => setDepartment(e.target.value)}>
          <option value="">All Departments</option>
          {departments.map((d) => (
            <option key={d}>{d}</option>
          ))}
        </select>

        <select className={styles.select} value={jobTitle} onChange={(e) => setJobTitle(e.target.value)}>
          <option value="">All Job Titles</option>
          {jobTitles.map((j) => (
            <option key={j}>{j}</option>
          ))}
        </select>

        <button
          className={styles.resetButton}
          onClick={() => {
            setSearch("");
            setLetter(null);
            setDepartment("");
            setJobTitle("");
          }}
        >
          Reset Filters
        </button>
      </div>

      {/* CARD GRID */}
      {loading ? (
        <div className={styles.loading}>Loading employees…</div>
      ) : (
        <div
          className={styles.cardGrid}
          style={{
            gridTemplateColumns: `repeat(auto-fill, ${Math.round((zoom / 100) * 310)}px)`
          }}
        >
          {filteredEmployees.map((emp) => (
            <EmployeeCard key={emp.Id} employee={emp} />
          ))}
        </div>
      )}

      {/* ORG CHART PANEL */}
      <OrgChartPanel
        isOpen={showOrgChart}
        onDismiss={() => setShowOrgChart(false)}
        data={orgChartData}
      />

      {/* FLOOR MAP PANEL */}
      <FloorMapPanel
        isOpen={showFloorMap}
        onDismiss={() => setShowFloorMap(false)}
        imageUrl="https://yourtenant.sharepoint.com/sites/intranet/Shared%20Documents/floor-map.png"
      />
    </div>
  );
};

export default EmployeeDirectory;
