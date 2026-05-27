import * as React from "react";
import { useHistory } from "react-router-dom";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import styles from "./Mdr.module.scss";
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import { SPFI } from "@pnp/sp/presets/all";
import { getSP } from "../loc/pnpjsConfig";
import UserContext from "../../../GlobalContext/context";
import Provider from "../../../GlobalContext/provider";
import { DefaultButton } from "@fluentui/react";
import "bootstrap/dist/css/bootstrap.min.css";
import { ErrorLogger } from "../../../utils/ErrorLogger";

interface IMDRProps {
  context: WebPartContext;
  siteUrl: string;
}

interface IProjectItem {
  ID: number;
  ProjectName: string;
  ProjectType: {
    ProjectType: string;
  };
}

interface IDeliverableItem {
  LatestMDRStatus: string;
  Id: number;
  DocNumber: string;

  Deliverables: {
    Deliverables: string;
  };

  Area: {
    Area: string;
  };

  Organisation: {
    Organisation: string;
  };

  DocumentType: {
    DocumentType: string;
  };

  RevisionNumber: number;
  DocumentComments: string;
  Status: string;

  ProjectCreationListID: {
    ProjectName: string;
  };

  AssignedTo: {
    Title: string;
  };
  IncomingRev: string[];
  IncomingStatus: string[];
  IncomingSubmitDate: string[];
  IncomingTransmittalNo: string[];
  OutgoingReturnDate: string[];
  OutgoingTransmittalNo: string[];
  OutgoingCodeStatus: string[];
  ServerLink: string[];
  FileName: string[];
}

interface IApprovalItem {
  ID: number;
  RevisionNumber: number;
  RequestedRole: string;
  IncomingDate: string;
  OutgoingDate: string;
  ApproverRole: string;
  OutgoingStatus: string;
}

interface IDocumentItem {
  ID: number;
  FileRef: string;
  SharedLink: string;
  FileLeafRef: string;
}

const MDR: React.FC<IMDRProps> = ({ context, siteUrl }) => {
  const elementRef = React.useRef<HTMLDivElement>(null);
  const sp: SPFI = getSP();
  const spContext = spfi().using(SPFx(context));
  const { useHide }: any = React.useContext(UserContext);

  const [projectName, setProjectName] = React.useState<string>("");
  const [deliverablesDetails, setDeliverablesDetails] = React.useState<
    IDeliverableItem[]
  >([]);
  const [submissions, setSubmissions] = React.useState<number[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);

  React.useEffect(() => {
    console.log(" UI Data:", deliverablesDetails);
    const getCreationIdFromUrl = () => {
      try {
        let params: URLSearchParams;

        if (window.location.search) {
          params = new URLSearchParams(window.location.search);
        } else if (window.location.hash) {
          const hash = window.location.hash;
          params = new URLSearchParams(hash.split("?")[1] || "");
        } else {
          return null;
        }

        const creationIdParam = params.get("ProjectId");
        console.log("CreationId:", creationIdParam);
        return creationIdParam ? parseInt(creationIdParam, 10) : null;
      } catch (error) {
        console.error("Error parsing URL:", error);
        ErrorLogger.logError(spContext, error, "parseURL", "MDRWebPart");
        return null;
      }
    };

    const creationIdFromUrl = getCreationIdFromUrl();
    console.log(" Parsed CreationId:", creationIdFromUrl);
    if (creationIdFromUrl) {
      getCreationDetails(creationIdFromUrl);
    } else {
      setLoading(false);
    }
  }, []);

  const getCreationDetails = async (creationId: number) => {
    console.log(" getCreationDetails CALLED:", creationId);
    setLoading(true);
    try {
      // throw new Error("Test Error Logging"); for Error Logging Testing
      const sp = spfi().using(SPFx(context));

      const projectItems: IProjectItem[] = await sp.web.lists
        .getByTitle("ProjectCreationList")
        .items.select("ID", "ProjectName", "ProjectType/ProjectType")
        .expand("ProjectType")
        .filter(`ID eq ${creationId}`)
        .top(1)();
      console.log(" Project Items:", projectItems);

      if (projectItems.length > 0) {
        setProjectName(projectItems[0].ProjectName);
        await getDeliverablesDetails(creationId);
      } else {
        setLoading(false);
      }
    } catch (error) {
      console.error("Error retrieving project details:", error);
      await ErrorLogger.logError(
        spContext,
        error,
        "getCreationDetails",
        "MDRWebPart",
      );
      setLoading(false);
    }
  };

  const getDocumentUrl = async (
    creationId: number,
    deliverableId: number,
    revision: number,
  ): Promise<{ url: string; name: string }> => {
    try {
      const sp = spfi().using(SPFx(context));

      const documents: IDocumentItem[] = await sp.web.lists
        .getByTitle("DeliverablesDocument")
        .items.select("ID", "FileRef", "FileLeafRef", "SharedLink")
        .filter(
          `ProjectID eq ${creationId} and DeliverablesDetailsId eq ${deliverableId} and Revision eq ${revision}`,
        )
        .orderBy("Created", false)
        .top(1)();

      if (documents.length > 0) {
        const fileUrl = `https://officeindia.sharepoint.com/sites/ESSA/SitePages/PDFViewer.aspx?docId=${documents[0].ID}`;

        return {
          url: fileUrl,
          name: documents[0].FileLeafRef,
        };
      } else {
        return { url: "", name: "" };
      }
    } catch (error) {
      console.error("Error fetching document link:", error);
      await ErrorLogger.logError(
        spContext,
        error,
        "getDocumentUrl",
        "MDRWebPart",
      );
      return { url: "", name: "" };
    }
  };

  const getDeliverablesDetails = async (creationId: number) => {
    console.log(" getDeliverablesDetails CALLED:", creationId);
    try {
      const sp = spfi().using(SPFx(context));

      const deliverables: IDeliverableItem[] = await sp.web.lists
        .getByTitle("DeliverablesDetails")
        .items.select(
          "Id",
          "DocNumber",

          //  LOOKUPS (IMPORTANT)
          "Deliverables/Deliverables",
          "Area/Area",
          "Organisation/Organisation",
          "DocumentType/DocumentType",

          //  OTHER FIELDS
          "RevisionNumber",
          "DocumentComments",
          "Status",

          //  LOOKUPS
          "ProjectCreationListID/ProjectName",
          "AssignedTo/Title",
        )
        .expand(
          "Deliverables",
          "Area",
          "Organisation",
          "DocumentType",
          "ProjectCreationListID",
          "AssignedTo",
        )
        .filter(`ProjectCreationListID/Id eq ${creationId}`)
        .orderBy("Created", false)();

      if (deliverables.length > 0) {
        const highestRevision = Math.max(
          ...deliverables.map((item) => Number(item.RevisionNumber) || 0),
        );
        console.log("Highest Revision:", highestRevision);
        console.log("Deliverables RAW:", deliverables);
        console.log("Deliverables LENGTH:", deliverables.length);

        const submissionsArray = [];
        for (let i = 0; i <= highestRevision; i++) {
          submissionsArray.push(i);
        }
        setSubmissions(submissionsArray);

        const processedDeliverables = await Promise.all(
          deliverables.map(async (item) => {
            const processedItem: IDeliverableItem = {
              ...item,
              IncomingRev: [],
              IncomingStatus: [],
              IncomingSubmitDate: [],
              IncomingTransmittalNo: [],
              OutgoingReturnDate: [],
              OutgoingTransmittalNo: [],
              OutgoingCodeStatus: [],
              ServerLink: [],
              FileName: [],
            };

            try {
              const approvals: IApprovalItem[] = await sp.web.lists
                .getByTitle("ProjectApprovals")
                .items.select(
                  "*",
                  "ID",
                  "RevisionNumber",
                  "RequestedRole",
                  "IncomingDate",
                  "OutgoingDate",
                  "ApproverRole",
                  "OutgoingStatus",
                )
                .filter(
                  `ProjectCreationListID/Id eq ${creationId} and DeliverablesDetailsId/Id eq ${item.Id} and ApproverRole eq 'Document Controller'`,
                )
                .orderBy("RevisionNumber", true)();

              if (approvals.length > 0) {
                const sortedApprovals = approvals.sort(
                  (a, b) => Number(a.RevisionNumber) - Number(b.RevisionNumber),
                );

                let latestStatus = "-";
                for (let i = sortedApprovals.length - 1; i >= 0; i--) {
                  if (
                    sortedApprovals[i].OutgoingStatus &&
                    sortedApprovals[i].OutgoingStatus.trim() !== ""
                  ) {
                    latestStatus = sortedApprovals[i].OutgoingStatus;
                    break;
                  }
                }

                processedItem.LatestMDRStatus = latestStatus;
              }

              for (let i = 0; i <= highestRevision; i++) {
                if (i <= item.RevisionNumber) {
                  const revDisplay = i === 0 ? "0" : i.toString();
                  processedItem.IncomingRev.push(revDisplay);

                  const matchingApprovals = approvals.filter(
                    (app) => Number(app.RevisionNumber) === i,
                  );

                  const incoming = matchingApprovals.find(
                    (app) => app.RequestedRole === "Vendor" && app.IncomingDate,
                  );
                  if (incoming) {
                    processedItem.IncomingSubmitDate.push(
                      incoming.IncomingDate,
                    );
                    processedItem.IncomingTransmittalNo.push(`TR-00${i + 1}`);
                    processedItem.IncomingStatus.push("IFA");
                  } else {
                    processedItem.IncomingSubmitDate.push("-");
                    processedItem.IncomingTransmittalNo.push("-");
                    processedItem.IncomingStatus.push("-");
                  }

                  const outgoing = matchingApprovals.find(
                    (app) => app.OutgoingDate,
                  );
                  if (outgoing) {
                    processedItem.OutgoingReturnDate.push(
                      outgoing.OutgoingDate,
                    );
                    processedItem.OutgoingTransmittalNo.push(`TR-00${i + 1}`);
                    processedItem.OutgoingCodeStatus.push(
                      outgoing.OutgoingStatus || "-",
                    );
                  } else {
                    processedItem.OutgoingReturnDate.push("-");
                    processedItem.OutgoingTransmittalNo.push("-");
                    processedItem.OutgoingCodeStatus.push("-");
                  }

                  const doc = await getDocumentUrl(creationId, item.Id, i);
                  processedItem.ServerLink.push(doc.url);
                  processedItem.FileName.push(
                    doc.name ? doc.name.replace(/^6_\d+_/, "") : "",
                  );
                } else {
                  processedItem.IncomingRev.push("-");
                  processedItem.IncomingStatus.push("-");
                  processedItem.IncomingSubmitDate.push("-");
                  processedItem.IncomingTransmittalNo.push("-");
                  processedItem.OutgoingReturnDate.push("-");
                  processedItem.OutgoingTransmittalNo.push("-");
                  processedItem.OutgoingCodeStatus.push("-");
                  processedItem.ServerLink.push("");
                  processedItem.FileName.push("");
                }
              }
            } catch (error) {
              console.error(
                `Error processing approvals for deliverable ${item.Id}:`,
                error,
              );
              await ErrorLogger.logError(
                spContext,
                error,
                `processDeliverable_${item.Id}`,
                "MDRWebPart",
              );
            }

            return processedItem;
          }),
        );
        console.log(" Processed Deliverables:", processedDeliverables);
        setDeliverablesDetails(processedDeliverables);
      }
    } catch (error) {
      console.error("Error retrieving deliverables details:", error);
      await ErrorLogger.logError(
        spContext,
        error,
        "getDeliverablesDetails",
        "MDRWebPart",
      );
    } finally {
      setLoading(false);
    }
  };

  const getOrdinal = (n: number): string => {
    if (n === 1) return "st";
    if (n === 2) return "nd";
    if (n === 3) return "rd";
    return "th";
  };

  const getStatusBadgeClass = (status: string): string => {
    switch (status) {
      case "In-Progress":
        return styles.badgeInProgress;
      case "Pending":
        return styles.badgePending;
      case "Rework":
        return styles.badgeRework;
      case "Rejected":
        return styles.badgeRejected;
      case "Approved":
        return styles.badgeApproved;
      default:
        return styles.badgeDefault;
    }
  };

  const formatDate = (dateString: string): string => {
    if (!dateString || dateString === "-") return "-";
    try {
      const date = new Date(dateString);
      return date.toLocaleDateString("en-GB");
    } catch (error) {
      ErrorLogger.logError(spContext, error, "formatDate", "MDRWebPart");
      return dateString;
    }
  };

  const handleBack = () => {
    window.location.href =
      "https://officeindia.sharepoint.com/sites/ESSA/SitePages/MyProject.aspx";
  };

  const exportMDRToExcel = () => {
    try {
      const table = document.querySelector(
        `.${styles.newTable}`,
      ) as HTMLElement;
      if (!table) return;

      const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:x="urn:schemas-microsoft-com:office:excel"
          xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="UTF-8">
        <style>
          ${getComputedStyleCSS()}
        </style>
      </head>
      <body>
        ${table.outerHTML}
      </body>
    </html>`;

      const blob = new Blob([html], { type: "application/vnd.ms-excel" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `MDR_Export_${new Date().toISOString().slice(0, 10)}.xls`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (error) {
      console.error("exportMDRToExcel error:", error);

      ErrorLogger.logError(spContext, error, "exportMDRToExcel", "MDRWebPart");
    }
  };

  function getComputedStyleCSS() {
    let css = "";
    for (const sheet of Array.from(document.styleSheets)) {
      try {
        for (const rule of Array.from(sheet.cssRules || [])) {
          css += rule.cssText;
        }
      } catch (e) {
        // Ignore CORS-restricted stylesheets
      }
    }
    return css;
  }

  if (loading) {
    return (
      <div
        id="wrapper"
        ref={elementRef}
        style={{
          background: "#fff",
          minHeight: "100vh",
        }}
      >
        <div className="app-menu" id="myHeader">
          <VerticalSideBar _context={sp} />
        </div>
        <div className="content-page">
          <HorizontalNavbar _context={sp} siteUrl={siteUrl} />
          <div
            className="content"
            style={{
              marginLeft: `${!useHide ? "240px" : "80px"}`,
              marginTop: "0rem",
            }}
          >
            <div className="container-fluid paddb">
              <div className={styles.mdrContainer}>
                <div
                  className={styles.loading}
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    justifyContent: "center",
                    alignItems: "center",
                    position: "fixed",
                    top: 0,
                    left: 0,
                    width: "100%",
                    height: "100vh",
                    zIndex: 9999,
                    background: "#fff",
                  }}
                >
                  <img
                    src={require("../../../CustomAsset/birdloader.gif")}
                    alt="Loading..."
                    style={{ width: "90px", height: "90px" }}
                  />
                  Loading MDR data...
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div id="wrapper" ref={elementRef}>
      <div className="app-menu" id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>

      <div className="content-page">
        <HorizontalNavbar _context={sp} siteUrl={siteUrl} />
        <div
          className="content"
          style={{
            marginLeft: `${!useHide ? "240px" : "80px"}`,
            marginTop: "0rem",
          }}
        >
          <div className="container-fluid paddb">
            <div className={styles.mdrContainer}>
              <div className={styles.pageHeader}>
                <h2>MDR - {projectName || "Project Details"}</h2>
                <div className={styles.headerActions}>
                  <DefaultButton
                    text="Export to Excel"
                    iconProps={{ iconName: "ExcelDocument" }}
                    onClick={exportMDRToExcel}
                    className={styles.btnExport}
                  />
                  <DefaultButton
                    text="Back"
                    iconProps={{ iconName: "NavigateBack" }}
                    onClick={handleBack}
                    className={styles.backButton}
                  />
                </div>
              </div>

              {/* MDR Table */}
              <div className={styles.mdrTableSection}>
                <div className={styles.card}>
                  <div className={styles.cardBody}>
                    <div className={styles.tableResponsive}>
                      <table className={styles.newTable}>
                        <thead>
                          <tr>
                            <th rowSpan={3} className={styles.tableHeader}>
                              SNo
                            </th>
                            <th rowSpan={3} className={styles.tableHeader}>
                              Project Name
                            </th>
                            <th rowSpan={3} className={styles.tableHeader}>
                              Document No
                            </th>
                            <th rowSpan={3} className={styles.tableHeader}>
                              Deliverables
                            </th>
                            <th rowSpan={3} className={styles.tableHeader}>
                              Area
                            </th>
                            <th rowSpan={3} className={styles.tableHeader}>
                              Organization
                            </th>
                            <th
                              colSpan={2}
                              rowSpan={2}
                              className={styles.tableHeader}
                            >
                              Current Status
                            </th>
                            <th rowSpan={3} className={styles.tableHeader}>
                              Document Type
                            </th>
                            {submissions.map((_, i) => (
                              <th
                                key={i}
                                colSpan={8}
                                className={styles.submissionHeader}
                              >
                                {i + 1}
                                {getOrdinal(i + 1)} Submitted
                              </th>
                            ))}
                          </tr>
                          <tr>
                            {submissions.map((_, i) => (
                              <React.Fragment key={i}>
                                <th
                                  colSpan={4}
                                  className={styles.incomingHeader}
                                >
                                  Incoming
                                </th>
                                <th
                                  colSpan={4}
                                  className={styles.outgoingHeader}
                                >
                                  Outgoing
                                </th>
                              </React.Fragment>
                            ))}
                          </tr>
                          <tr>
                            <th className={styles.latestRevision}>
                              Latest Revision
                            </th>
                            <th className={styles.commentHeader}>
                              Comment Status
                            </th>
                            {submissions.map((_, i) => (
                              <React.Fragment key={i}>
                                <th className={styles.revHeader}>Rev</th>
                                <th className={styles.statusHeader}>Status</th>
                                <th className={styles.dateHeader}>
                                  Submit Date
                                </th>
                                <th className={styles.transmittalHeader}>
                                  Transmittal No
                                </th>
                                <th className={styles.dateHeader}>
                                  Return Date
                                </th>
                                <th className={styles.transmittalHeader}>
                                  Transmittal No
                                </th>
                                <th className={styles.statusHeader}>
                                  Code Status
                                </th>
                                <th className={styles.documentHeader}>
                                  Document Link
                                </th>
                              </React.Fragment>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {deliverablesDetails.map((item, index) => (
                            <tr key={item.Id}>
                              <td className={styles.tableCell}>{index + 1}</td>
                              <td className={styles.tableCell}>
                                {item.ProjectCreationListID?.ProjectName}
                              </td>
                              <td className={styles.tableCell}>
                                {item.DocNumber}
                              </td>
                              <td>{item.Deliverables?.Deliverables || "-"}</td>

                              <td>{item.Area?.Area || "-"}</td>
                              <td>{item.Organisation?.Organisation || "-"}</td>
                              <td className={styles.tableCell}>
                                <div
                                  className={styles.revisionCircle}
                                  style={
                                    item.RevisionNumber
                                      ? {
                                          backgroundColor: "#e2e8ff",
                                          color: "#7989c7",
                                        }
                                      : {}
                                  }
                                >
                                  {item.RevisionNumber}
                                </div>
                              </td>
                              <td className={styles.tableCell}>
                                {item.LatestMDRStatus ? (
                                  <span
                                    className={`${styles.badge} ${getStatusBadgeClass(item.LatestMDRStatus)}`}
                                  >
                                    {item.LatestMDRStatus}
                                  </span>
                                ) : (
                                  "-"
                                )}
                              </td>
                              <td>{item.DocumentType?.DocumentType || "-"}</td>

                              {submissions.map((_, submissionIndex) => {
                                const incomingRev =
                                  item.IncomingRev[submissionIndex] || "-";
                                const incomingStatus =
                                  item.IncomingStatus[submissionIndex] || "-";
                                const incomingSubmitDate =
                                  item.IncomingSubmitDate[submissionIndex] ||
                                  "-";
                                const incomingTransmittalNo =
                                  item.IncomingTransmittalNo[submissionIndex] ||
                                  "-";
                                const outgoingReturnDate =
                                  item.OutgoingReturnDate[submissionIndex] ||
                                  "-";
                                const outgoingTransmittalNo =
                                  item.OutgoingTransmittalNo[submissionIndex] ||
                                  "-";
                                const outgoingCodeStatus =
                                  item.OutgoingCodeStatus[submissionIndex] ||
                                  "-";
                                const serverLink =
                                  item.ServerLink[submissionIndex] || "";
                                const fileName =
                                  item.FileName[submissionIndex] || "";

                                return (
                                  <React.Fragment key={submissionIndex}>
                                    <td className={styles.tableCell}>
                                      <div
                                        className={styles.revisionCircle}
                                        style={
                                          incomingRev !== "-"
                                            ? {
                                                backgroundColor: "#e2e8ff",
                                                color: "#7989c7",
                                              }
                                            : {}
                                        }
                                      >
                                        {incomingRev}
                                      </div>
                                    </td>
                                    <td className={styles.tableCell}>
                                      {incomingStatus !== "-" ? (
                                        <span
                                          className={`${styles.badge} ${getStatusBadgeClass(incomingStatus)}`}
                                        >
                                          {incomingStatus}
                                        </span>
                                      ) : (
                                        "-"
                                      )}
                                    </td>
                                    <td className={styles.tableCell}>
                                      <div
                                        className={styles.dateCircle}
                                        style={
                                          incomingSubmitDate !== "-"
                                            ? {
                                                backgroundColor: "#c7eecf",
                                                color: "#000",
                                              }
                                            : {}
                                        }
                                      >
                                        {formatDate(incomingSubmitDate)}
                                      </div>
                                    </td>
                                    <td className={styles.tableCell}>
                                      {incomingTransmittalNo}
                                    </td>
                                    <td className={styles.tableCell}>
                                      <div className={styles.returnDateCircle}>
                                        {formatDate(outgoingReturnDate)}
                                      </div>
                                    </td>
                                    <td className={styles.tableCell}>
                                      {outgoingTransmittalNo}
                                    </td>
                                    <td className={styles.tableCell}>
                                      {outgoingCodeStatus !== "-" ? (
                                        <span
                                          className={`${styles.badge} ${getStatusBadgeClass(outgoingCodeStatus)}`}
                                        >
                                          {outgoingCodeStatus}
                                        </span>
                                      ) : (
                                        "-"
                                      )}
                                    </td>
                                    <td className={styles.tableCell}>
                                      {fileName ? (
                                        <a
                                          href={serverLink}
                                          target="_blank"
                                          rel="noopener noreferrer"
                                          className={styles.documentLink}
                                          title={fileName}
                                        >
                                          Link
                                        </a>
                                      ) : (
                                        "-"
                                      )}
                                    </td>
                                  </React.Fragment>
                                );
                              })}
                            </tr>
                          ))}
                        </tbody>
                      </table>

                      {deliverablesDetails.length === 0 && (
                        <div className={styles.noData}>
                          <p>No deliverables found for this project.</p>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

// Wrap with Provider for UserContext
const MDRWithProvider: React.FC<IMDRProps> = (props) => (
  <Provider>
    <MDR {...props} />
  </Provider>
);

export default MDRWithProvider;
