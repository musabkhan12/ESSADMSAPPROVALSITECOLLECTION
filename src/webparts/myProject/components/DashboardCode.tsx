import * as React from "react";

import { getAllItems } from "../loc/pnpjsConfig";
import styles from "./Dashboard.module.scss";
import { DefaultButton } from "@fluentui/react/lib/Button";
import { IContextualMenuProps } from "@fluentui/react/lib/ContextualMenu";
import Provider from "../../../GlobalContext/provider";
import { IMyProjectProps } from "./IMyProjectProps";
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import UserContext from "../../../GlobalContext/context";
import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import { ErrorLogger } from "../../../utils/ErrorLogger";
import { SPFI } from "@pnp/sp/presets/all";
import NewRequest from "./NewRequest";
import { getSP } from "../loc/pnpjsConfig";
import { spfi, SPFx } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "./Dash.css";
import {
  PieChart,
  DataVizPalette,
  IDataPoint as FluentIDataPoint,
} from "@fluentui/react-charting";

interface IDataPoint extends FluentIDataPoint {
  x: string | number;
  color: string;
}

interface IProjectItem {
  ID: number;
  ProjectName: string;
  ProjectOverview: string;
  PreparedBy: {
    Title: string;
  };
  ProjectType: {
    ProjectType: string;
  };
  ClientName: string;
  ProjectStartDate: string;
  Status: string;
}

interface IApprovalItem {
  ID: number;
  ProjectCreationListIDId: number;
  DeliverablesDetailsIdId: number;
  Status: string;
  Level?: string;
  AssignedTo?: {
    Title: string;
  };
  ApprovalDate?: string;
  Comments?: string;
  Remarks?: string;
  ApproverRole?: string;
}

interface IApprovalHierarchyItem {
  ID: number;
  DeliverablesDetailsIdId: number;
  AssignedTo: {
    Title: string;
  }[];
  ApproverRole: string;
  Level: number;
  Status: string;
}

interface IDeliverableItem {
  ID: number;
  ProjectCreationListIDId: number;
  Status: string;
  DeliverablesDocumentIDId: number | null;
  Deliverables: {
    ID: number;
    Deliverables: string;
  };
  Area: {
    ID: number;
    Area: string;
  };
  Organisation: {
    ID: number;
    Organisation: string;
  };
  DocumentType: {
    ID: number;
    DocumentType: string;
  };
  AssignedTo: {
    Title: string;
  };
  DocNumber: string;
  Modified: string;
  RevisionNumber: string;
  FileLeafRef: string;
  DocumentComments: string;
  IsReworked: string;
}

interface ICommentItem {
  ID: number;
  ProjectID: number;
}

interface ICard {
  id: number;
  title: string;
  type: string;
  user: string;
  statusCounts: number[];
  description: string;
  documents: number;
  comments: number;
  progress: { current: number; total: number };
  status: string;
}

interface IProjectDetails {
  ProjectName: string;
  ProjectOverview: string;
  PreparedBy: string;
  ProjectType: string;
  ClientName: string;
  ProjectStartDate: string;
  Status: string;
  ProjCreationId: number;
  CompletedTask: number;
  TotalNoComment: number;
  Members: number;
  DeliverablesDetailsArr: IDeliverableItem[];
}

const Dashboard = ({ props }: any) => {
  const elementRef = React.useRef<HTMLDivElement>(null);
  const sp: SPFI = getSP();
  const siteUrl = props.siteUrl;
  const { useHide }: any = React.useContext(UserContext);
  const [cards, setCards] = React.useState<ICard[]>([]);
  const [loading, setLoading] = React.useState(true);
  const [showDetails, setShowDetails] = React.useState(false);
  const [selectedProject, setSelectedProject] =
    React.useState<IProjectDetails | null>(null);
  const [deliverableItems, setDeliverableItems] = React.useState<
    IDeliverableItem[]
  >([]);
  const [commentItems, setCommentItems] = React.useState<ICommentItem[]>([]);
  const [approvalItems, setApprovalItems] = React.useState<IApprovalItem[]>([]);
  const [showAuditHistory, setShowAuditHistory] = React.useState(false);
  const [auditHistoryData, setAuditHistoryData] = React.useState<
    IApprovalItem[]
  >([]);

  const [showApprovalHierarchy, setShowApprovalHierarchy] =
    React.useState(false);
  const [selectedHierarchyDeliverableId, setSelectedHierarchyDeliverableId] =
    React.useState<number | null>(null);
  const [approvalHierarchyData, setApprovalHierarchyData] = React.useState<
    IApprovalHierarchyItem[]
  >([]);

  const [statusChartData, setStatusChartData] = React.useState<IDataPoint[]>(
    [],
  );

  const [showNewRequest, setShowNewRequest] = React.useState(false);
  const [searchTerm, setSearchTerm] = React.useState<string>("");
  const [filteredCards, setFilteredCards] = React.useState<ICard[]>([]);
  const [hierarchyLoading, setHierarchyLoading] = React.useState(false);
  const [showEditDraft, setShowEditDraft] = React.useState(false);
  const [draftPayload, setDraftPayload] = React.useState<any>(null);

  React.useEffect(() => {
    if (searchTerm.trim() === "") {
      setFilteredCards(cards);
    } else {
      const filtered = cards.filter(
        (card) =>
          card.title.toLowerCase().includes(searchTerm.toLowerCase()) ||
          card.description.toLowerCase().includes(searchTerm.toLowerCase()) ||
          card.user.toLowerCase().includes(searchTerm.toLowerCase()) ||
          card.type.toLowerCase().includes(searchTerm.toLowerCase()),
      );
      setFilteredCards(filtered);
    }
  }, [searchTerm, cards]);

  // Initialize filteredCards when cards load
  React.useEffect(() => {
    if (cards.length > 0) {
      setFilteredCards(cards);
    }
  }, [cards]);

  React.useEffect(() => {
    const initializeSP = async () => {
      try {
        const sp = spfi().using(SPFx(props.context));
        await loadDashboardData(sp);
      } catch (error) {
        console.error("Error initializing SharePoint:", error);
        await ErrorLogger.logError(
          sp,
          error,
          "initializeSP",
          "DashboardWebPart",
        );
        setLoading(false);
      }
    };

    initializeSP();
  }, [props.context]);

  // Update chart data when project details change
  React.useEffect(() => {
    if (selectedProject && selectedProject.DeliverablesDetailsArr.length > 0) {
      updateChartData(selectedProject.DeliverablesDetailsArr);
    }
  }, [selectedProject]);

  const updateChartData = (deliverables: IDeliverableItem[]) => {
    const statusCount: { [key: string]: number } = {};

    // Count deliverables by status
    deliverables.forEach((item) => {
      const status = item.Status || "Unknown";
      statusCount[status] = (statusCount[status] || 0) + 1;
    });

    // Define color mapping for different statuses
    const statusColors: { [key: string]: string } = {
      "In-Progress": DataVizPalette.color8, // Blue
      Pending: DataVizPalette.color9, // Orange
      Rework: DataVizPalette.color10, // Yellow
      Rejected: DataVizPalette.color11, // Red
      Approved: DataVizPalette.color12, // Green
      Completed: DataVizPalette.color12, // Green (same as Approved)
      Unknown: DataVizPalette.color7, // Gray
    };

    // Convert to IDataPoint format for PieChart
    const chartData: IDataPoint[] = Object.entries(statusCount).map(
      ([status, count]) => ({
        x: status,
        y: count,
        color: statusColors[status] || DataVizPalette.color7,
        legend: status,
      }),
    );

    setStatusChartData(chartData);
  };

  const loadDashboardData = async (sp: any) => {
    try {
      console.log("Loading dashboard data...");

      const projectItems: IProjectItem[] = await sp.web.lists
        .getByTitle("ProjectCreationList")
        .items.select(
          "*",
          "ID",
          "ProjectName",
          "ProjectOverview",
          "PreparedBy/Title",
          "ProjectType/ProjectType",
          "ClientName",
          "ProjectStartDate",
          "Status",
        )
        .expand("PreparedBy", "ProjectType")();

      console.log("Project items loaded:", projectItems);

      const approvalItems: IApprovalItem[] = await getAllItems<IApprovalItem>(
        sp,
        "ProjectApprovals",
        ["*", "AssignedTo/Title"],
        ["AssignedTo"],
      );

      console.log("Approval items loaded:", approvalItems);
      setApprovalItems(approvalItems);

      const deliverables: IDeliverableItem[] =
        await getAllItems<IDeliverableItem>(
          sp,
          "DeliverablesDetails",
          [
            "*",
            "Deliverables/ID",
            "Deliverables/Deliverables",
            "Area/ID",
            "Area/Area",
            "Organisation/ID",
            "Organisation/Organisation",
            "DocumentType/ID",
            "DocumentType/DocumentType",
            "AssignedTo/Title",
            "DocNumber",
            "Modified",
            "RevisionNumber",
            "FileLeafRef",
            "DocumentComments",
            "ProjectCreationListIDId",
            "Status",
            "IsReworked",
            "DeliverablesDocumentIDId",
          ],
          [
            "Deliverables",
            "Area",
            "Organisation",
            "DocumentType",
            "AssignedTo",
          ],
        );

      console.log("Deliverable items loaded:", deliverables);
      setDeliverableItems(deliverables);
      const commentsRaw = await getAllItems<any>(
        sp,
        "DocumentComments",
        ["*", "ID", "ProjectID"],
        [],
      );

      console.log("Raw comment items loaded:", commentsRaw);

      const comments: ICommentItem[] = commentsRaw.map((item: any) => ({
        ID: item.ID,
        ProjectID: item.ProjectID ? Number(item.ProjectID) : 0,
      }));

      console.log("Processed comment items:", comments);
      setCommentItems(comments);

      const projectCards: ICard[] = projectItems.map((proj: IProjectItem) => {
        const projectDeliverables = deliverables.filter(
          (d: IDeliverableItem) => d.ProjectCreationListIDId === proj.ID,
        );
        const projectComments = comments.filter(
          (c: ICommentItem) => c.ProjectID === proj.ID,
        );

        const statusMap = [
          "Pending",
          "In-Progress",
          "Rework",
          "Approved",
          "Rejected",
        ];

        const statusCounts = statusMap.map((status) => {
          return projectDeliverables.filter((d: IDeliverableItem) => {
            // Pending Count
            if (status === "Pending") {
              return d.Status === "Pending" && d.IsReworked !== "Yes";
            }

            // Rework Count
            if (status === "Rework") {
              return (
                d.Status === "Rework" ||
                (d.Status === "Pending" && d.IsReworked === "Yes")
              );
            }

            // Other Status
            return d.Status === status;
          }).length;
        });

        const totaldeliverables = projectDeliverables.length;
        const completed = projectDeliverables.filter(
          (d: IDeliverableItem) =>
            d.Status === "Approved" || d.Status === "Completed",
        ).length;

        return {
          id: proj.ID,
          title: proj.ProjectName || "Untitled Project",
          type: proj.ProjectType?.ProjectType || "N/A",
          user: proj.PreparedBy?.Title || "Unknown User",
          statusCounts,
          description: proj.ProjectOverview || "No description available",
          documents: projectDeliverables.length,
          comments: projectComments.length,
          progress: {
            current: completed,
            total: totaldeliverables > 0 ? totaldeliverables : 0,
          },
          status: proj.Status || "",
        };
      });

      setCards(projectCards);
      setFilteredCards(projectCards);
    } catch (err) {
      console.error("Error loading dashboard data:", err);
      await ErrorLogger.logError(
        sp,
        err,
        "loadDashboardData",
        "DashboardWebPart",
      );
    } finally {
      setLoading(false);
    }
  };

  const handleOpenAuditHistory = async (deliverableId: number) => {
    try {
      setLoading(true);

      const filteredApprovals = approvalItems.filter(
        (approval) => approval.DeliverablesDetailsIdId === deliverableId,
      );

      console.log(
        `Audit history for deliverable ${deliverableId}:`,
        filteredApprovals,
      );
      setAuditHistoryData(filteredApprovals);
      setShowAuditHistory(true);
    } catch (error) {
      console.error("Error loading audit history:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "handleOpenAuditHistory",
        "DashboardWebPart",
      );
      alert("Error loading audit history. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleCloseAuditHistory = () => {
    setShowAuditHistory(false);
    setAuditHistoryData([]);
  };

  const handleOpenHierarchy = async (deliverableId: number) => {
    console.log("handleOpenHierarchy called with ID:", deliverableId);

    try {
      console.log("Setting hierarchy loading state...");
      setHierarchyLoading(true);
      setSelectedHierarchyDeliverableId(deliverableId);

      console.log("Fetching hierarchy data from SharePoint...");

      const hierarchyData: IApprovalHierarchyItem[] = await sp.web.lists
        .getByTitle("ApprovalHierarchy")
        .items.select("*,ID,DeliverablesDetailsIdId,AssignedTo/Title")
        .expand("AssignedTo")
        .filter(`DeliverablesDetailsIdId eq ${deliverableId}`)
        .orderBy("SerialNumber", true)()
        .catch((error: any) => {
          console.error("SharePoint query error:", error);
          throw error;
        });

      console.log("Hierarchy data received:", hierarchyData);
      setApprovalHierarchyData(hierarchyData);
      setShowApprovalHierarchy(true);
      console.log("Modal should be visible now");
    } catch (error) {
      console.error("Error loading approval hierarchy:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "handleOpenHierarchy",
        "DashboardWebPart",
      );
      alert("Error loading approval hierarchy. Please try again.");
    } finally {
      console.log("Setting hierarchy loading to false");
      setHierarchyLoading(false);
    }
  };

  const handleCloseApprovalHierarchy = () => {
    setShowApprovalHierarchy(false);
    setSelectedHierarchyDeliverableId(null);
    setApprovalHierarchyData([]);
  };

  const getDocumentsByMetadata = async (
    projectId: number,
    deliverableId: number,
  ): Promise<
    Array<{
      fileName: string;
      fileUrl: string;
      documentId: number;
      revision: number;
      isConsolidatorCopy: string;
    }>
  > => {
    try {
      const sp = spfi().using(SPFx(props.context));
      const items = await sp.web.lists
        .getByTitle("DeliverablesDocument")
        .items.select(
          "Id",
          "FileLeafRef",
          "FileRef",
          "ProjectID",
          "DeliverablesDetailsId",
          "FSObjType",
          "Revision",
          "IsConsolidatorCopy",
        )
        .filter(
          `ProjectID eq '${projectId}' and DeliverablesDetailsId eq '${deliverableId}' and FSObjType eq 0`,
        )
        .orderBy("Modified", false)();

      if (items && items.length > 0) {
        return items.map((doc) => ({
          fileName: doc.FileLeafRef,
          fileUrl: doc.FileRef,
          documentId: doc.Id,
          revision: parseInt(doc.Revision) || 0, // Parse revision number
          isConsolidatorCopy: doc.IsConsolidatorCopy || "No",
        }));
      }

      return [];
    } catch (error) {
      console.error("Error fetching documents by metadata:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "getDocumentsByMetadata",
        "DashboardWebPart",
      );
      return [];
    }
  };

  // Updated subfolder function - also try to get Revision from list item
  const getDocumentsFromSubfolder = async (
    projectName: string,
    docNumber: string,
  ): Promise<
    Array<{
      fileName: string;
      fileUrl: string;
      documentId: number;
      revision: number;
      isConsolidatorCopy: string;
    }>
  > => {
    try {
      const sp = spfi().using(SPFx(props.context));
      const libraryName = "DeliverablesDocument";

      const sanitizedProjectName = projectName.replace(/[<>:"/\\|?*]/g, "_");
      const safeDocNumber = docNumber.replace(/[<>:"/\\|?*]/g, "_");
      const folderPath = `${libraryName}/${sanitizedProjectName}/${safeDocNumber}`;

      try {
        const folderItems = (await sp.web
          .getFolderByServerRelativePath(folderPath)
          .files.select(
            "Name",
            "ServerRelativeUrl",
            "TimeLastModified",
            "ListItemAllFields/Id",
            "ListItemAllFields/Revision",
            "ListItemAllFields/IsConsolidatorCopy",
          )
          .expand("ListItemAllFields")
          .orderBy("TimeLastModified", false)()) as Array<{
          Name: string;
          ServerRelativeUrl: string;
          ListItemAllFields?: {
            Id?: number;
            Revision?: string;
            IsConsolidatorCopy?: string;
          };
        }>;

        if (folderItems && folderItems.length > 0) {
          return folderItems
            .filter((file) => file.Name && !file.Name.endsWith("/"))
            .map((file) => ({
              fileName: file.Name,
              fileUrl: file.ServerRelativeUrl,
              documentId: file.ListItemAllFields?.Id || 0,
              revision: parseInt(file.ListItemAllFields?.Revision) || 0,
              isConsolidatorCopy:
                file.ListItemAllFields?.IsConsolidatorCopy || "No",
            }));
        }
      } catch (folderError) {
        console.log(`Folder not found or empty: ${folderPath}`);
      }

      return [];
    } catch (error) {
      console.error("Error fetching documents from subfolder:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "getDocumentsFromSubfolder",
        "DashboardWebPart",
      );
      return [];
    }
  };

  // DocumentCell component with proper revision handling
  const DocumentCell: React.FC<{
    deliverableId: number;
    docNumber: string;
    projectId: number;
    projectName: string;
  }> = ({ deliverableId, docNumber, projectId, projectName }) => {
    const [documents, setDocuments] = React.useState<
      Array<{
        fileName: string;
        fileUrl: string;
        documentId: number;
        revision: number;
        isConsolidatorCopy: string;
      }>
    >([]);
    const [loading, setLoading] = React.useState(false);

    React.useEffect(() => {
      const loadDocuments = async () => {
        if (!projectName || !docNumber) {
          setDocuments([]);
          return;
        }

        setLoading(true);
        try {
          let docs = await getDocumentsByMetadata(projectId, deliverableId);

          if (docs.length === 0) {
            docs = await getDocumentsFromSubfolder(projectName, docNumber);
          }

          const latestDocs = getLatestRevisionDocuments(docs);
          setDocuments(latestDocs);
        } catch (error) {
          console.error("Error loading documents:", error);
          await ErrorLogger.logError(
            sp,
            error,
            "DocumentCell.loadDocuments",
            "DashboardWebPart",
          );
          setDocuments([]);
        } finally {
          setLoading(false);
        }
      };

      loadDocuments();
    }, [deliverableId, docNumber, projectId, projectName]);

    const getLatestRevisionDocuments = (
      docs: Array<{
        fileName: string;
        fileUrl: string;
        documentId: number;
        revision: number;
        isConsolidatorCopy: string;
      }>,
    ): Array<{
      fileName: string;
      fileUrl: string;
      documentId: number;
      revision: number;
      isConsolidatorCopy: string;
    }> => {
      if (docs.length === 0) return [];

      // Step 1: Check if ANY document is a consolidator copy
      const consolidatorDocs = docs.filter(
        (d) => d.isConsolidatorCopy === "Yes" || d.isConsolidatorCopy === "yes",
      );

      if (consolidatorDocs.length > 0) {
        // Show ONLY consolidator documents
        return consolidatorDocs;
      }

      // Step 2: No consolidator - find the highest revision
      const maxRevision = Math.max(...docs.map((d) => d.revision));

      // Step 3: Show ALL documents with the highest revision
      const latestDocs = docs.filter((d) => d.revision === maxRevision);

      return latestDocs.sort((a, b) => a.fileName.localeCompare(b.fileName));
    };

    const getCleanDisplayName = (fileName: string): string => {
      // Remove path and get just the filename
      let cleanName = fileName.split("/").pop() || fileName;

      // Remove file extension temporarily
      const extension = cleanName.split(".").pop();
      const nameWithoutExt = cleanName.substring(0, cleanName.lastIndexOf("."));

      // Clean the name part
      let cleanedName = nameWithoutExt
        .replace(/^\d+[_\s-]+/, "") // Remove leading numbers with separators
        //.replace(/[_\s-]+\d+[_\s-]+/g, ' ') // Remove number blocks in middle
        .replace(/\d{6,}/g, "") // Remove long number sequences (timestamps)
        .replace(/-/g, " ") // Replace hyphens with spaces
        .replace(/_ConsolidatorBackup$/i, "") // Remove _ConsolidatorBackup suffix (case insensitive)
        .replace(/\s{2,}/g, " ") // Replace multiple spaces with single space
        .replace(/_+$/, "") // Remove trailing underscores at the end
        .replace(/^_+/, "") // Remove leading underscores at the start
        .trim();

      // Reattach extension
      return extension ? `${cleanedName}.${extension}` : cleanedName;
    };

    const handleDocumentClick = (documentId: number) => {
      const tronUrl = `https://officeindia.sharepoint.com/sites/ESSA/SitePages/PDFViewer.aspx?docId=${documentId}`;
      window.open(tronUrl, "_blank", "noopener,noreferrer");
    };

    if (!docNumber) {
      return <span className={styles.noDocument}>-</span>;
    }

    if (loading) {
      return <span className={styles.loadingText}>Loading...</span>;
    }

    if (documents.length === 0) {
      return <span className={styles.noDocument}>No document</span>;
    }

    return (
      <div className={styles.fileList}>
        {documents.map((doc, index) => (
          <div key={index} className={styles.fileItem}>
            <a
              href="#"
              onClick={(e) => {
                e.preventDefault();
                handleDocumentClick(doc.documentId);
              }}
              className={styles.fileLink}
            >
              📄 {getCleanDisplayName(doc.fileName)}
            </a>
          </div>
        ))}
      </div>
    );
  };

  const handleCardClick = async (projectId: number, status: string) => {
    if (status === "Save as draft") {
      await handleDraftClick(projectId);
      return;
    }
    try {
      setLoading(true);
      const sp = spfi().using(SPFx(props.context));

      // Get project details
      const projectItem: IProjectItem = await sp.web.lists
        .getByTitle("ProjectCreationList")
        .items.getById(projectId)
        .select(
          "ProjectName",
          "ProjectOverview",
          "PreparedBy/Title",
          "ProjectType/ProjectType",
          "ClientName",
          "ProjectStartDate",
          "Status",
        )
        .expand("PreparedBy", "ProjectType")();

      // Get deliverables for this project
      const projectDeliverables = deliverableItems.filter(
        (d) => d.ProjectCreationListIDId === projectId,
      );

      // Get comments for this project
      const projectComments = commentItems.filter(
        (c) => c.ProjectID === projectId,
      );

      // Get unique assigned team members from deliverables
      const uniqueAssignedMembers = new Set<string>();
      projectDeliverables.forEach((deliverable) => {
        if (deliverable.AssignedTo?.Title) {
          uniqueAssignedMembers.add(deliverable.AssignedTo.Title);
        }
      });

      const projectDetails: IProjectDetails = {
        ProjCreationId: projectId || 0,
        ProjectName: projectItem.ProjectName || "Untitled Project",
        ProjectOverview: projectItem.ProjectOverview || "No overview available",
        PreparedBy: projectItem.PreparedBy?.Title || "Unknown",
        ProjectType: projectItem.ProjectType?.ProjectType || "N/A",
        ClientName: projectItem.ClientName || "N/A",
        ProjectStartDate: projectItem.ProjectStartDate
          ? new Date(projectItem.ProjectStartDate).toLocaleDateString("en-GB", {
              day: "2-digit",
              month: "2-digit",
              year: "numeric",
            })
          : "N/A",
        Status: projectItem.Status || "Unknown",
        CompletedTask: projectDeliverables.filter(
          (d) => d.Status === "Approved" || d.Status === "Completed",
        ).length,
        TotalNoComment: projectComments.length,
        Members: uniqueAssignedMembers.size,
        DeliverablesDetailsArr: projectDeliverables,
      };

      setSelectedProject(projectDetails);
      setShowDetails(true);
    } catch (error) {
      console.error("Error loading project details:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "handleCardClick",
        "DashboardWebPart",
      );
      setShowDetails(true);
    } finally {
      setLoading(false);
    }
  };

  const handleDraftClick = async (projectId: number) => {
    try {
      setLoading(true);
      const sp = spfi().using(SPFx(props.context));
      // Step 1: Get project details from ProjectCreationList
      const projectItem = await sp.web.lists
        .getByTitle("ProjectCreationList")
        .items.getById(projectId)
        .select(
          "ID",
          "ProjectName",
          "ProjectOverview",
          "PreparedBy/Title",
          "PreparedBy/Id",
          "ProjectType/ProjectType",
          "ProjectType/Id",
          "ClientName",
          "ProjectStartDate",
          "Status",
          "SubmitStatus",
        )
        .expand("PreparedBy", "ProjectType")();
      // Step 2: Get payload from RequestProcessingQueue
      const queueItems = await sp.web.lists
        .getByTitle("RequestProcessingQueue")
        .items.select("ID", "PayloadJSON", "Status")
        .filter(`ProjectCreationID eq ${projectId}`)
        .orderBy("ID", false)();
      let payloadData = null;
      if (queueItems.length > 0) {
        try {
          payloadData = JSON.parse(queueItems[0].PayloadJSON);
        } catch (e) {
          console.error("Error parsing PayloadJSON:", e);
        }
      }
      // Step 3: Prepare draft data object
      const draftDataObj = {
        projectId: projectItem.ID,
        projectName: projectItem.ProjectName || "",
        projectTypeId: projectItem.ProjectType?.Id || "",
        projectTypeName: projectItem.ProjectType?.ProjectType || "",
        status: projectItem.Status || "",
        submitStatus: projectItem.SubmitStatus || "",
        clientName: projectItem.ClientName || "",
        preparedById: projectItem.PreparedBy?.Id || 0,
        preparedByName: projectItem.PreparedBy?.Title || "",
        startDate: projectItem.ProjectStartDate
          ? new Date(projectItem.ProjectStartDate)
          : null,
        overview: projectItem.ProjectOverview || "",
        // Payload data for deliverables
        deliverables: payloadData?.deliverables || [],
      };
      setDraftPayload(draftDataObj);
      setShowEditDraft(true);
    } catch (error) {
      console.error("Error loading draft data:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "handleDraftClick",
        "DashboardWebPart",
      );
      alert("Error loading draft data. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleBackToDashboard = () => {
    setShowDetails(false);
    setSelectedProject(null);
  };

  const handleViewMDR = () => {
    if (selectedProject?.ProjCreationId) {
      const url = `https://officeindia.sharepoint.com/sites/ESSA/SitePages/ProjectMDR.aspx?ProjectId=${selectedProject.ProjCreationId}`;
      window.open(url, "_blank");
    }
  };

  const getStatusBadgeClass = (status: string) => {
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

  const menuProps: IContextualMenuProps = {
    items: [
      {
        key: "newProject",
        text: "New Project",
        onClick: () => {
          console.log("New Project clicked in menu");
          setShowNewRequest(true);
        },
      },
    ],
  };

  if (loading)
    return (
      <div
        className={styles.loading}
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          alignItems: "center",
          height: "100vh",
           width: "100%",
        }}
      >
        <img
          src={require("../../../CustomAsset/birdloader.gif")}
          className="alignrightl"
          alt="Loading..."
          style={{ width: "90px", height: "90px" }}
        />
        <div>Loading...</div>
      </div>
    );

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
          <div className="container-fluid  paddb">
            <div className={styles.dashboard}>
              {showNewRequest ? (
                <NewRequest
                  context={props.context}
                  onCancel={async () => {
                    setShowNewRequest(false);

                    setTimeout(async () => {
                      const sp = spfi().using(SPFx(props.context));
                      await loadDashboardData(sp);
                    }, 1000);
                  }}
                />
              ) : showEditDraft && draftPayload ? (
                <NewRequest
                  context={props.context}
                  onCancel={() => setShowEditDraft(false)}
                  initialData={draftPayload}
                />
              ) : showDetails && selectedProject ? (
                <div className={styles.projectDetails}>
                  {/* Horizontal Scroll Wrapper for Entire Details View */}
                  <div className={styles.detailsScrollWrapper}>
                    {/* Header */}
                    <div className={styles.pageTitleRow}>
                      <div className={styles.pageTitleBox}>
                        <h4 className={styles.pageTitle}>
                          {selectedProject.ProjectName}
                        </h4>
                        <div className={styles.pageTitleRight}>
                          <div style={{ display: "flex", gap: "10px" }}>
                            <DefaultButton
                              style={{ width: "130px" }}
                              text="View MDR"
                              iconProps={{ iconName: "RedEye" }}
                              onClick={handleViewMDR}
                              className={styles.btnDark}
                            />

                            <DefaultButton
                              style={{ width: "130px" }}
                              text="Back"
                              iconProps={{ iconName: "NavigateBack" }}
                              onClick={handleBackToDashboard}
                              className={styles.backButton}
                            />
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Stats Cards */}
                    <div className={styles.statsRow}>
                      <div className={styles.statsCol}>
                        <div className={styles.statsCard}>
                          <div className={styles.cardBody}>
                            <div className={styles.statsRowInner}>
                              <div className={styles.statsIconCol}>
                                <div
                                  className={`${styles.avatarMd} ${styles.bgSuccess}`}
                                >
                                  {" "}
                                  <img
                                    src={require("../assets/file.png")}
                                    alt="Document(s)"
                                  />
                                </div>
                              </div>
                              <div className={styles.statsTextCol}>
                                <div className={styles.textEnd}>
                                  <h3 className={styles.statsNumber}>
                                    {selectedProject.CompletedTask}
                                  </h3>
                                  <p className={styles.statsLabel}>
                                    Document(s)
                                  </p>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className={styles.statsCol}>
                        <div className={styles.statsCard}>
                          <div className={styles.cardBody}>
                            <div className={styles.statsRowInner}>
                              <div className={styles.statsIconCol}>
                                <div
                                  className={`${styles.avatarMd} ${styles.bgDanger}`}
                                >
                                  <img
                                    src={require("../assets/comment.png")}
                                    alt="Comment(s)"
                                  />
                                </div>
                              </div>
                              <div className={styles.statsTextCol}>
                                <div className={styles.textEnd}>
                                  <h3 className={styles.statsNumber}>
                                    {selectedProject.TotalNoComment}
                                  </h3>
                                  <p className={styles.statsLabel}>
                                    Comment(s)
                                  </p>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className={styles.statsCol}>
                        <div className={styles.statsCard}>
                          <div className={styles.cardBody}>
                            <div className={styles.statsRowInner}>
                              <div className={styles.statsIconCol}>
                                <div
                                  className={`${styles.avatarMd} ${styles.bgWarning}`}
                                >
                                  <img
                                    src={require("../assets/users.png")}
                                    alt="Member(s)"
                                  />
                                </div>
                              </div>
                              <div className={styles.statsTextCol}>
                                <div className={styles.textEnd}>
                                  <h3 className={styles.statsNumber}>
                                    {selectedProject.Members}
                                  </h3>
                                  <p className={styles.statsLabel}>Member(s)</p>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Project Details */}
                    <div className={styles.detailsRow}>
                      <div className={styles.projectInfoCol}>
                        <div className={styles.projectCard}>
                          <div className={styles.cardBody}>
                            <h5 className={styles.sectionTitle}>Overview:</h5>
                            <p className={styles.projectOverview}>
                              {selectedProject.ProjectOverview}
                            </p>

                            <div className={styles.detailsGrid}>
                              <div className={styles.detailItem}>
                                <h5 className={styles.detailTitle}>
                                  Project Date
                                </h5>
                                <p>{selectedProject.ProjectStartDate}</p>
                              </div>

                              <div className={styles.detailItem}>
                                <h5 className={styles.detailTitle}>
                                  Prepared By
                                </h5>
                                <p>{selectedProject.PreparedBy}</p>
                              </div>

                              <div className={styles.detailItem}>
                                <h5 className={styles.detailTitle}>
                                  Project Type
                                </h5>
                                <p>{selectedProject.ProjectType}</p>
                              </div>

                              <div className={styles.detailItem}>
                                <h5 className={styles.detailTitle}>
                                  Client Name
                                </h5>
                                <p>{selectedProject.ClientName}</p>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                      <div className={styles.chartCol}>
                        <div className={styles.chartCard}>
                          <div className={styles.cardBody}>
                            <h4
                              style={{ fontWeight: "bold" }}
                              className="header-title line18 font-8 text-dark newtextdark fw-bold mb-0"
                            >
                              Deliverables by Status
                            </h4>
                            <div className={styles.chartContainer}>
                              {statusChartData.length > 0 ? (
                                <div className={styles.pieChartWrapper}>
                                  <PieChart
                                    data={statusChartData}
                                    width={250}
                                    height={250}
                                  />
                                  <div className={styles.chartLegend}>
                                    {statusChartData.map((item, index) => (
                                      <div
                                        key={index}
                                        className={styles.legendItem}
                                      >
                                        <div
                                          className={styles.legendColor}
                                          style={{
                                            backgroundColor: item.color,
                                          }}
                                        ></div>
                                        <span className={styles.legendText}>
                                          {item.x}: {item.y}
                                        </span>
                                      </div>
                                    ))}
                                  </div>
                                </div>
                              ) : (
                                <div className={styles.chartPlaceholder}>
                                  <div className={styles.chartStats}>
                                    <div className={styles.chartStat}>
                                      <span className={styles.statCount}>
                                        {
                                          selectedProject.DeliverablesDetailsArr.filter(
                                            (d) =>
                                              d.Status === "Approved" ||
                                              d.Status === "Completed",
                                          ).length
                                        }
                                      </span>
                                      <span className={styles.statLabel}>
                                        Completed
                                      </span>
                                    </div>
                                    <div className={styles.chartStat}>
                                      <span className={styles.statCount}>
                                        {
                                          selectedProject.DeliverablesDetailsArr.filter(
                                            (d) => d.Status === "In-Progress",
                                          ).length
                                        }
                                      </span>
                                      <span className={styles.statLabel}>
                                        In Progress
                                      </span>
                                    </div>
                                    <div className={styles.chartStat}>
                                      <span className={styles.statCount}>
                                        {
                                          selectedProject.DeliverablesDetailsArr.filter(
                                            (d) => d.Status === "Pending",
                                          ).length
                                        }
                                      </span>
                                      <span className={styles.statLabel}>
                                        Pending
                                      </span>
                                    </div>
                                  </div>
                                </div>
                              )}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* Deliverables Table */}
                    <div className={styles.deliverablesSection}>
                      <div className={styles.deliverablesCard}>
                        <div className={styles.cardBody}>
                          <h4 className="header-title line18 font-8 text-dark newtextdark fw-bold mb-0">
                            Deliverables Status
                          </h4>

                          <div className={styles.tableContainer}>
                            <table className={styles.deliverablesTable}>
                              <thead>
                                <tr>
                                  <th>Deliverables</th>
                                  <th>Area</th>
                                  <th style={{ minWidth: "100px" }}>
                                    Assigned To
                                  </th>
                                  <th>Document Number</th>
                                  <th>Organization</th>
                                  <th  style={{ minWidth: "100px" }}>Action Date Time</th>
                                  <th style={{ minWidth: "100px" }}>
                                    Revision Number
                                  </th>
                                  <th style={{ minWidth: "100px" }}>
                                    Attachment
                                  </th>
                                  <th style={{ minWidth: "100px" }}>Status</th>
                                  <th>Remark</th>
                                  <th style={{ minWidth: "100px" }}>
                                    Audit History
                                  </th>
                                  <th style={{ minWidth: "100px" }}>
                                    Approval Hierarchy
                                  </th>
                                </tr>
                              </thead>
                              <tbody>
                                {selectedProject.DeliverablesDetailsArr.map(
                                  (deliverable) => (
                                    <tr key={deliverable.ID}>
                                      <td>
                                        {deliverable.Deliverables
                                          ?.Deliverables || ""}
                                      </td>

                                      <td>{deliverable.Area?.Area || ""}</td>

                                      <td style={{ minWidth: "100px" }}>
                                        {deliverable.AssignedTo?.Title}
                                      </td>
                                      <td>{deliverable.DocNumber}</td>

                                      <td>
                                        {deliverable.Organisation
                                          ?.Organisation || ""}
                                      </td>

                                      <td  style={{ minWidth: "100px" }}>
                                        {deliverable.Status !== "Pending" &&
                                        deliverable.Modified
                                          ? new Date(
                                              deliverable.Modified,
                                            ).toLocaleString("en-GB", {
                                              day: "2-digit",
                                              month: "2-digit",
                                              year: "numeric",
                                              hour: "2-digit",
                                              minute: "2-digit",
                                              second: "2-digit",
                                              hour12: true,
                                            })
                                          : ""}
                                      </td>
                                      <td
                                        style={{ minWidth: "100px" }}
                                        className={styles.textCenter}
                                      >
                                        {deliverable.RevisionNumber}
                                      </td>
                                      <td style={{ minWidth: "150px" }}>
                                        <DocumentCell
                                          deliverableId={deliverable.ID}
                                          docNumber={deliverable.DocNumber}
                                          projectId={
                                            selectedProject.ProjCreationId
                                          }
                                          projectName={
                                            selectedProject.ProjectName
                                          }
                                        />
                                      </td>
                                      <td style={{ minWidth: "100px" }}>
                                        <span
                                          className={`${styles.statusBadge} ${getStatusBadgeClass(deliverable.Status)}`}
                                        >
                                          {deliverable.Status}
                                        </span>
                                      </td>
                                      <td>{deliverable.DocumentComments}</td>
                                      <td
                                        style={{ minWidth: "100px" }}
                                        className={styles.textCenter}
                                      >
                                        <button
                                          style={{ minWidth: "40px" }}
                                          type="button"
                                          className={styles.btnOutlineSuccess}
                                          onClick={() =>
                                            handleOpenAuditHistory(
                                              deliverable.ID,
                                            )
                                          }
                                        >
                                          🔍
                                        </button>
                                      </td>
                                      <td
                                        style={{ minWidth: "100px" }}
                                        className={styles.textCenter}
                                      >
                                        <button
                                          style={{ minWidth: "40px" }}
                                          type="button"
                                          className={styles.btnOutlineDanger}
                                          onClick={(e) => {
                                            e.preventDefault();
                                            e.stopPropagation();
                                            handleOpenHierarchy(deliverable.ID);
                                          }}
                                          disabled={loading} // Add disabled state while loading
                                        >
                                          {loading &&
                                          selectedHierarchyDeliverableId ===
                                            deliverable.ID ? (
                                            <span>⏳</span>
                                          ) : (
                                            "👁️"
                                          )}
                                        </button>
                                      </td>
                                    </tr>
                                  ),
                                )}
                              </tbody>
                            </table>
                            {selectedProject.DeliverablesDetailsArr.length ===
                              0 && (
                              <div className={styles.noData}>
                                <p>No deliverables found for this project.</p>
                              </div>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>{" "}
                  {/* End of detailsScrollWrapper */}
                  {/* AUDIT HISTORY MODAL */}
                  {showAuditHistory && (
                    <div className={styles.modalOverlay}>
                      <div className={styles.modalContent}>
                        <div className={styles.modalHeader}>
                          <h3>Audit History</h3>
                          <button
                            type="button"
                            className={styles.closeButton}
                            onClick={handleCloseAuditHistory}
                          >
                            ×
                          </button>
                        </div>

                        <div className={styles.modalBody}>
                          {auditHistoryData.length > 0 ? (
                            <div className={styles.tableContainer}>
                              <table className={styles.auditTable}>
                                <thead>
                                  <tr>
                                    <th>Approval Level</th>
                                    <th>Approval Role</th>
                                    <th>Approver</th>
                                    <th>Approval Date</th>
                                    <th>Status</th>
                                    <th>Remarks</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {auditHistoryData.map((approval) => (
                                    <tr key={approval.ID}>
                                      <td>{approval.Level}</td>
                                      <td>{approval.ApproverRole}</td>
                                      <td>
                                        {approval.AssignedTo?.Title || "N/A"}
                                      </td>
                                      <td>
                                        {approval.ApprovalDate
                                          ? new Date(
                                              approval.ApprovalDate,
                                            ).toLocaleDateString("en-GB")
                                          : "N/A"}
                                      </td>
                                      <td>
                                        <span
                                          className={`${styles.statusBadge} ${getStatusBadgeClass(approval.Status)}`}
                                        >
                                          {approval.Status}
                                        </span>
                                      </td>
                                      <td>
                                        {approval.Remarks || "No comments"}
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          ) : (
                            <div className={styles.noData}>
                              <p>
                                No audit history found for this deliverable.
                              </p>
                            </div>
                          )}
                        </div>

                        <div className={styles.modalFooter}>
                          <button
                            type="button"
                            className={styles.btnDark}
                            onClick={handleCloseAuditHistory}
                          >
                            Close
                          </button>
                        </div>
                      </div>
                    </div>
                  )}
                  {/* Approval Hierarchy Modal */}
                  {showApprovalHierarchy && (
                    <div className={styles.modalOverlay}>
                      <div className={styles.modalContent}>
                        <div className={styles.modalHeader}>
                          <h3>Approval Hierarchy</h3>
                          <button
                            type="button"
                            className={styles.closeButton}
                            onClick={handleCloseApprovalHierarchy}
                          >
                            ×
                          </button>
                        </div>

                        <div className={styles.modalBody}>
                          <div className={styles.auditInfo}>
                            <h4>
                              Deliverable ID: {selectedHierarchyDeliverableId}
                            </h4>
                            <p>
                              Total Approval Steps:{" "}
                              {approvalHierarchyData.length}
                            </p>
                          </div>

                          {approvalHierarchyData.length > 0 ? (
                            <div className={styles.tableContainer}>
                              <table className={styles.auditTable}>
                                <thead>
                                  <tr>
                                    <th>Approval Level</th>
                                    <th>Approval Role</th>
                                    <th>Approver</th>
                                    {/* <th>Status</th> */}
                                  </tr>
                                </thead>
                                <tbody>
                                  {approvalHierarchyData.map((hierarchy) => (
                                    <tr key={hierarchy.ID}>
                                      <td className={styles.textCenter}>
                                        {hierarchy.Level}
                                      </td>
                                      <td>{hierarchy.ApproverRole}</td>
                                      <td>
                                        {hierarchy.AssignedTo &&
                                        hierarchy.AssignedTo.length > 0 ? (
                                          <div className={styles.userList}>
                                            {hierarchy.AssignedTo.map(
                                              (user, index) => (
                                                <div
                                                  key={index}
                                                  className={styles.userItem}
                                                >
                                                  {user.Title}
                                                  {index <
                                                    hierarchy.AssignedTo
                                                      .length -
                                                      1 && ", "}
                                                </div>
                                              ),
                                            )}
                                          </div>
                                        ) : (
                                          "N/A"
                                        )}
                                      </td>
                                      {/* <td>
                                        <span
                                          className={`${styles.statusBadge} ${getStatusBadgeClass(hierarchy.Status)}`}
                                        >
                                          {hierarchy.Status}
                                        </span>
                                      </td> */}
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          ) : (
                            <div className={styles.noData}>
                              <p>
                                No approval hierarchy found for this
                                deliverable.
                              </p>
                            </div>
                          )}
                        </div>

                        <div className={styles.modalFooter}>
                          <button
                            type="button"
                            className={styles.btnDark}
                            onClick={handleCloseApprovalHierarchy}
                          >
                            Close
                          </button>
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              ) : (
                <>
                  <div className={styles.controls}>
                    <h2
                      style={{ margin: "0px", fontWeight: "600" }}
                      className="fw-bold text-dark header-title"
                    >
                      Dashboard
                    </h2>

                    <div className={styles.searchArea}>
                      <input
                        type="text"
                        placeholder="Search projects..."
                        className={styles.searchInput}
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                      />
                      {searchTerm && (
                        <button
                          //className={styles.clearButton}
                          onClick={() => setSearchTerm("")}
                          title="Clear search"
                        >
                          ✕
                        </button>
                      )}
                      <button
                        className={styles.filterButton}
                        onClick={() => setSearchTerm("")}
                        type="button"
                      >
                        {searchTerm ? "Clear" : "All"}
                      </button>
                      <DefaultButton
                        text="New Request"
                        menuProps={menuProps}
                        className={styles.newRequest}
                      />
                    </div>
                  </div>

                  <div className={styles.cardsGrid}>
                    {filteredCards.length > 0 ? (
                      filteredCards.map((card, index) => (
                        <div
                          key={index}
                          className={styles.card}
                          onClick={() => handleCardClick(card.id, card.status)}
                          style={{ cursor: "pointer" }}
                        >
                          <div className={styles.cardHeader}>
                            <div
                              style={{
                                display: "flex",
                                justifyContent: "space-between",
                                alignItems: "center",
                                width: "100%",
                              }}
                            >
                              <h3
                                style={{
                                  margin: 0,
                                  fontSize: "14px",
                                  fontWeight: "600",
                                }}
                              >
                                {card.title}
                              </h3>
                              {card.status === "Save as draft" && (
                                <span className={styles.draftBadge}>Draft</span>
                              )}
                            </div>
                            <span>...</span>
                          </div>
                          <span className={styles.type}>{card.type}</span>
                          <div className={styles.userInfo}>
                            <span>
                              <img
                                src={require("../assets/usernew.png")}
                                alt="Users"
                              />{" "}
                              {card.user}
                            </span>
                          </div>
                          <div className={styles.statusCounts}>
                            {card.statusCounts.map((count, idx) => {
                              const statusLabels = [
                                "Pending",
                                "In Progress",
                                "Rework",
                                "Approved",
                                "Reject",
                              ];

                              return (
                                <span
                                  key={idx}
                                  className={styles.statusBadge}
                                  title={statusLabels[idx]}
                                >
                                  {count}
                                </span>
                              );
                            })}
                          </div>
                          <p className={styles.description}>
                            {card.description}
                          </p>
                          <div className={styles.stats}>
                            <span>
                              <img
                                style={{ marginTop: "-3px" }}
                                src={require("../assets/docnew.png")}
                                alt="Users"
                              />{" "}
                              {card.documents} Documents
                            </span>
                            <span>
                              <img
                                src={require("../assets/commnew.png")}
                                alt="Users"
                              />{" "}
                              {card.comments} Comments
                            </span>
                          </div>
                          <div className={styles.progress}>
                            <div className="d-flex align-items-center justify-content-between">
                              {" "}
                              <span> Documents</span>
                              <span>
                                {card.progress.current}/{card.progress.total}
                              </span>
                            </div>
                            <div className={styles.progressBar}>
                              <div
                                className={styles.progressFill}
                                style={{
                                  width: `${(card.progress.current / card.progress.total) * 100}%`,
                                }}
                              ></div>
                            </div>
                          </div>
                        </div>
                      ))
                    ) : (
                      <div className={styles.noData}>
                        <p>No projects found.</p>
                      </div>
                    )}
                  </div>
                </>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

const MyProjects: React.FC<IMyProjectProps> = (props) => (
  <Provider>
    <Dashboard props={props} />
  </Provider>
);

export default MyProjects;
