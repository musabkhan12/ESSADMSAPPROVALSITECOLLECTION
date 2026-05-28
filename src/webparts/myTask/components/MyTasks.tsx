import * as React from "react";
import styles from "./MyTasks.module.scss";
import { PrimaryButton, DefaultButton } from "@fluentui/react";

import { WebPartContext } from "@microsoft/sp-webpart-base";

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/files/web";
import CustomPopup from "../../myProject/components/CustomPopup";
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
//import CustomBreadcrumb from '../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb';
import UserContext from "../../../GlobalContext/context";
import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
import { SPFI } from "@pnp/sp/presets/all";
import { getSP } from "../loc/pnpjsConfig";
import loaderGif from "../assets/Loder.gif";
import totalim from "../assets/total.png";
import approve from "../assets/approve.png";
import pen from "../assets/pen.png";
import newlogo from "../assets/logo-high.png";
import eye from "../assets/eye.png";
import { ErrorLogger } from "../../../utils/ErrorLogger";
import { encryptParams } from "../../../utils/CryptoUtils";

interface IMyTaskProps {
  context: WebPartContext;
  siteUrl: string;
  // description: string;
  // siteUrl: string;
  // userDisplayName: string;
  // isDarkTheme: boolean;
  // environmentMessage: string;
  // hasTeamsContext: boolean;
}

type ActionType = "Submitted";

const actionConfig: Record<ActionType, { subject: string; message: string }> = {
  Submitted: {
    subject: "Approval Action Assigned",
    message:
      "The assigned task has been completed and the document has been submitted for your review. Kindly review the submitted document and proceed with the next steps.",
  },
};

// Represents a single task item shown in UI
interface Task {
  projectId: any;
  sno: number;
  projectName: string;
  projectType: string;
  deliverable: string;
  deliverableId: number;
  area: string;
  docType: string;
  docNumber: string;
  revisionnumber: string;
  assignedTo: string;
  preparedBy: string;
  preparedByEmail?: string;
  IsReworked: string;
  org: string;
  status: "Pending" | "Approved" | "In-Progress";
  clientName: string;
  creationDate: Date | undefined;
  DueDate?: Date | undefined;
}

// Audit history for approvals
interface AuditHistoryItem {
  sno: number;
  approvalLevel: string;
  assignedTo: string;
  assignedToRole: string;
  requestorName: string;
  requestedDate: string;
  actionTakenBy: string;
  actionTakenOn: string;
  remark: string;
  status: string;
}

// Comments on documents
interface DocumentComment {
  id: number;
  userName: string;
  commentDate: string;
  pageNumber: string;
  revision: string;
  comment: string;
}

// Uploaded file structure
interface UploadedFile {
  name: string;
  file: File | null;
  url?: string;
  id?: number;
}

// Main component handling My Tasks page
const MyTask = ({ props }: any) => {
  const elementRef = React.useRef<HTMLDivElement>(null);
  const sp: SPFI = getSP();
  const siteUrl = props.siteUrl;
  const { useHide }: any = React.useContext(UserContext);

  // Stores all tasks fetched from SharePoint
  const [tasks, setTasks] = React.useState<Task[]>([]);

  // Current selected filter (Pending / Approved)
  const [currentFilter, setCurrentFilter] = React.useState<string>("Pending");

  // Filtered tasks shown in UI
  const [filteredTasks, setFilteredTasks] = React.useState<Task[]>([]);

  // Date filters
  const [fromDate, setFromDate] = React.useState<string>("");
  const [toDate, setToDate] = React.useState<string>("");

  // Selected task when user clicks view
  const [selectedTask, setSelectedTask] = React.useState<Task | null>(null);

  // Controls form visibility
  const [showForm, setShowForm] = React.useState<boolean>(false);
  const [auditHistory, setAuditHistory] = React.useState<AuditHistoryItem[]>(
    [],
  );
  const [showNoAuditHistory, setShowNoAuditHistory] =
    React.useState<boolean>(false);
  //const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const [selectedFiles, setSelectedFiles] = React.useState<UploadedFile[]>([]);
  const fileInputRef = React.useRef<HTMLInputElement | null>(null);
  const [fileError, setFileError] = React.useState<boolean>(false);
  const [showSubmitLoader, setShowSubmitLoader] =
    React.useState<boolean>(false);

  const [comment, setComment] = React.useState<string>("");
  // const [documentControllerId, setDocumentControllerId] = React.useState<
  //   number | null
  // >(null);
  // const [documentControllerName, setDocumentControllerName] = React.useState<
  //   string | null
  // >(null);
  const [uploadedFileName, setUploadedFileName] = React.useState<string>("");
  const [uploadedFileUrl, setUploadedFileUrl] = React.useState<string>("");
  const [documentComments, setDocumentComments] = React.useState<
    DocumentComment[]
  >([]);
  const [allDocumentComments, setAllDocumentComments] = React.useState<
    DocumentComment[]
  >([]);
  const [versionList, setVersionList] = React.useState<string[]>([]);
  const [selectedVersion, setSelectedVersion] = React.useState<string>("");
  const [showDocumentComments, setShowDocumentComments] =
    React.useState<boolean>(false);

  const [existingDocuments, setExistingDocuments] = React.useState<
    Array<{
      name: string;
      url: string;
      id: number;
      revision?: string; // Optional
      created?: string; // Optional
    }>
  >([]);

  const [popup, setPopup] = React.useState<{
    isOpen: boolean;
    type: "confirmation" | "validation" | "success" | "error";
    title: string;
    message: string;
    onConfirm?: () => void;
  }>({
    isOpen: false,
    type: "success",
    title: "",
    message: "",
  });

  const spContext = spfi().using(SPFx(props.context));

  // const getDocumentController = async (): Promise<number | null> => {
  //   try {
  //     const items = await spContext.web.lists
  //       .getByTitle("ProjectConfiguration")
  //       .items.select(
  //         "*",
  //         "DocumentController/Title",
  //         "DocumentController/EMail",
  //         "DocumentController/ID",
  //       )
  //       .expand("DocumentController")
  //       .orderBy("Created", false)();

  //     if (items.length > 0) {
  //       const documentControllerId = items[0].DocumentControllerId;
  //       const documentControllerName = items[0].DocumentController?.Title;
  //       console.log("Document Controller ID:", documentControllerId);
  //       return documentControllerId;
  //     }
  //     return null;
  //   } catch (error) {
  //     console.error("Error fetching Document Controller:", error);

  //     await ErrorLogger.logError(
  //       spContext,
  //       error,
  //       "getDocumentController",
  //       "MyTaskWebPart",
  //     );

  //     return null;
  //   }
  // };
  // const getDocumentControllerName = async (): Promise<string | null> => {
  //   try {
  //     const items = await spContext.web.lists
  //       .getByTitle("ProjectConfiguration")
  //       .items.select(
  //         "*",
  //         "DocumentController/Title",
  //         "DocumentController/EMail",
  //         "DocumentController/ID",
  //       )
  //       .expand("DocumentController")
  //       .orderBy("Created", false)();

  //     if (items.length > 0) {
  //       const documentControllerName = items[0].DocumentController?.Title;
  //       console.log("Document Controller ID:", documentControllerName);
  //       return documentControllerName;
  //     }
  //     return null;
  //   } catch (error) {
  //     console.error("Error fetching Document Controller:", error);

  //     await ErrorLogger.logError(
  //       spContext,
  //       error,
  //       "getDocumentControllerName",
  //       "MyTaskWebPart",
  //     );

  //     return null;
  //   }
  // };

  // Runs once on component load

  const getDocumentControllerByDocType = async (docType: string) => {
    try {
      const items = await spContext.web.lists
        .getByTitle("ProjectConfiguration")
        .items.select(
          "*",
          "DocumentController/Title",
          "DocumentController/EMail",
          "DocumentController/ID",
          "DocumentType/DocumentType",
        )
        .expand("DocumentController", "DocumentType")
        .filter(`DocumentType/DocumentType eq '${docType}'`)();

      if (items.length > 0) {
        return {
          id: items[0].DocumentControllerId,
          name: items[0].DocumentController?.Title,
        };
      }

      return null;
    } catch (error) {
      console.error("Error fetching controller:", error);

      await ErrorLogger.logError(
        spContext,
        error,
        "getDocumentControllerByDocType",
        "MyTaskWebPart",
      );

      return null;
    }
  };
  React.useEffect(() => {
    const initializeData = async () => {
      try {
        await fetchTasks();
      } catch (error) {
        await ErrorLogger.logError(
          spContext,
          error,
          "initializeData",
          "MyTaskWebPart",
        );
      }

      setTimeout(() => {
        const menu = document.querySelector(".sidebar-openBa");
        if (menu instanceof HTMLElement) menu.style.display = "none";

        const search = document.querySelector(".search_bar");
        if (search instanceof HTMLElement) search.style.display = "none";
      }, 1000);
    };

    initializeData();
  }, []);

  // Fetch tasks assigned to current user from SharePoint
  const fetchTasks = async () => {
    const currentUserId = (await spContext.web.currentUser()).Id;
    try {
      const items = await spContext.web.lists
        .getByTitle("DeliverablesDetails")
        .items.select(
          "*",
          "Area/Area",
          "ProjectCreationListID/ID",
          "ProjectCreationListID/ProjectName",
          "Deliverables/Deliverables",
          "DocumentType/DocumentType",
          "Organisation/Organisation",

          "AssignedTo/Id",
          "AssignedTo/Title",
        )
        .expand(
          "Area",
          "ProjectCreationListID",
          "AssignedTo",
          "Deliverables",
          "DocumentType",
          "Organisation",
        )
        .filter(`AssignedTo/Id eq ${currentUserId}`)();

      const transformedTasks: Task[] = await Promise.all(
        items.map(async (item: any, index: number) => {
          try {
            const creationitem = await spContext.web.lists
              .getByTitle("ProjectCreationList")
              .items.getById(item.ProjectCreationListID.ID)
              .select(
                "*",
                "ProjectType/ProjectType",
                "ProjectType/Id",
                "PreparedBy/Id",
                "PreparedBy/Title",
                "PreparedBy/EMail",
              )
              .expand("ProjectType", "PreparedBy")();
            console.log("Creation Item:", creationitem);
            console.log("Creation Item PreparedBy:", creationitem.PreparedBy);

            return {
              sno: index + 1,
              projectName: item.ProjectCreationListID?.ProjectName || "",
              projectType: creationitem.ProjectType?.ProjectType || "",
              deliverable: item.Deliverables?.Deliverables || "",
              docType: item.DocumentType?.DocumentType || "",
              org: item.Organisation?.Organisation || "",
              deliverableId: item.Id || 0,
              area: item.Area?.Area || "",

              docNumber: item.DocNumber || "",
              DueDate: item.DueDate || undefined,
              assignedTo: item.AssignedTo?.Title || "",
              IsReworked: item.IsReworked || "",
              preparedBy: creationitem.PreparedBy?.Title || "",
              preparedByEmail: creationitem.PreparedBy?.EMail || "",

              status: item.Status,
              projectId: item.ProjectCreationListID?.ID,
              revisionnumber: item.RevisionNumber || "0",
              creationDate: creationitem.Created || undefined,
              clientName: creationitem.ClientName || "",
            };
          } catch (error) {
            console.error(
              `Error fetching creation item for project ${item.ProjectCreationListID?.ID}:`,
              error,
            );

            await ErrorLogger.logError(
              spContext,
              error,
              "fetchTasks_inner_map",
              "MyTaskWebPart",
            );

            return {
              sno: index + 1,
              projectName: item.ProjectCreationListID?.ProjectName || "",
              projectType: "Type A",
              deliverable: item.Deliverables || "",
              deliverableId: item.Id || 0,
              area: item.Area?.Area || "",
              docType: item.DocumentType || "",
              docNumber: item.DocNumber || "",
              assignedTo: item.AssignedTo?.Title || "",
              IsReworked: item.IsReworked || "",
              preparedBy: "",
              preparedByEmail: "",
              org: item.Organization || "",
              status: item.Status,
              projectId: item.ProjectCreationListID?.ID,
              revisionnumber: item.RevisionNumber || "0",
              creationDate: undefined as Date | undefined,
              clientName: "",
            };
          }
        }),
      );

      setTasks(transformedTasks);
      const pendingTasks = transformedTasks.filter(
        (task) => task.status === "Pending",
      );
      setFilteredTasks(pendingTasks);
    } catch (error) {
      console.error("Error loading tasks:", error);

      await ErrorLogger.logError(
        spContext,
        error,
        "fetchTasks",
        "MyTaskWebPart",
      );
    }
  };

  const getDocumentComments = async (
    projectCreationID: number,
    deliverableDetailsID: number,
    revision: string,
  ) => {
    try {
      const items = await spContext.web.lists
        .getByTitle("DocumentComments")
        .items.select("*, ProjectCreationListID/Id, DeliverableDetailsID/Id")
        .expand("ProjectCreationListID", "DeliverableDetailsID")
        .filter(
          `ProjectCreationListID/Id eq ${projectCreationID} and DeliverableDetailsID/Id eq ${deliverableDetailsID}`,
        )
        .orderBy("ID", false)();
      if (items.length > 0) {
        const comments: DocumentComment[] = items.map((item: any) => ({
          id: item.Id,
          userName: item.UserName || "",
          commentDate: item.CommentDate
            ? new Date(item.CommentDate).toLocaleString("en-GB")
            : "",
          pageNumber: item.PageNumber || "",
          revision: item.Revision || "",
          comment: item.Comment || "",
        }));
        setAllDocumentComments(comments);
        console.log("All Document Comments:", comments);

        const filteredComments = comments.filter(
          (comment) => comment.revision === revision,
        );
        setDocumentComments(filteredComments);

        const uniqueVersions = [
          ...new Set(comments.map((comment) => comment.revision)),
        ].sort();
        setVersionList(uniqueVersions);
        setSelectedVersion(revision);
        setShowDocumentComments(true);
        console.log("Unique Versions:", uniqueVersions);
        console.log("Filtered Document Comments:", filteredComments);
      } else {
        setShowDocumentComments(false);
        setDocumentComments([]);
        setAllDocumentComments([]);
        setVersionList([]);
      }
    } catch (error) {
      console.error("Error fetching document comments:", error);

      await ErrorLogger.logError(
        spContext,
        error,
        "getDocumentComments",
        "MyTaskWebPart",
      );

      setShowDocumentComments(false);
    }
  };

  const onVersionChange = (version: string) => {
    try {
      setSelectedVersion(version);
      if (!version) {
        setDocumentComments(allDocumentComments);
      } else {
        const filteredComments = allDocumentComments.filter(
          (comment) => comment.revision === version,
        );
        setDocumentComments(filteredComments);
      }
    } catch (error) {
      console.error("onVersionChange error:", error);

      ErrorLogger.logError(
        spContext,
        error,
        "onVersionChange",
        "MyTaskWebPart",
      );
    }
  };

  const exportCommentsToExcel = () => {
    try {
      const headers = [
        "Users",
        "Comment Date",
        "Page No.",
        "Revision",
        "Comments/Clarifications",
      ];
      const csvContent = [
        headers.join(","),
        ...documentComments.map((comment) =>
          [
            `"${comment.userName}"`,
            `"${comment.commentDate}"`,
            `"${comment.pageNumber}"`,
            `"${comment.revision}"`,
            `"${comment.comment}"`,
          ].join(","),
        ),
      ].join("\n");
      const blob = new Blob([csvContent], { type: "text/csv" });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `DocumentComments_${selectedTask?.docNumber || "export"}.csv`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error("exportCommentsToExcel error:", error);

      ErrorLogger.logError(
        spContext,
        error,
        "exportCommentsToExcel",
        "MyTaskWebPart",
      );
    }
  };

  const refreshDocComment = () => {
    try {
      if (selectedTask) {
        getDocumentComments(
          selectedTask.projectId,
          selectedTask.deliverableId,
          selectedTask.revisionnumber,
        );
      }
    } catch (error) {
      ErrorLogger.logError(
        spContext,
        error,
        "refreshDocComment",
        "MyTaskWebPart",
      );
    }
  };

  const pendingCount = tasks.filter((t) => t.status === "Pending").length;
  const completedCount = tasks.filter(
    (t) => t.status === "Approved" || t.status === "In-Progress",
  ).length;
  const TotalDocumentCount = tasks.length;

  const getStatusClass = (status: string) => {
    switch (status) {
      case "Pending":
        return styles.pendingStatus;
      case "In-Progress":
        return styles.inProgressStatus;
      case "Approved":
        return styles.completedStatus;
      default:
        return styles.pendingStatus;
    }
  };

  const getDeliverablesDetails = (status: string) => {
    try {
      setCurrentFilter(status);
      if (status === "All") {
        setFilteredTasks(tasks);
      } else if (status === "Completed" || status === "Approved") {
        setFilteredTasks(
          tasks.filter(
            (task) =>
              task.status === "Approved" || task.status === "In-Progress",
          ),
        );
      } else {
        setFilteredTasks(tasks.filter((task) => task.status === status));
      }
    } catch (error) {
      ErrorLogger.logError(
        spContext,
        error,
        "getDeliverablesDetails",
        "MyTaskWebPart",
      );
    }
  };

  const applyDueDateFilter = () => {
    try {
      let baseTasks: Task[] = [];

      if (currentFilter === "Approved" || currentFilter === "Completed") {
        baseTasks = tasks.filter(
          (task) => task.status === "Approved" || task.status === "In-Progress",
        );
      } else if (currentFilter === "Pending") {
        baseTasks = tasks.filter((task) => task.status === "Pending");
      } else {
        baseTasks = tasks;
      }

      let filtered = baseTasks;

      if (fromDate) {
        const from = new Date(fromDate);
        from.setHours(0, 0, 0, 0);

        filtered = filtered.filter((task) => {
          if (!task.DueDate) return false;
          const due = new Date(task.DueDate);
          due.setHours(0, 0, 0, 0);
          return due >= from;
        });
      }

      if (toDate) {
        const to = new Date(toDate);
        to.setHours(23, 59, 59, 999);

        filtered = filtered.filter((task) => {
          if (!task.DueDate) return false;
          const due = new Date(task.DueDate);
          return due <= to;
        });
      }

      setFilteredTasks(filtered);
    } catch (error) {
      ErrorLogger.logError(
        // ✅ await hata diya
        spContext,
        error,
        "applyDueDateFilter",
        "MyTaskWebPart",
      );
    }
  };

  const clearDueDateFilter = () => {
    try {
      setFromDate("");
      setToDate("");

      if (currentFilter === "Approved" || currentFilter === "Completed") {
        setFilteredTasks(
          tasks.filter(
            (task) =>
              task.status === "Approved" || task.status === "In-Progress",
          ),
        );
      } else if (currentFilter === "Pending") {
        setFilteredTasks(tasks.filter((task) => task.status === "Pending"));
      } else {
        setFilteredTasks(tasks);
      }
    } catch (error) {
      console.error("clearDueDateFilter error:", error);

      ErrorLogger.logError(
        spContext,
        error,
        "clearDueDateFilter",
        "MyTaskWebPart",
      );
    }
  };

  const getAuditHistoryDeliverables = async (
    item: any,
    context: WebPartContext,
  ): Promise<AuditHistoryItem[]> => {
    const auditHistoryArr: AuditHistoryItem[] = [];

    try {
      const items = await spContext.web.lists
        .getByTitle("ProjectApprovals")
        .items.select(
          "*",
          "AssignedTo/Id",
          "AssignedTo/Title",
          "AssignedTo/EMail",
          "DeliverablesDetailsId/Id",
          // "DeliverablesDetailsId/Deliverables",
          "DeliverablesDetailsId/ID", //Vishnu Added
          "Author/Id",
          "Author/Title",
          "Author/EMail",
        )
        .expand("AssignedTo", "DeliverablesDetailsId", "Author")
        .filter(`DeliverablesDetailsId/ID eq ${item.deliverableId}`)();

      if (items.length > 0) {
        items.forEach((auditHistoryItem: any, index: number) => {
          let status = auditHistoryItem.Status;
          if (status && status.toLowerCase() === "pending") {
            status = "Pending";
          }

          auditHistoryArr.push({
            sno: index + 1,
            approvalLevel: auditHistoryItem.ApprovalLevel || `Level ${index}`,
            assignedTo: auditHistoryItem.AssignedTo?.Title || "N/A",
            assignedToRole: auditHistoryItem.ApproverRole || "",
            requestorName:
              auditHistoryItem.Author?.Title ||
              auditHistoryItem.RequestorName ||
              "N/A",
            requestedDate: auditHistoryItem.Created
              ? new Date(auditHistoryItem.Created).toLocaleString("en-GB")
              : "N/A",
            actionTakenBy:
              auditHistoryItem.ActionTakenBy?.Title ||
              auditHistoryItem.ModifiedBy?.Title ||
              "N/A",
            actionTakenOn: auditHistoryItem.Modified
              ? new Date(auditHistoryItem.Modified).toLocaleString("en-GB")
              : "",
            remark: auditHistoryItem.Remarks || "",
            status: status || "Pending",
          });
        });
      }
    } catch (error) {
      console.error("Error fetching audit history:", error);

      await ErrorLogger.logError(
        spContext,
        error,
        "getAuditHistoryDeliverables",
        "MyTaskWebPart",
      );
    }

    return auditHistoryArr;
  };

  const handleViewClick = async (task: Task) => {
    setFileError(false);
    setSelectedTask(task);
    setShowForm(true);
    setUploadedFileName("");
    setUploadedFileUrl("");
    setComment("");
    setSelectedFiles([]);
    setExistingDocuments([]);

    try {
      const allFiles = await spContext.web.lists
        .getByTitle("DeliverablesDocument")
        .items.select(
          "ID",
          "FileLeafRef",
          "File/ServerRelativeUrl",
          "Revision",
          "Created",
        )
        .expand("File")
        .filter(
          `ProjectID eq '${task.projectId}' 
   and DeliverablesDetailsId eq '${task.deliverableId}'`,
        )
        .orderBy("Created", false)();

      if (allFiles.length > 0) {
        // const documents = allFiles.map((file) => {
        //   const serverRelativeUrl = file.File.ServerRelativeUrl;
        //   const fileUrl = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;

        //   return {
        //     name: file.FileLeafRef,
        //     url: fileUrl,
        //     id: file.ID,
        //     revision: file.Revision,
        //     created: file.Created,
        //   };
        // });

        // Step 1: Map files
        const mappedDocs = allFiles
          .filter((file) => file.File && file.File.ServerRelativeUrl) // ✅ SAFETY
          .map((file) => ({
            name: file.FileLeafRef,
            url: file.File.ServerRelativeUrl,
            id: file.ID,
            revision: Number(file.Revision || 0),
            created: new Date(file.Created),
          }));

        // Step 2: Find latest revision
        const maxRevision = Math.max(...mappedDocs.map((d) => d.revision));

        // Step 3: Filter only latest revision
        const latestDocs = mappedDocs.filter((d) => d.revision === maxRevision);

        // Step 4: If multiple files same revision → take latest by created date
        const latestDate = Math.max(
          ...latestDocs.map((d) => d.created.getTime()),
        );

        const finalDocs = latestDocs.filter(
          (d) => d.created.getTime() === latestDate,
        );

        // Step 5: Remove duplicates
        const uniqueMap = new Map();
        finalDocs.forEach((doc) => {
          if (!uniqueMap.has(doc.id)) {
            uniqueMap.set(doc.id, doc);
          }
        });

        const uniqueDocs = Array.from(uniqueMap.values());

        // Step 6: Final documents
        const documents = uniqueDocs.map((doc) => {
          const fileUrl = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(doc.url)}`;

          return {
            name: doc.name.replace(/_ConsolidatorBackup/i, ""), // clean name
            url: fileUrl,
            id: doc.id,
            revision: doc.revision,
            created: doc.created,
          };
        });

        setExistingDocuments(documents);
      }

      // Load Audit History
      const history = await getAuditHistoryDeliverables(task, props.context);
      setAuditHistory(history);
      setShowNoAuditHistory(history.length === 0);

      // Load Document Comments
      await getDocumentComments(
        task.projectId,
        task.deliverableId,
        task.revisionnumber,
      );

      // Load file and comment if applicable
      if (task.status === "Approved" || task.status === "In-Progress") {
        const deliverablesList = spContext.web.lists.getByTitle(
          "DeliverablesDetails",
        );
        const deliverableItem = await deliverablesList.items
          .getById(task.deliverableId)
          .select("DeliverablesDocumentID/ID", "DocumentComments")
          .expand("DeliverablesDocumentID")();

        if (deliverableItem.DocumentComments) {
          setComment(deliverableItem.DocumentComments);
        }

        const deliverablesDocumentId =
          deliverableItem.DeliverablesDocumentID?.ID;
        if (deliverablesDocumentId) {
          const fileItem = await spContext.web.lists
            .getByTitle("DeliverablesDocument")
            .items.getById(deliverablesDocumentId)
            .select("ID", "FileLeafRef", "File/ServerRelativeUrl")
            .expand("File")();

          if (fileItem.File) {
            setUploadedFileName(fileItem.FileLeafRef);
            const serverRelativeUrl = fileItem.File.ServerRelativeUrl;
            const fileUrl = `${props.context.pageContext.web.absoluteUrl}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(serverRelativeUrl)}`;
            setUploadedFileUrl(fileUrl);
          }
        }
      } else {
        setComment("");
        setUploadedFileName("");
        setUploadedFileUrl("");
      }
    } catch (error) {
      console.error("Error loading task details:", error);

      await ErrorLogger.logError(
        spContext,
        error,
        "handleViewClick",
        "MyTaskWebPart",
      );

      setAuditHistory([]);
      setShowNoAuditHistory(true);
      setShowDocumentComments(false);
    }
  };

  const handleBackClick = () => {
    try {
      setSelectedTask(null);
      setShowForm(false);
      setAuditHistory([]);
      setShowNoAuditHistory(false);
      setShowDocumentComments(false);
      setDocumentComments([]);
      setExistingDocuments([]); // Empty array with proper type
    } catch (error) {
      console.error("handleBackClick error:", error);

      ErrorLogger.logError(
        spContext,
        error,
        "handleBackClick",
        "MyTaskWebPart",
      );
    }
  };

  const getSecurePdfUrl = async (documentId: number): Promise<string> => {
    const token = await encryptParams(documentId, "read");

    return `https://officeindia.sharepoint.com/sites/ESSA/SitePages/PDFViewer.aspx?token=${token}`;
  };

  // Upload files to SharePoint document library
  const uploadFileToFolder = async (
    folderServerRelativeUrl: string,
    file: File,
    fileName: string,
    digest: string,
  ) => {
    try {
      const url = `${props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${folderServerRelativeUrl}')/Files/add(url='${fileName}', overwrite=true)`;
      const fileBuffer = await file.arrayBuffer();

      const response = await fetch(url, {
        method: "POST",
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-Type": "application/octet-stream",
          "X-RequestDigest": digest,
        },
        body: fileBuffer,
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(`Upload failed: ${response.status} - ${errorText}`);
      }

      const result = await response.json();
      return result.d.ListItemAllFields.Id;
    } catch (error) {
      await ErrorLogger.logError(
        spContext,
        error,
        "uploadFileToFolder",
        "MyTaskWebPart",
      );

      throw error;
    }
  };

  // Handles final submission of task
  const handleSubmitClick = async () => {
    setShowSubmitLoader(true);

    if (!selectedTask || selectedFiles.length === 0) {
      setFileError(true);
      setPopup({
        isOpen: true,
        type: "validation",
        title: "Validation",
        message: "Please fill mandatory fields.",
      });
      setShowSubmitLoader(false);
      return;
    }

    try {
      const controllerData = await getDocumentControllerByDocType(
        selectedTask.docType,
      );
      console.log("Selected DocType:", selectedTask.docType);
      console.log("Controller Data:", controllerData);

      if (!controllerData) {
        setPopup({
          isOpen: true,
          type: "error",
          title: "Configuration Missing",
          message: `No Document Controller found for ${selectedTask.docType}`,
        });
        setShowSubmitLoader(false);
        return;
      }

      const documentControllerId = controllerData.id;
      const documentControllerName = controllerData.name;
      const currentRevision = parseInt(selectedTask.revisionnumber || "0");
      const newRevision =
        selectedTask.IsReworked === "Yes"
          ? (currentRevision + 1).toString()
          : currentRevision.toString();

      // ---------- SANITIZE NAMES (same as NewRequest.tsx) ----------
      // const sanitize = (name: string) => name.replace(/[<>:"/\\|?*]/g, "_");
      const sanitize = (name: string) =>
        name
          .replace(/[<>:"/\\|?*]/g, "_")
          .replace(/\s+/g, " ")
          .trim();
      const projectName = sanitize(selectedTask.projectName);
      const docNumber = sanitize(selectedTask.docNumber);

      // ---------- CORRECT FOLDER PATH ----------
      const folderPath = `DeliverablesDocument/${projectName}/${docNumber}`;
      console.log("Target folder path (web-relative):", folderPath);

      // ---------- FOLDER CHECK (PnPjs) ----------
      let targetFolder;
      try {
        targetFolder = spContext.web.getFolderByServerRelativePath(folderPath);
        await targetFolder.files(); // test existence
        console.log(" Folder exists");
      } catch (err) {
        setPopup({
          isOpen: true,
          type: "error",
          title: "Folder Missing",
          message: `Subfolder "${docNumber}" does not exist under project "${projectName}".`,
        });
        setShowSubmitLoader(false);
        return;
      }

      // ---------- GET REQUEST DIGEST ----------
      const digestRes = await fetch(
        `${props.context.pageContext.web.absoluteUrl}/_api/contextinfo`,
        {
          method: "POST",
          headers: { Accept: "application/json;odata=verbose" },
        },
      );
      const digestData = await digestRes.json();
      const digest = digestData.d.GetContextWebInformation.FormDigestValue;
      console.log(" Digest obtained");

      // ---------- UPLOAD FILES (REST API - binary) ----------
      const uploadedFileIds: number[] = [];
      // Server-relative path (starts with /sites/ESSA)
      // const serverRelativeFolderPath = `${props.context.pageContext.web.serverRelativeUrl}/${folderPath}`;
      // // Encode each part separately to keep slashes
      // const encodedFolderPath = serverRelativeFolderPath
      //   .split("/")
      //   .map((part) => encodeURIComponent(part))
      //   .join("/");
      // console.log("Encoded server-relative folder path:", encodedFolderPath);
      const serverRelativeFolderPath = `${props.context.pageContext.web.serverRelativeUrl}/${folderPath}`;

      console.log("Server-relative folder path:", serverRelativeFolderPath);

      for (const fileObj of selectedFiles) {
        if (!fileObj.file) continue;
        // const fileName = `${Date.now()}_${fileObj.file.name}`; TimeStampIssueLine
        const fileName = fileObj.file.name; //Timestamp removed
        const fileBuffer = await fileObj.file.arrayBuffer();

        const uploadUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/GetFolderByServerRelativeUrl('${serverRelativeFolderPath}')/Files/add(url='${fileName}', overwrite=true)`;
        console.log("Upload URL:", uploadUrl);

        const uploadRes = await fetch(uploadUrl, {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/octet-stream",
            "X-RequestDigest": digest,
          },
          body: fileBuffer,
        });

        if (!uploadRes.ok) {
          const errText = await uploadRes.text();
          throw new Error(`Upload failed: ${uploadRes.status} - ${errText}`);
        }

        const result = await uploadRes.json();

        //  IMPORTANT FIX START
        console.log("Upload response:", result);

        // Get file URL from response
        const fileServerRelativeUrl = result.d.ServerRelativeUrl;

        // Get correct List Item ID using PnP
        const fileItem = await spContext.web
          .getFileByServerRelativePath(fileServerRelativeUrl)
          .getItem<{ Id: number }>();

        const uploadedFileItemId = fileItem.Id;

        console.log("Correct File Item ID:", uploadedFileItemId);
        //  IMPORTANT FIX END

        // Update metadata
        const updateUrl = `${props.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('DeliverablesDocument')/items(${uploadedFileItemId})`;
        const updateRes = await fetch(updateUrl, {
          method: "POST",
          headers: {
            Accept: "application/json;odata=verbose",
            "Content-Type": "application/json;odata=nometadata",
            "X-RequestDigest": digest,
            "X-HTTP-Method": "MERGE",
            "IF-MATCH": "*",
          },
          body: JSON.stringify({
            ProjectID: selectedTask.projectId.toString(),
            DeliverablesDetailsId: selectedTask.deliverableId.toString(),
            Revision: newRevision,
            IsMyTask: "Yes",
          }),
        });
        if (!updateRes.ok) {
          const err = await updateRes.text();
          console.error("Metadata update failed:", err);
        } else {
          console.log("Metadata updated successfully");
        }

        uploadedFileIds.push(uploadedFileItemId);
        console.log(` Uploaded: ${fileName}`);
      }

      // ---------- REST OF THE CODE (same as before) ----------
      // Update DeliverablesDetails
      await spContext.web.lists
        .getByTitle("DeliverablesDetails")
        .items.getById(selectedTask.deliverableId)
        .update({
          DocumentComments: comment,
          RevisionNumber: newRevision,
          Status: "In-Progress",
          IsReworked: "",
          DeliverablesDocumentIDId:
            uploadedFileIds.length > 0 ? uploadedFileIds : [],
        });
      console.log(" DeliverablesDetails updated");

      // Mark Vendor approval completed
      const currentUser = await spContext.web.currentUser();
      const vendorApprovals = await spContext.web.lists
        .getByTitle("ProjectApprovals")
        .items.select("*")
        .filter(
          `ProjectCreationListIDId eq ${selectedTask.projectId} and DeliverablesDetailsIdId eq ${selectedTask.deliverableId} and AssignedToId eq ${currentUser.Id} and ApproverRole eq 'Vendor' and Status eq 'Pending'`,
        )();
      if (vendorApprovals.length > 0) {
        await spContext.web.lists
          .getByTitle("ProjectApprovals")
          .items.getById(vendorApprovals[0].Id)
          .update({
            Status: "Completed",
            Remarks: comment,
            ApprovalDate: new Date(),
          });
        console.log(" Vendor approval completed");
      }

      // Get Prepared By
      const projectItem = await spContext.web.lists
        .getByTitle("ProjectCreationList")
        .items.getById(selectedTask.projectId)
        .select("PreparedBy/Id", "PreparedBy/Title")
        .expand("PreparedBy")();
      const preparedByUserId = projectItem.PreparedBy?.Id;
      const preparedByUserName = projectItem.PreparedBy?.Title;

      // Create Document Controller approval
      await spContext.web.lists.getByTitle("ProjectApprovals").items.add({
        DeliverablesDetailsIdId: selectedTask.deliverableId,
        ProjectCreationListIDId: selectedTask.projectId,
        AssignedToId: documentControllerId,
        DocumentType: selectedTask.docType,
        ApproverRole: "Document Controller",
        ProjectType: selectedTask.projectType,
        Level: "Level 1",
        SerialNumber: 0,
        ApprovalCriteria: "Anyone",
        RequestedById: currentUser.Id,
        RequestedDate: new Date(),
        IncomingDate: new Date(),
        RequestedRole: "Vendor",
        RevisionNumber: newRevision,
        Status: "Pending",
      });
      console.log(" Document Controller approval created");

      // ---------- EMAIL TRIGGER START ----------
      const fileDetails = [];

      const latestId = uploadedFileIds[uploadedFileIds.length - 1];

      if (latestId) {
        const item = await spContext.web.lists
          .getByTitle("DeliverablesDocument")
          .items.getById(latestId)
          .select("FileLeafRef", "Revision")();

        const originalName = item.FileLeafRef;

        const cleanedName = originalName.includes("_")
          ? originalName.substring(originalName.indexOf("_") + 1)
          : originalName;

        fileDetails.push({
          name: cleanedName,
          link: `https://officeindia.sharepoint.com/sites/ESSA/SitePages/PDFViewer.aspx?docId=${latestId}`,
          docrevisionNo: item.Revision,
        });
      }

      const emailBody = buildApprovalEmailBody(
        "Submitted",
        documentControllerName || "",
        preparedByUserName || "",
        fileDetails.map((d) => ({
          deliverable: selectedTask?.deliverable,
          fileName: d.name,
          sharedLink: d.link,
          docrevisionNo: d.docrevisionNo,
        })),
      );

      await addEmailTriggerDetails(
        selectedTask.projectName,
        documentControllerId!,
        preparedByUserId,
        emailBody,
        uploadedFileIds,
      );
      // ---------- EMAIL TRIGGER END ----------

      // Success
      setShowSubmitLoader(false);
      setPopup({
        isOpen: true,
        type: "success",
        title: "Success",
        message: "Task submitted successfully.",
        onConfirm: () => {
          setPopup((p) => ({ ...p, isOpen: false }));
          handleBackClick();
          window.location.reload();
        },
      });
    } catch (error) {
      console.error("Error in submission:", error);

      await ErrorLogger.logError(
        spContext,
        error,
        "handleSubmitClick",
        "MyTaskWebPart",
      );

      setShowSubmitLoader(false);

      setPopup({
        isOpen: true,
        type: "error",
        title: "Invalid File Name",
        message: "Please remove special characters from file name.",
      });
    }
  };

  // Builds HTML email body for approval notification
  const buildApprovalEmailBody = (
    actionType: ActionType,
    approverName: string,
    senderName: string,
    docs: {
      deliverable: string;
      fileName: string;
      sharedLink: string;
      docrevisionNo: string;
    }[],
  ) => {
    const { message } = actionConfig[actionType];

    // Format current date as dd MMM yyyy
    const today = new Date();
    const sentDate = today.toLocaleDateString("en-GB", {
      day: "numeric",
      month: "short",
      year: "numeric",
    });

    let rows = docs
      .map((d, i) => {
        const no = i + 1;
        const transmittal = `TR-${String(no).padStart(3, "0")}`;
        const remarks = "Submitted for Review";

        // <td style="border:1px solid #ccc;padding:6px;">${selectedTask?.docNumber}</td>
        return `
        <tr>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${no}</td>
          <td style="border:1px solid #ccc;padding:6px;">
            <a href="${d.sharedLink}" target="_blank">${selectedTask?.docNumber}</a>
          </td>
          <td style="border:1px solid #ccc;padding:6px;">${d.deliverable}</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${d.docrevisionNo ?? "-"}</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
          <td style="border:1px solid #ccc;padding:6px;">${transmittal}</td>
          <td style="border:1px solid #ccc;padding:6px;">${sentDate}</td>
          <td style="border:1px solid #ccc;padding:6px;">Attached</td>
        </tr>
      `;
      })
      .join("");

    // --- Extra Dropbox Row ---
    const extraNo = docs.length + 1;
    const extraTransmittal = `TR-${String(extraNo).padStart(3, "0")}`;
    //const dropboxUrl = `https://officeindia.sharepoint.com/sites/ESSA/DeliverablesDocument/Forms/AllItems.aspx?id=%2Fsites%2FESSA%2FDeliverablesDocument%2FBNS%201107%2FBNS%201107%2DX%2DX01%2DCIV%2DMAR%2D0001&viewid=e0f53e58%2D6b2f%2D46c3%2D8c04%2Dae8262124b60`;

    //code Line 1126 to 1134 dynamic foder link creation for dropbox link in email
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const serverRel = props.context.pageContext.web.serverRelativeUrl;

    const projectName = encodeURIComponent(selectedTask?.projectName || "");
    const docNumber = encodeURIComponent(selectedTask?.docNumber || "");

    const dropboxUrl = `${siteUrl}/DeliverablesDocument/Forms/AllItems.aspx?id=${serverRel}%2FDeliverablesDocument%2F${projectName}%2F${docNumber}`;
    rows += `
    <tr>
      <td style="border:1px solid #ccc;padding:6px;text-align:center;">${extraNo}</td>
      <td style="border:1px solid #ccc;padding:6px;"></td>
      <td style="border:1px solid #ccc;padding:6px;">${selectedTask.projectName}</td>
      <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
      <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
      <td style="border:1px solid #ccc;padding:6px;">${extraTransmittal}</td>
      <td style="border:1px solid #ccc;padding:6px;">${sentDate}</td>
      <td style="border:1px solid #ccc;padding:6px;">
        <a href="${dropboxUrl}" target="_blank">Dropbox Link</a>
      </td>
    </tr>
  `;

    return `
      Dear ${approverName},<br/><br/>
      ${message}<br/><br/>
  
      <table style="border-collapse:collapse;width:100%;max-width:900px;font-size:13px;">
        <thead>
          <tr style="background:#8eaada;font-weight:bold;">
            <th style="border:1px solid #ccc;padding:6px;text-align:center;">NO.#</th>
            <th style="border:1px solid #ccc;padding:6px;">Doc. #</th>
            <th style="border:1px solid #ccc;padding:6px;">Deliverables</th>
            <th style="border:1px solid #ccc;padding:6px;text-align:center;">Rev</th>
            <th style="border:1px solid #ccc;padding:6px;text-align:center;">Status</th>
            <th style="border:1px solid #ccc;padding:6px;">Transmittal</th>
            <th style="border:1px solid #ccc;padding:6px;">Sent Dated</th>
            <th style="border:1px solid #ccc;padding:6px;">Remarks</th>
          </tr>
        </thead>
        <tbody>
          ${rows}
        </tbody>
      </table>
  
      <br/><br/>
  
      <a href="https://officeindia.sharepoint.com/sites/ESSA/SitePages/ESSAAPPROVALS.aspx" target="_blank">
        Click here to view in portal
      </a><br/><br/>
  
      Regards,<br/>
      ${senderName}
    `;
  };

  const waitForShareLink = async (
    fileItemId: number,
    maxAttempts = 20,
    delayMs = 3000,
  ) => {
    try {
      for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        const item = await spContext.web.lists
          .getByTitle("DeliverablesDocument")
          .items.getById(fileItemId)
          .select("File/Name", "SharedLink", "Revision")
          .expand("File")();

        if (item.SharedLink) {
          return {
            shareLink: item.SharedLink,
            fileName: item.File?.Name,
            revisionNo: item.Revision,
          };
        }

        // wait before next attempt
        await new Promise((res) => setTimeout(res, delayMs));
      }

      throw new Error(`ShareLink not generated for FileItemId: ${fileItemId}`);
    } catch (error) {
      //

      await ErrorLogger.logError(
        spContext,
        error,
        "waitForShareLink",
        "MyTaskWebPart",
      );

      throw error;
    }
  };

  // 🔹 EMAIL TRIGGER DETAILS (My Task)
  const addEmailTriggerDetails = async (
    projectName: string,
    documentControllerId: number,
    preparedByUserId: number,
    emailbody: string,
    attachmentIds: number[],
  ) => {
    try {
      const currentUser = await spContext.web.currentUser();

      await spContext.web.lists.getByTitle("EmailTriggerDetails").items.add({
        Title: projectName,
        IsMyTask: "Yes",
        Body: emailbody,
        Subject: `Document Submission Confirmation – ${selectedTask?.deliverable}`,

        //  ToUser = Document Controller
        ToUserId: [documentControllerId],

        //  CCUser = Logged-in user + Prepared By
        CCUserId: [currentUser.Id, preparedByUserId],

        DeliverablesDocumentIDId: attachmentIds,
      });
    } catch (error) {
      await ErrorLogger.logError(
        spContext,
        error,
        "addEmailTriggerDetails",
        "MyTaskWebPart",
      );

      console.error("Error in addEmailTriggerDetails:", error);
    }
  };

  return (
    <div id="wrapper" ref={elementRef}>
      {/* 🔄 SUBMIT TASK LOADER */}
      {showSubmitLoader && (
        <div
          style={{
            position: "fixed",
            top: 0,
            left: 0,
            width: "100%",
            height: "100%",
            backgroundColor: "rgba(0,0,0,0.45)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            zIndex: 9999,
          }}
        >
          <div className="text-center">
            <img
              src={require("../../../CustomAsset/birdloader.gif")}
              alt="Submitting..."
              style={{ width: "90px", height: "90px" }}
            />
            <p className="text-white mt-3 fw-semibold">
              Submitting, please wait...
            </p>
          </div>
        </div>
      )}

      <div style={{ display: "none" }} className="app-menu" id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>

      <div className="content-page">
        <HorizontalNavbar
          _context={sp}
          className="app-menu-hide"
          siteUrl={siteUrl}
        />
        <div
          className="content"
          style={{
            marginLeft: `${!useHide ? "15px" : "0px"}`,

            marginTop: "0rem",
          }}
        >
          <div className="container-fluid  paddb">
            <img
              style={{
                position: "fixed",
                top: "8px",
                left: "1px",
                zIndex: "99",
              }}
              src={newlogo}
            />
            <div className={styles.myTask}>
              {!showForm ? (
                <>
                  <div className="row mb-2">
                    <div className="col-md-6">
                      <h4 className="page-title fw-bold mb-1 text-dark font-20">
                        Vendor Portal
                      </h4>
                      <ol className="breadcrumb mb-0">
                        {/* <li className="breadcrumb-item">
                          <a href="">Home</a> </li>
                        <li className="breadcrumb-item active"><a href="#">My Task</a></li> */}
                      </ol>
                    </div>
                  </div>

                  <div className="row nrefield">
                    <div className="col-sm-4">
                      <div className="card">
                        <div className="card-body d-flex justify-content-between align-items-center">
                          <div>
                            <p
                              style={{ fontSize: "20px" }}
                              className="mb-2 text-dark font-20 fw-bold"
                            >
                              Total Documents
                            </p>
                            <h3
                              style={{ fontSize: "16px" }}
                              className="mb-0 font-16 text-dark"
                            >
                              <span data-target="438">
                                {TotalDocumentCount}
                              </span>
                            </h3>
                          </div>
                          <div className="avatar fs-60 avatar-img-size">
                            <img src={totalim} alt="projects" />
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="col-sm-4">
                      <div
                        style={{ width: "100%" }}
                        className={`${styles.tileCard} ${currentFilter === "Approved" ? styles.activeTile : ""}`}
                        onClick={() => getDeliverablesDetails("Approved")}
                      >
                        <div className="card-body d-flex justify-content-between align-items-center">
                          <div>
                            <p
                              style={{ fontSize: "20px" }}
                              className="mb-2 text-dark font-20 fw-bold"
                            >
                              Completed
                            </p>
                            <h3
                              style={{ fontSize: "16px" }}
                              className="mb-0 font-16 text-dark"
                            >
                              <span data-target="438">{completedCount}</span>
                            </h3>
                          </div>
                          <div className="avatar fs-60 avatar-img-size">
                            <img src={approve} alt="projects" />
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="col-sm-4">
                      <div
                        style={{ width: "100%" }}
                        className={`${styles.tileCard} ${currentFilter === "Pending" ? styles.activeTile : ""}`}
                        onClick={() => getDeliverablesDetails("Pending")}
                      >
                        <div className="card-body d-flex justify-content-between align-items-center">
                          <div>
                            <p
                              style={{ fontSize: "20px" }}
                              className="mb-2 text-dark font-20 fw-bold"
                            >
                              Pending
                            </p>
                            <h3
                              style={{ fontSize: "16px" }}
                              className="mb-0 font-16 text-dark"
                            >
                              {pendingCount}
                            </h3>
                          </div>
                          <div className="avatar fs-60 avatar-img-size">
                            <img src={pen} alt="projects" />
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>

                  <div className="card mt-0">
                    <div className="card-body pb-0">
                      <div className="d-flex align-items-center justify-content-between">
                        <div className={styles.tilesContainer}>
                          <div
                            className={`${styles.tileCard} ${currentFilter === "Pending" ? styles.activeTile : ""}`}
                            onClick={() => getDeliverablesDetails("Pending")}
                          >
                            <div className={styles.tileBody}>
                              <div className={styles.tileRow}>
                                <div className={styles.tileContent}>
                                  <p className={styles.tileLabel}>Pending</p>{" "}
                                  <h3 className={styles.tileCount}>
                                    {pendingCount}
                                  </h3>
                                </div>
                              </div>
                            </div>
                          </div>

                          <div
                            className={`${styles.tileCard} ${currentFilter === "Approved" ? styles.activeTile : ""}`}
                            onClick={() => getDeliverablesDetails("Approved")}
                          >
                            <div className={styles.tileBody}>
                              <div className={styles.tileRow}>
                                <div className={styles.tileContent}>
                                  <p className={styles.tileLabel}>Completed</p>
                                  <h3 className={styles.tileCount}>
                                    {completedCount}
                                  </h3>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>

                        <div
                          style={{ marginTop: "-12px" }}
                          className="d-flex align-items-center justify-content-end gap-2"
                        >
                          <div>
                            <div className="form-group d-flex align-items-center justify-content-end gap-1">
                              <label
                                style={{ width: "120px" }}
                                className="mb-1"
                              >
                                From Date
                              </label>
                              <input
                                type="date"
                                className="form-control"
                                value={fromDate}
                                onChange={(e) => setFromDate(e.target.value)}
                              />
                            </div>
                          </div>
                          <div className="form-group d-flex align-items-center justify-content-end gap-1">
                            <label style={{ width: "100px" }} className="mb-1">
                              To Date
                            </label>
                            <input
                              type="date"
                              className="form-control"
                              value={toDate}
                              onChange={(e) => setToDate(e.target.value)}
                            />
                          </div>
                          <div
                            className="btn btn-primary pr7"
                            onClick={applyDueDateFilter}
                          >
                            Filter
                          </div>
                          <div
                            className="btn btn-secondary pr7 ms-2"
                            onClick={clearDueDateFilter}
                          >
                            Clear
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                  {/* Tiles */}

                  {/* Table */}
                  <div
                    style={{ clear: "both", marginTop: "15px" }}
                    className="mt-2"
                  >
                    <div className="">
                      <div className={styles.mainTableContainer}>
                        <div className={styles.tableCard}>
                          <div className={styles.tableWrapper}>
                            <table className={styles.taskTable}>
                              <thead>
                                <tr>
                                  <th
                                    style={{
                                      minWidth: "70px",
                                      textAlign: "center",
                                    }}
                                  >
                                    S.No
                                  </th>
                                  <th style={{ textAlign: "center" }}>
                                    Project Name
                                  </th>
                                  <th>Project Type</th>
                                  <th>Deliverable</th>
                                  <th style={{ textAlign: "center" }}>
                                    Document Type
                                  </th>
                                  <th style={{ textAlign: "center" }}>
                                    Due Date
                                  </th>
                                  <th>Organization</th>
                                  <th
                                    style={{
                                      minWidth: "90px",
                                      textAlign: "center",
                                    }}
                                  >
                                    Status
                                  </th>
                                  <th
                                    style={{
                                      minWidth: "70px",
                                      textAlign: "center",
                                    }}
                                  >
                                    Action
                                  </th>
                                </tr>
                              </thead>
                              <tbody>
                                {filteredTasks.map((task, index) => (
                                  <tr key={task.sno}>
                                    <td style={{ minWidth: "70px" }}>
                                      <div
                                        className="indexdesign"
                                        style={{ marginLeft: "8px" }}
                                      >
                                        {index + 1}{" "}
                                      </div>
                                    </td>
                                    <td style={{ textAlign: "center" }}>
                                      {task.projectName}
                                    </td>
                                    <td>{task.projectType}</td>
                                    <td>{task.deliverable}</td>
                                    <td style={{ textAlign: "center" }}>
                                      {task.docType}
                                    </td>
                                    <td>
                                      {task.DueDate
                                        ? new Date(
                                            task.DueDate,
                                          ).toLocaleDateString("en-GB")
                                        : "-"}
                                    </td>

                                    <td>{task.org}</td>
                                    <td
                                      style={{
                                        minWidth: "90px",
                                        textAlign: "center",
                                      }}
                                    >
                                      <span
                                        className={`${styles.statusBadge} ${getStatusClass(task.status)}`}
                                      >
                                        {task.status}
                                      </span>
                                    </td>
                                    <td
                                      style={{
                                        minWidth: "70px",
                                        textAlign: "center",
                                      }}
                                    >
                                      <img
                                        src={eye}
                                        style={{ cursor: "pointer" }}
                                        onClick={() => handleViewClick(task)}
                                        title={`View ${task.deliverable}`}
                                      />
                                      {/* <span
                                        className={styles.actionIcon}
                                       
                                      >
                                        👁️
                                      </span> */}
                                    </td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                        {/* <div className={styles.scrollHint}><em>Scroll horizontally → to view all columns</em></div> */}
                      </div>
                    </div>
                  </div>
                </>
              ) : (
                <>
                  {/* Form View */}
                  <div className={styles.formContainer}>
                    <div className={styles.formHeader}>
                      <h3 style={{ margin: "0px" }}>
                        My Task &gt;&gt; {selectedTask?.docNumber}
                      </h3>
                      <DefaultButton
                        text="Back"
                        iconProps={{ iconName: "NavigateBack" }}
                        onClick={handleBackClick}
                        className={styles.backButton}
                      />
                    </div>
                    <div className={styles.formView}>
                      <div className={styles.formInner}>
                        <div className={styles.formBody}>
                          <div className={styles.formGrid}>
                            <div>
                              <label>Project Name</label>
                              <input
                                type="text"
                                value={selectedTask?.projectName || ""}
                                disabled
                              />
                            </div>
                            <div>
                              <label>Project Type</label>
                              <input
                                type="text"
                                value={selectedTask?.projectType || ""}
                                disabled
                              />
                            </div>
                            <div>
                              <label>Client Name</label>
                              <input
                                type="text"
                                value={selectedTask?.clientName || ""}
                                disabled
                              />
                            </div>
                            <div>
                              <label>Project Date</label>
                              <input
                                type="text"
                                value={
                                  selectedTask?.creationDate
                                    ? new Date(
                                        selectedTask.creationDate,
                                      ).toLocaleDateString("en-GB")
                                    : ""
                                }
                                disabled
                              />
                            </div>
                            <div>
                              <label>Prepared By</label>
                              <input
                                type="text"
                                value={selectedTask?.preparedBy || ""}
                                disabled
                              />
                            </div>
                            <div>
                              <label>Deliverable</label>
                              <input
                                type="text"
                                value={selectedTask?.deliverable || ""}
                                disabled
                                title={selectedTask?.deliverable || ""}
                              />
                            </div>
                            <div>
                              <label>Document Type</label>
                              <input
                                type="text"
                                value={selectedTask?.docType || ""}
                                disabled
                              />
                            </div>
                            <div>
                              <label>Document Number</label>
                              <input
                                type="text"
                                value={selectedTask?.docNumber || ""}
                                disabled
                              />
                            </div>
                            <div>
                              <label>Area</label>
                              <input
                                type="text"
                                value={selectedTask?.area || ""}
                                disabled
                              />
                            </div>
                            <div>
                              <label>Organization</label>
                              <input
                                type="text"
                                value={selectedTask?.org || ""}
                                disabled
                              />
                            </div>

                            <div>
                              <label>Revision Number</label>
                              <input
                                type="text"
                                value={
                                  selectedTask?.IsReworked === "Yes"
                                    ? (
                                        Number(
                                          selectedTask?.revisionnumber || 0,
                                        ) + 1
                                      ).toString()
                                    : selectedTask?.revisionnumber || "0"
                                }
                                disabled
                              />
                            </div>
                          </div>

                          {/* Upload & Comment Section */}
                          <div className="row">
                            <div className="col-sm-6">
                              <div className={styles.formActions}>
                                {/* Show upload field only for Pending tasks */}

                                {selectedTask?.status === "Pending" && (
                                  <>
                                    <label>Upload Documents*</label>
                                    <input
                                      ref={fileInputRef}
                                      type="file"
                                      className={
                                        fileError ? styles.inputError : ""
                                      }
                                      multiple
                                      accept=".pdf,.doc,.docx,.ppt,application/pdf,application/msword,application/vnd.openxmlformats-officedocument.wordprocessingml.document,application/vnd.ms-powerpoint"
                                      onChange={(e) => {
                                        setFileError(false);

                                        if (e.target.files) {
                                          const allowedExtensions =
                                            /\.(pdf|doc|docx|ppt)$/i;

                                          const validFiles: UploadedFile[] = [];

                                          for (const file of Array.from(
                                            e.target.files,
                                          )) {
                                            // Existing file type validation
                                            if (
                                              !allowedExtensions.test(file.name)
                                            ) {
                                              setPopup({
                                                isOpen: true,
                                                type: "validation",
                                                title: "Invalid File Type",
                                                message:
                                                  "Only PDF, Word (.doc, .docx), and PowerPoint (.ppt) files are allowed. Images are not supported.",
                                              });
                                              continue;
                                            }

                                            // New special character validation
                                            if (
                                              /[&%#<>:"/\\|?*]/.test(file.name)
                                            ) {
                                              setPopup({
                                                isOpen: true,
                                                type: "error",
                                                title: "Invalid File Name",
                                                message:
                                                  "Please remove special characters from file name.",
                                              });
                                              if (fileInputRef.current) {
                                                fileInputRef.current.value = "";
                                              }
                                              continue;
                                            }

                                            validFiles.push({
                                              name: file.name,
                                              file: file,
                                              url: undefined,
                                            });
                                          }

                                          if (validFiles.length > 0) {
                                            setSelectedFiles((prev) => [
                                              ...prev,
                                              ...validFiles,
                                            ]);
                                          }
                                        }
                                      }}
                                    />

                                    {/* Show selected files list */}
                                    {selectedFiles.length > 0 && (
                                      <div className={styles.selectedFilesList}>
                                        <label>Selected Files:</label>
                                        {selectedFiles.map((file, index) => (
                                          <div
                                            key={index}
                                            className={styles.fileItem}
                                          >
                                            <span>{file.name}</span>
                                            <button
                                              style={{ minWidth: "auto" }}
                                              type="button"
                                              onClick={() => {
                                                const newFiles = [
                                                  ...selectedFiles,
                                                ];
                                                newFiles.splice(index, 1);
                                                setSelectedFiles(newFiles);
                                                if (fileInputRef.current) {
                                                  fileInputRef.current.value =
                                                    "";
                                                }
                                              }}
                                              className={styles.removeFileBtn}
                                            >
                                              ✕
                                            </button>
                                          </div>
                                        ))}
                                      </div>
                                    )}
                                  </>
                                )}

                                {/* Existing Documents Section - for Pending tasks */}
                                {selectedTask?.status === "Pending" &&
                                  existingDocuments.length > 0 && (
                                    <>
                                      <label>
                                        Previously Uploaded Documents:
                                      </label>
                                      <div className={styles.existingFilesList}>
                                        {existingDocuments.map((doc, index) => (
                                          <p className="mb-0" key={index}>
                                            <a
                                              className="font-12"
                                              href="#"
                                              onClick={async (e) => {
                                                e.preventDefault();

                                                const secureUrl =
                                                  await getSecurePdfUrl(doc.id);

                                                window.open(
                                                  secureUrl,
                                                  "_blank",
                                                  "noopener,noreferrer",
                                                );
                                              }}
                                            >
                                              {doc.name}
                                            </a>
                                          </p>
                                        ))}
                                      </div>
                                    </>
                                  )}

                                {/* For Completed/In-Progress tasks - show all uploaded files */}
                                {(selectedTask?.status === "Approved" ||
                                  selectedTask?.status === "In-Progress") &&
                                  existingDocuments.length > 0 && (
                                    <>
                                      <label>All Uploaded Documents:</label>
                                      <div className={styles.existingFilesList}>
                                        {existingDocuments.map((doc, index) => (
                                          <p key={index}>
                                            <a
                                              href="#"
                                              onClick={async (e) => {
                                                e.preventDefault();

                                                const secureUrl =
                                                  await getSecurePdfUrl(doc.id);

                                                window.open(
                                                  secureUrl,
                                                  "_blank",
                                                  "noopener,noreferrer",
                                                );
                                              }}
                                            >
                                              {doc.name}
                                            </a>
                                          </p>
                                        ))}
                                      </div>
                                    </>
                                  )}
                              </div>
                            </div>
                            <div className="col-sm-6">
                              <label>Comment</label>
                              <textarea
                                className="form-control"
                                rows={2}
                                placeholder="Enter your comment"
                                value={comment}
                                onChange={(e) => setComment(e.target.value)}
                                disabled={
                                  selectedTask?.status === "Approved" ||
                                  selectedTask?.status === "In-Progress"
                                }
                              />
                            </div>
                          </div>

                          {/* Document Comments Accordion */}
                          {showDocumentComments && (
                            <div className={styles.accordionItem}>
                              <h2 className={styles.accordionHeader}>
                                <div className={styles.accordionButton}>
                                  Document Comments
                                </div>
                              </h2>

                              <div className={styles.accordionBody}>
                                <div className={styles.customCard}>
                                  <div
                                    className={styles.documentCommentsHeader}
                                  >
                                    <select
                                      className={styles.formSelect}
                                      value={selectedVersion}
                                      onChange={(e) =>
                                        onVersionChange(e.target.value)
                                      }
                                    >
                                      <option value="">
                                        -- Select Revision --
                                      </option>
                                      {versionList.map((version) => (
                                        <option key={version} value={version}>
                                          {version}
                                        </option>
                                      ))}
                                    </select>
                                    <button
                                      className={styles.btnOutlineSuccess}
                                      type="button"
                                      onClick={exportCommentsToExcel}
                                    >
                                      Export to Excel
                                    </button>
                                    <button
                                      className={styles.btnOutlineSuccess}
                                      type="button"
                                      onClick={refreshDocComment}
                                    >
                                      ↻
                                    </button>
                                  </div>

                                  <div className={styles.ribbonContent}>
                                    <table className={styles.commentsTable}>
                                      <thead>
                                        <tr>
                                          <th
                                            style={{
                                              minWidth: "80px",
                                              maxWidth: "80px",
                                            }}
                                          >
                                            Users
                                          </th>
                                          <th
                                            style={{
                                              minWidth: "100px",
                                              maxWidth: "100px",
                                            }}
                                          >
                                            Comment Date
                                          </th>
                                          <th
                                            style={{
                                              minWidth: "80px",
                                              maxWidth: "80px",
                                            }}
                                          >
                                            Page No.
                                          </th>
                                          <th
                                            style={{
                                              minWidth: "80px",
                                              maxWidth: "80px",
                                            }}
                                          >
                                            Revision
                                          </th>
                                          <th
                                            style={{
                                              minWidth: "200px",
                                              maxWidth: "200px",
                                            }}
                                          >
                                            Comments
                                          </th>
                                        </tr>
                                      </thead>
                                      <tbody
                                        style={{
                                          maxHeight: "250px",
                                          overflowY: "auto",
                                        }}
                                      >
                                        {documentComments.map((commentItem) => (
                                          <tr key={commentItem.id}>
                                            <td
                                              style={{
                                                padding: "10px",
                                                verticalAlign: "top",
                                                minWidth: "80px",
                                                maxWidth: "80px",
                                              }}
                                            >
                                              {commentItem.userName}
                                            </td>
                                            <td
                                              style={{
                                                padding: "10px",
                                                verticalAlign: "top",
                                                minWidth: "100px",
                                                maxWidth: "100px",
                                              }}
                                            >
                                              {commentItem.commentDate}
                                            </td>
                                            <td
                                              style={{
                                                padding: "10px",
                                                verticalAlign: "top",
                                                minWidth: "80px",
                                                maxWidth: "80px",
                                              }}
                                            >
                                              {commentItem.pageNumber}
                                            </td>
                                            <td
                                              style={{
                                                padding: "10px",
                                                verticalAlign: "top",
                                                minWidth: "80px",
                                                maxWidth: "80px",
                                              }}
                                            >
                                              {commentItem.revision}
                                            </td>
                                            <td
                                              style={{
                                                padding: "15px",
                                                verticalAlign: "top",
                                                minWidth: "200px",
                                                maxWidth: "200px",
                                              }}
                                            >
                                              {commentItem.comment}
                                            </td>
                                          </tr>
                                        ))}
                                      </tbody>
                                    </table>
                                  </div>
                                </div>
                              </div>
                            </div>
                          )}

                          {/* Audit History */}
                          <div className={styles.auditHistory}>
                            <h4>Audit History</h4>
                            {showNoAuditHistory && auditHistory.length === 0 ? (
                              <p>No audit history available</p>
                            ) : (
                              <table className="mtablemyt">
                                <thead>
                                  <tr>
                                    <th
                                      style={{
                                        minWidth: "70px",
                                        maxWidth: "70px",
                                      }}
                                    >
                                      SNo
                                    </th>
                                    <th>Approval Level</th>
                                    <th>Assigned To</th>
                                    <th>Assigned To Role</th>
                                    <th>Requestor Name</th>
                                    <th>Requested Date</th>
                                    <th>Action Taken By</th>
                                    <th>Action Taken On</th>
                                    <th>Remark</th>
                                    <th
                                      style={{
                                        minWidth: "110px",
                                        maxWidth: "110px",
                                      }}
                                    >
                                      Status
                                    </th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {auditHistory.map((historyItem) => (
                                    <tr key={historyItem.sno}>
                                      <td
                                        style={{
                                          minWidth: "70px",
                                          maxWidth: "70px",
                                        }}
                                      >
                                        {historyItem.sno}
                                      </td>
                                      <td title={historyItem.approvalLevel}>
                                        {historyItem.approvalLevel}
                                      </td>
                                      <td title={historyItem.assignedTo}>
                                        {historyItem.assignedTo}
                                      </td>
                                      <td title={historyItem.assignedToRole}>
                                        {historyItem.assignedToRole}
                                      </td>
                                      <td title={historyItem.requestorName}>
                                        {historyItem.requestorName}
                                      </td>
                                      <td title={historyItem.requestedDate}>
                                        {historyItem.requestedDate}
                                      </td>
                                      <td title={historyItem.assignedTo}>
                                        {historyItem.assignedTo}
                                      </td>
                                      <td title={historyItem.actionTakenOn}>
                                        {historyItem.actionTakenOn}
                                      </td>
                                      <td title={historyItem.remark}>
                                        {historyItem.remark}
                                      </td>
                                      <td
                                        style={{
                                          minWidth: "110px",
                                          maxWidth: "110px",
                                        }}
                                        title={historyItem.status}
                                      >
                                        <span
                                          className={`${styles.statusBadge} ${getStatusClass(historyItem.status)}`}
                                        >
                                          {historyItem.status}
                                        </span>
                                      </td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            )}
                          </div>

                          <div className={styles.formButtons}>
                            {selectedTask?.status === "Pending" && (
                              <>
                                <PrimaryButton
                                  text="Submit"
                                  iconProps={{ iconName: "CheckMark" }}
                                  onClick={() => {
                                    if (
                                      !selectedTask ||
                                      selectedFiles.length === 0
                                    ) {
                                      setFileError(true);

                                      setPopup({
                                        isOpen: true,
                                        type: "validation",
                                        title: "Validation",
                                        message:
                                          "Please fill mandatory fields.",
                                      });

                                      return;
                                    }

                                    setPopup({
                                      isOpen: true,
                                      type: "confirmation",
                                      title: "Confirm Submission",
                                      message:
                                        "Are you sure you want to submit this task?",
                                      onConfirm: () => {
                                        setPopup((prev) => ({
                                          ...prev,
                                          isOpen: false,
                                        }));

                                        handleSubmitClick();
                                      },
                                    });
                                  }}
                                  disabled={showSubmitLoader}
                                  className={styles.submitButton}
                                />

                                <DefaultButton
                                  text="Cancel"
                                  iconProps={{ iconName: "Cancel" }}
                                  onClick={handleBackClick}
                                  className={styles.cancelButton}
                                />
                              </>
                            )}

                            {selectedTask?.status === "Approved" && (
                              <DefaultButton
                                text="Back"
                                iconProps={{ iconName: "NavigateBack" }}
                                onClick={handleBackClick}
                                className={styles.backButton}
                              />
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>
                </>
              )}

              {/* Reusable popup for confirmation, success, error */}
              <CustomPopup
                isOpen={popup.isOpen}
                type={popup.type}
                title={popup.title}
                message={popup.message}
                onConfirm={popup.onConfirm}
                onCancel={() =>
                  setPopup((prev) => ({ ...prev, isOpen: false }))
                }
                onClose={() => setPopup((prev) => ({ ...prev, isOpen: false }))}
                onSuccessOk={() => {
                  setPopup((prev) => ({ ...prev, isOpen: false }));
                  if (popup.onConfirm) popup.onConfirm();
                }}
              />
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default MyTask;
