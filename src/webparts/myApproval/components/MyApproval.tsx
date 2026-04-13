
import { escape, set } from "@microsoft/sp-lodash-subset";

import styles from "./MyApproval.module.scss";
import React, { useRef, useState } from "react";

import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";

import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";

import "bootstrap/dist/css/bootstrap.min.css";

import "../../../CustomCss/mainCustom.scss";

import "../components/MyApproval.scss";

import "bootstrap/dist/js/bootstrap.bundle.min.js";

import "../../../CustomJSComponents/CustomTable/CustomTable.scss";
// import "./CustomTable.scss";

import "../../verticalSideBar/components/VerticalSidebar.scss";
//import "./CustomTable.scss";
import { IMyApprovalProps } from "./IMyApprovalProps";

import Provider from "../../../GlobalContext/provider";

import UserContext from "../../../GlobalContext/context";

import CustomBreadcrumb from "../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb";
import { getType } from "../../../APISearvice/CustomService";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";

import {
  faEdit,
  faPaperclip,
  faSort,
  faEye,
  faTrashAlt,
} from "@fortawesome/free-solid-svg-icons";

import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";

import * as XLSX from "xlsx";
let currentItemID = "";
import moment from "moment";
import { Web } from "@pnp/sp/webs";
import { AssignFrom } from "@pnp/core";

// import {

//   addItem,

//   GetCategory,

//   getChoiceFieldOption,

//   getDiscussionComments,

//   getDiscussionFilter,

//   getDiscussionFilterAll,

//   getDiscussionForum,

//   getDiscussionMe,

//   getDiscussionMeAll,

//   updateItem,

// } from "../../../APISearvice/DiscussionForumService";

import { encryptId } from "../../../APISearvice/CryptoService";

import { getNews } from "../../../APISearvice/NewsService";

import Swal from "sweetalert2";

import { getCategory, getEntity } from "../../../APISearvice/CustomService";

import ReactQuill from "react-quill";

import { uploadFileToLibrary } from "../../../APISearvice/MediaService";

import "react-quill/dist/quill.snow.css";

import { SPFI } from "@pnp/sp/presets/all";

// import { fetchUserInformationList } from "../../../APISearvice/GroupTeamService";

// import Multiselect from "multiselect-react-dropdown";

import { getSP } from "../loc/pnpjsConfig";

import { Eye, Edit } from "react-feather";

import {
  getDataByID,
  getMyApproval,
  getMyRequest,
  updateItemApproval,
} from "../../../APISearvice/ApprovalService";
import DMSMyApprovalAction from "./DMSApprovalAction";
import { getApprovalListsData } from "../../../APISearvice/BusinessAppsService";
// import DMSMyFolderApprovalAction from "./DMSFolderApprovalAction";
import DMSMyFolderApprovalAction from "./DMSFolderApprovalAction1";
import Select from 'react-select';
import loaderGif from "../assets/Loder.gif";//priyanshu
import CustomPopup from "../../../APISearvice/CustomPopup";//priyanshu


interface ApprovalHierarchyItem {
  id?: number;
  level: string;
  approverRole: string;
  approver: string;
  approvers: string[]; // New array for multiple approvers
  approvalCriteria: string;
  serialNumber?: number;
  assignedTo?: any[];
}
interface DocumentComment {
  id: number;
  userName: string;
  commentDate: string;
  pageNumber: string;
  revision: string;
  comment: string;
  // DocumentName: string;
}
interface UserOption {
  value: string;
  label: string;
  email: string;
}

// Id: details.Id,
//       Status: details.Status,
//       Author: { Title: details.Author?.Title || "" },
//       Created: details.Created,
//       InitiatedBy: creationItem?.AuthorId ?? creationItem?.AuthorId,
//       VendorId: DeliverablesItem?.AssignedToId ?? DeliverablesItem?.AssignedToId,
//       ProjectName: creationItem.ProjectName || "",
//       ProjectId: details.ProjectCreationListID?.ID,
//       DeliverableId: details.DeliverablesDetailsId?.ID,
//       PreparedBy: creationItem?.PreparedBy?.Title || details?.PreparedBy?.Title || "",
//       DocType: DeliverablesItem.DocumentType || "",

interface ProjectItem {
  Id: number;
  PreparedById: number;
  Status: string;
  Author: { Title: string };
  Created: string;
  InitiatedBy?: number;
  VendorId?: number;
  ProjectName: string;
  PreparedBy: string;
  DocType: string;
  ProjectId?: number;
  DeliverableId?: number;
  Deliverable: string;
  RevisionNumber: string;
  VendorName: string;
  documentNumber: string;
  docrevisionNumber: string;
  deliverablesDocumentIDs: any[];
  // ProjectType: string;
  // RequestID: string;
  // Title: string;
  // ApprovalTitle: string;
  // ProcessName: string;
  // Requester: { Title: string; EMail: string };
  // ClientName: string;
  // Area: string;
  // DocNumber: string;
  // DoYouNeedApproval: string;
  // CurrentApprovalLevel: string;
  ApprovalRole: string;
  // ApprovalSN: number;
  // ApprovalCriteria: string;
  // AssignedTo: string;
  // Org: string;
  // DocumentNumber: string;
  // ProjectDate: string;
  // Remarks: string;
  // RedirectionLink: string;
  // SNo: number;
}


let actingforuseremail: any
const MyApprovalContext = ({ props }: any) => {
  const sp: SPFI = getSP(props.context);
  const [activeComponent, setActiveComponent] = useState<string>("");
  const { useHide }: any = React.useContext(UserContext);
  const [showNestedDMSTable, setShowNestedDMSTable] = useState(false);
  const [announcementData, setAnnouncementData] = React.useState([]);

  const [myApprovalsData, setMyApprovalsData] = React.useState([]);
  const [myApprovalsDataAll, setMyApprovalsDataAll] = React.useState([]);
  const [myApprovalsDataAutomation, setMyApprovalsDataAutomation] =
    React.useState([]);
  const handleShowNestedDMSTable = () => {
    setShowNestedDMSTable(true); // Show nested table within DMS
  };
  const elementRef = React.useRef<HTMLDivElement>(null);

  const SiteUrl = props.siteUrl;
  const [folderActionOrFileAction, setFolderActionOrFileAction] = useState("");
  const [newsData, setNewsData] = React.useState([]);

  const [TypeData, setTypeData] = React.useState([]);

  const [BnnerImagepostArr, setBannerImagepostArr] = React.useState([]);

  const [DocumentpostArr, setDocumentpostArr] = React.useState([]);

  const [DocumentpostArr1, setDocumentpostArr1] = React.useState([]);

  const [ImagepostArr, setImagepostArr] = React.useState([]);

  const [ImagepostArr1, setImagepostArr1] = React.useState([]);

  const [GrouTypeData, setGroupTypeData] = React.useState([]);

  const [DocumentpostIdsArr, setDocumentpostIdsArr] = React.useState([]);

  const [selectedValue, setSelectedValue] = useState([]);

  const [EnityData, setEnityData] = React.useState([]);

  const [options, setOpions] = useState([]);

  const [approved, setApproved] = useState("yes");

  const [filters, setFilters] = React.useState({
    SNo: "",

    RequestID: "",

    ProcessName: "",

    RequestedBy: "",

    RequestedDate: "",

    Status: "",
    Title: "",
  });

  const [StatusTypeData, setStatusTypeData] = useState([
    { id: "Pending", name: "Pending" },
    { id: "Approved", name: "Approved" },
    { id: "Rejected", name: "Rejected" },
  ]);
  const [isOpen, setIsOpen] = React.useState(false);

  const [IsinvideHide, setIsinvideHide] = React.useState(false);
  // const [Mylistdata, setMylistdata] = useState([]);
  const [Mylistdata, setMylistdata] = useState<any[]>([]);
  const [ProjectWorkflowdata, setProjectWorkflowdata] = useState([]);
  const [selectedProjectTask, setSelectedProjectTask] = React.useState<any>(null);
  const [showProjectForm, setShowProjectForm] = React.useState(false);
  const [projectNeedsFurtherApproval, setProjectNeedsFurtherApproval] = React.useState<string>("Select");
  const [projectOutgoingstatus, setprojectOutgoingstatus] = React.useState<string>("Select");
  const [projectCodestatus, setprojectCodestatus] = React.useState<string>("Select");
  const [projectWantsToPublishInDossier, setProjectWantsToPublishInDossier] = React.useState<string>("Select");
  const [projectRemarks, setProjectRemarks] = React.useState<string>("");
  const [projectHierarchy, setProjectHierarchy] = React.useState<ApprovalHierarchyItem[]>([]);
  const [projectDocumentInfo, setProjectDocumentInfo] = React.useState<{
    documentUrl: string;
    fileName: string;
    fileLeafRef: string;
    fileRef: string;
    sharedLink: string;
    id: number;
  }[]>([]);
  // } | null>(null);
  const [documentComments, setDocumentComments] = React.useState<DocumentComment[]>([]);
  const [allDocumentComments, setAllDocumentComments] = React.useState<DocumentComment[]>([]);
  const [versionList, setVersionList] = React.useState<string[]>([]);
  const [selectedVersion, setSelectedVersion] = React.useState<string>("");
  const [showDocumentComments, setShowDocumentComments] = React.useState<boolean>(false);
  const [users, setUsers] = React.useState<UserOption[]>([]);
  // State for selected project items
  const [selectedProjectItems, setSelectedProjectItems] = useState<ProjectItem[]>([]);
  const [mdrStatus, setMdrStatus] = React.useState<string>("");
  const [currentUserId, setCurrentUserId] = React.useState<number | null>(null);

  const [showSubmitLoader, setShowSubmitLoader] = React.useState<boolean>(false);
  const [popupOpen, setPopupOpen] = React.useState(false);
  const [popupType, setPopupType] = React.useState<
    'confirmation' | 'validation' | 'success' | 'error'
  >('confirmation');
  const [popupTitle, setPopupTitle] = React.useState('');
  const [popupMessage, setPopupMessage] = React.useState('');
  const [pendingAction, setPendingAction] = React.useState<
    'Approved' | 'Rejected' | 'Rework' | null
  >(null);

  type ActionType =
    | "APPROVAL_ASSIGNED"
    | "APPROVED"
    | "REWORK_REQUESTED"
    | "SENT_FOR_PUBLISH"
    | "PUBLISHED"
    | "REWORK_BY_APPROVER"
    | "REWORK_BY_DOCUMENT_CONTROLLER";

  const actionConfig: Record<ActionType, { subject: string; message: string }> = {
    APPROVAL_ASSIGNED: {
      subject: "Approval Action Assigned",
      message: "An approval action has been assigned to you for the following deliverable(s)."
    },
    APPROVED: {
      subject: "Deliverable Approved",
      message: "The deliverable(s) mentioned below have been approved."
    },
    REWORK_REQUESTED: {
      subject: "Rework Requested",
      message: "Rework has been requested for the following deliverable(s)."
    },
    SENT_FOR_PUBLISH: {
      subject: "Deliverable Pending Publication",
      message: "The deliverable(s) mentioned below are pending publication."
    },

    PUBLISHED: {
      subject: "Deliverable Published",
      message: "The deliverable(s) mentioned below have been published."
    },
    REWORK_BY_APPROVER: {
      subject: "Rework Requested by Approver",
      message: "The Approver has requested rework for the following deliverable(s)."
    },
    REWORK_BY_DOCUMENT_CONTROLLER: {
      subject: "Rework Requested by Document Controller",
      message: "The Document Controller has requested rework for the following deliverable(s)."
    }
  };








  const handleReturnToMain = (Name: any) => {
    setActiveComponent(Name); // Reset to show the main component
    console.log(activeComponent, "activeComponent updated");
  };
  const getUserTitleByEmail = async (userEmail: any) => {
    try {
      const user = await sp.web.siteUsers.getByEmail(userEmail)();
      return user.Title;
    } catch (error) {
      console.error("Error fetching user title:", error);
      return null;
    }
  };
  const [loading, setLoading] = useState(true);
  const [actingForUser, setSetActingForUser] = useState([]);

  const myActingfordata = async () => {
    try {
      const currentUserEmail = currentUserEmailRef.current;
      console.log("currentUserEmail myActingfordata", currentUserEmail);
      const today = new Date().toISOString();
      const delegateListItems = await sp.web.lists.getByTitle('ARGDelegateList').items.select(
        "DelegateName/EMail",
        "ActingFor/EMail",
        "ActingFor/Title",
        "DelegateName/Title",
        "StartDate",
        "EndDate",
        "Status"
      )
        .expand("DelegateName", "ActingFor")
        .filter(`ActingFor/EMail eq '${currentUserEmail}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`)();

      console.log("delegateListItems myActingfordata", delegateListItems);

      // Extract unique ActingFor.Title and EMail values
      const uniqueTitlesAndEmails = [
        ...new Map(
          delegateListItems.map((item) => [item.DelegateName?.Title, { title: item.DelegateName?.Title, email: item.DelegateName?.EMail }])
        ).values(),
      ];

      // Set state with unique titles and emails
      setSetActingForUser(uniqueTitlesAndEmails.map((item, index) => ({ id: index.toString(), name: item.title, email: item.email })));
      console.log("setSetActingForUser", actingForUser);
      console.log("uniqueTitlesAndEmails", uniqueTitlesAndEmails);

    } catch (error) {
      console.error("Error fetching data:", error);
    } finally {
      //  setLoading(false);
    }
  };



  const getProjectWorkflowApprovals = async (status: string, actingfor?: any) => {
    try {
      const currentUserEmail = actingfor || currentUserEmailRef.current;

      // Get current user ID for filtering
      const currentUser = await sp.web.currentUser();
      const currentUserId = currentUser.Id;

      // Fetch main approval items with the same fields and expands as your example
      const items = await sp.web.lists.getByTitle("ProjectApprovals").items
        .select(
          "*",
          "DeliverablesDetailsId/Deliverables",
          "DeliverablesDetailsId/DocNumber",
          "DeliverablesDetailsId/Organization",
          "DeliverablesDetailsId/Area",
          "DeliverablesDetailsId/ID",
          "ProjectCreationListID/ProjectName",
          "ProjectCreationListID/ID",
          "AssignedTo/ID",
          "AssignedTo/Title",
          "AssignedTo/EMail",
          "Author/Title",
          "Author/EMail"
        )
        .expand("DeliverablesDetailsId", "ProjectCreationListID", "AssignedTo", "Author")
        .filter(`AssignedTo/ID eq ${currentUserId} and Status eq '${status}' and ApproverRole ne 'Vendor'`)
        .orderBy("Created", false)();

      console.log(items, "ProjectApprovals List with current user filter");

      // Transform items with additional data from ProjectCreationList (similar to your example)
      const approvals = await Promise.all(
        items.map(async (item: any, index: number) => {
          try {
            let creationItem = null;

            // Fetch additional project creation data if ProjectCreationListID exists
            if (item.ProjectCreationListID?.ID) {
              creationItem = await sp.web.lists.getByTitle("ProjectCreationList").items.getById(item.ProjectCreationListID.ID)
                .select(
                  "*",
                  "PreparedBy/Title",
                  "ProjectType/ProjectType",
                  "ProjectType/Id",
                  "ClientName"
                )
                .expand("ProjectType", "PreparedBy")();
            }
            let DeliverablesItem = null;

            // Fetch additional project creation data if DeliverablesDetails exists
            if (item.DeliverablesDetailsId?.ID) {
              DeliverablesItem = await sp.web.lists.getByTitle("DeliverablesDetails").items.getById(item.DeliverablesDetailsId.ID)
                .select(
                  "*",
                  "AssignedTo/ID",
                  "AssignedTo/Title",
                )
                .expand("AssignedTo")();
            }

            // Format the data to match your table structure
            return {
              Id: item.Id,
              RequestID: item.DocNumber || `PROJ-${item.Id}`,
              Title: item.ProjectCreationListID?.ProjectName || "Project Approval",
              ApprovalTitle: item.ProjectCreationListID?.ProjectName || "Project Approval",
              ProcessName: "Project Workflow",
              Status: item.Status,
              Requester: {
                Title: item.Author?.Title || "",
                EMail: item.Author?.EMail || ""
              },
              Author: {
                Title: item.Author?.Title || "",
              },
              Created: item.Created,
              InitiatedBy: creationItem?.AuthorId,
              VendorId: DeliverablesItem?.AssignedToId,

              // Project-specific fields from your example
              ProjectName: item.ProjectCreationListID?.ProjectName || "",
              ProjectType: creationItem?.ProjectType?.ProjectType || "",
              ClientName: creationItem?.ClientName || "",
              PreparedBy: creationItem?.PreparedBy?.Title || "",
              Deliverable: item.DeliverablesDetailsId?.Deliverables || "",
              Area: item.DeliverablesDetailsId?.Area || "",
              DocType: item.DocumentType || "",
              DocNumber: item.DeliverablesDetailsId?.DocNumber || "",
              DoYouNeedApproval: item.Doyouneedapproval || "",
              CurrentApprovalLevel: item.Level || "",
              ApprovalRole: item.ApproverRole || "",
              ApprovalSN: item.SerialNumber || 0,
              ApprovalCriteria: item.ApprovalCriteria || "",
              AssignedTo: item.AssignedTo?.Title || "",
              Org: item.DeliverablesDetailsId?.Organization || "",
              RevisionNumber: item.RevisionNumber || "0",
              DocumentNumber: item.DocNumber || "",
              ProjectDate: item.Created ? new Date(item.Created).toLocaleDateString('en-GB') : "",
              ProjectId: item.ProjectCreationListID?.ID,
              DeliverableId: item.DeliverablesDetailsId?.ID,
              Remarks: item.Remarks || "",

              // For redirection or additional actions
              RedirectionLink: item.RedirectionLink || "", // Add if you have this field

              // Additional fields that might be useful for filtering
              SNo: index + 1
            };
          } catch (error) {
            console.error(`Error fetching creation item for project ${item.ProjectCreationListID?.ID}:`, error);

            // Return fallback data if creationItem fetch fails
            return {
              Id: item.Id,
              RequestID: item.DocNumber || `PROJ-${item.Id}`,
              Title: item.ProjectCreationListID?.ProjectName || "Project Approval",
              ApprovalTitle: item.ProjectCreationListID?.ProjectName || "Project Approval",
              ProcessName: "Project Workflow",
              Status: item.Status,
              Requester: {
                Title: item.Author?.Title || "",
                EMail: item.Author?.EMail || ""
              },
              Author: {
                Title: item.Author?.Title || "",
              },
              Created: item.Created,
              ProjectName: item.ProjectCreationListID?.ProjectName || "",
              ProjectType: "",
              ClientName: "",
              PreparedBy: "",
              Deliverable: item.DeliverablesDetailsId?.Deliverables || "",
              Area: item.DeliverablesDetailsId?.Area || "",
              DocType: item.DocumentType || "",
              DocNumber: item.DeliverablesDetailsId?.DocNumber || "",
              DoYouNeedApproval: item.Doyouneedapproval || "",
              CurrentApprovalLevel: item.Level || "",
              ApprovalRole: item.ApproverRole || "",
              ApprovalSN: item.SerialNumber || 0,
              ApprovalCriteria: item.ApprovalCriteria || "",
              AssignedTo: item.AssignedTo?.Title || "",
              Org: item.DeliverablesDetailsId?.Organization || "",
              RevisionNumber: item.RevisionNumber || "0",
              DocumentNumber: item.DocNumber || "",
              ProjectDate: item.Created ? new Date(item.Created).toLocaleDateString('en-GB') : "",
              ProjectId: item.ProjectCreationListID?.ID,
              DeliverableId: item.DeliverablesDetailsId?.ID,
              Remarks: item.Remarks || "",
              RedirectionLink: item.RedirectionLink || "",
              SNo: index + 1
            };
          }
        })
      );

      console.log("Transformed Project Approvals:", approvals);
      return approvals;

    } catch (error) {
      console.error("Error fetching Project Workflow items:", error);
      return [];
    }
  };

  // Add these functions
  const addNewProjectApprovalRow = (item?: any) => {
    const newRow: ApprovalHierarchyItem = {
      level: `Level ${projectHierarchy.length + 1}`,
      approverRole: item?.ApproverRole || '',
      approver: item?.AssignedTo?.Title || '',
      approvers: item?.AssignedTo ? [item.AssignedTo.Title] : [],
      approvalCriteria: item?.ApprovalCriteria || 'Anyone',
      assignedTo: item?.AssignedTo ? [item.AssignedTo] : []
    };
    setProjectHierarchy(prev => [...prev, newRow]);
  };

  const deleteProjectApprovalRow = (index: number) => {
    setProjectHierarchy(prev => {
      const updatedHierarchy = prev.filter((_, i) => i !== index);
      // Renumber the levels sequentially
      return updatedHierarchy.map((row, i) => ({
        ...row,
        level: `Level ${i + 1}`
      }));
    });
  };

  const updateProjectApprovalRow = (index: number, field: keyof ApprovalHierarchyItem, value: string) => {
    setProjectHierarchy(prev => prev.map((row, i) =>
      i === index ? { ...row, [field]: value } : row
    ));
  };

  const handleProjectApproverChange = (index: number, selectedOptions: any) => {
    setProjectHierarchy(prev => prev.map((row, i) => {
      if (i === index) {
        const selectedUsers = selectedOptions || [];
        const approverNames = selectedUsers.map((user: any) => user.label).join(', ');
        const assignedToArray = selectedUsers.map((user: any) => ({
          ID: parseInt(user.value),
          Title: user.label,
          EMail: user.email
        }));

        return {
          ...row,
          approver: approverNames,
          assignedTo: assignedToArray
        };
      }
      return row;
    }));
  };

  // const getApprovalmasterTasklist = async (value: any, actingfor?: any) => {
  //   // alert(`Status value is ${value} is acting for ${actingfor} in DMS`)

  //   try {
  //     // Retrieve current user email
  //     const currentUserEmail = currentUserEmailRef.current;

  //     // Fetch the ARGDelegateList items where the current user is in the ActingFor column
  //     const today = new Date().toISOString(); // Get today's date in YYYY-MM-DD format
  //     // console.log("today", today);


  //     let arr = [];
  //     let approvalData: any[] = [];
  //     if (!actingfor) {
  //       const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
  //         "Log", "CurrentUser", "Remark"
  //         , "LogHistory"
  //         , "FileUID/FileUID"
  //         , "FileUID/SiteName"
  //         , "FileUID/DocumentLibraryName"
  //         , "FileUID/FileName"
  //         , "FileUID/RequestNo"
  //         , "FileUID/Processname"
  //         //  ,"FileUID/FilePreviewUrl" 
  //         , "FileUID/Status"
  //         , "FileUID/FolderPath"
  //         , "FileUID/RequestedBy"
  //         , "FileUID/Created"
  //         , "FileUID/ApproveAction"
  //         , "MasterApproval/ApprovalType"
  //         , "MasterApproval/Level"
  //         , "MasterApproval/DocumentLibraryName"

  //       )
  //         .expand("FileUID", "MasterApproval")
  //         .filter(`CurrentUser eq '${currentUserEmailRef.current}' and FileUID/Status eq '${value}'`).orderBy("Created", false)();
  //       console.log(items, "DMSFileApprovalTaskList");
  //       items.map((item) => {
  //         if (item.CurrentUser !== currentUserEmailRef.current) {
  //           arr.push(item)
  //           // alert(`Delegate user ${item.CurrentUser} is acting for ${item.FileUID.FileName}`)
  //         }

  //       });
  //       const updatedItems = await Promise.all(items.map(async (item) => {
  //         const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
  //         return { ...item, RequestedByTitle: requestedbyuserTitle };
  //       }));
  //       approvalData = updatedItems
  //       setMylistdata(updatedItems);
  //     }
  //     if (actingfor !== "" && actingfor !== undefined) {
  //       const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
  //         "Log", "CurrentUser", "Remark"
  //         , "LogHistory"
  //         , "FileUID/FileUID"
  //         , "FileUID/SiteName"
  //         , "FileUID/DocumentLibraryName"
  //         , "FileUID/FileName"
  //         , "FileUID/RequestNo"
  //         , "FileUID/Processname"
  //         //  ,"FileUID/FilePreviewUrl" 
  //         , "FileUID/Status"
  //         , "FileUID/FolderPath"
  //         , "FileUID/RequestedBy"
  //         , "FileUID/Created"
  //         , "FileUID/ApproveAction"
  //         , "MasterApproval/ApprovalType"
  //         , "MasterApproval/Level"
  //         , "MasterApproval/DocumentLibraryName"

  //       )
  //         .expand("FileUID", "MasterApproval")
  //         .filter(`CurrentUser eq '${actingfor}' and FileUID/Status eq '${value}'`).orderBy("Created", false)();
  //       console.log(items, "DMSFileApprovalTaskList");
  //       items.map((item) => {
  //         if (item.CurrentUser !== currentUserEmailRef.current) {
  //           arr.push(item)
  //           // alert(`Delegate user ${item.CurrentUser} is acting for ${item.FileUID.FileName}`)
  //         }

  //       });
  //       const updatedItems = await Promise.all(items.map(async (item) => {
  //         const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
  //         return { ...item, RequestedByTitle: requestedbyuserTitle };
  //       }));
  //       approvalData = updatedItems;
  //       setMylistdata(updatedItems);
  //     }

  //     // const updatedItems2 = await Promise.all(items2.map(async (item) => {
  //     //   const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
  //     //   return { ...item, RequestedByTitle: requestedbyuserTitle };
  //     // }));
  //     if (!actingfor) {
  //       const Item2: any = await sp.web.lists.getByTitle('DMSFolderDeligationApprovalTask').items.select(
  //         "*",
  //         "Folderdetail"
  //         , "Folderdetail/SiteTitle"
  //         , "Folderdetail/DocumentLibraryName"
  //         , "Folderdetail/CurrentUser"
  //         , "Folderdetail/FolderPath"
  //         , "Folderdetail/FolderName"
  //         , "Folderdetail/ParentFolderId"
  //         , "Folderdetail/Department"
  //         , "Folderdetail/Devision"
  //         , "Folderdetail/RequestNo"
  //         , "FolderMeta"
  //         , "FolderMeta/SiteName"
  //         , "FolderMeta/DocumentLibraryName"
  //         , "FolderMeta/ColumnName",
  //         "Folderdetail/Processname",
  //         "Folderdetail/Status",
  //         "Approver"
  //       ).expand("Folderdetail", "FolderMeta")
  //         .filter(`Approver eq '${currentUserEmailRef.current}' and Folderdetail/Status eq '${value}'`)();
  //       console.log("Item2 get from dmsfolderdeligationapprovaltasklist", Item2)
  //       const normalizeItem2 = async (item: any) => ({
  //         Log: item?.Log || '', // Replace with appropriate mappings
  //         CurrentUser: item?.Folderdetail?.CurrentUser || '',
  //         Remark: item?.Remark || '',
  //         LogHistory: item?.LogHistory || '',
  //         // ProcessName:  item?.Folderdetail?.Processname,
  //         RequestedByTitle: await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
  //         FileUID: {
  //           FileUID: item?.FolderMeta?.FileUID || item?.Folderdetail?.RequestNo,
  //           SiteName: item?.FolderMeta?.SiteName || '',
  //           DocumentLibraryName: item?.FolderMeta?.DocumentLibraryName || '',
  //           // FileName: item?.FolderMeta?.FolderName || '',
  //           FileName: item?.Folderdetail?.FolderName === null ? item?.Folderdetail?.DocumentLibraryName : item?.Folderdetail?.FolderName,
  //           RequestNo: item?.Folderdetail?.RequestNo || '',
  //           Status: item?.Folderdetail?.Status || '',
  //           FolderPath: item?.Folderdetail?.FolderPath || '',
  //           // RequestedBy: item?.RequestedBy || item?.Folderdetail?.CurrentUser || '',
  //           // RequestedByTitle:await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
  //           Created: item?.Created || '',
  //           ApproveAction: item?.ApproveAction || '',
  //           Processname: item?.Folderdetail?.Processname
  //         },
  //         MasterApproval: {
  //           ApprovalType: item?.ApprovalType || '',
  //           Level: item?.Level || '',
  //           DocumentLibraryName: item?.DocumentLibraryName || ''
  //         }
  //       });
  //       // const normalizeItem3 = Item2.map(normalizeItem2);
  //       const normalizeItem3 = await Promise.all(Item2.map(normalizeItem2));
  //       console.log("normalizeItem3", normalizeItem3);
  //       console.log(approvalData, "approvalData 666");
  //       const CombinedItems = [...approvalData, ...normalizeItem3];
  //       console.log(CombinedItems, "CombinedItems")
  //       setMylistdata(CombinedItems);
  //       // setMylistdata(updatedItems);


  //       // return arr = CombinedItems
  //     }
  //     if (actingfor !== "" && actingfor !== undefined) {
  //       const Item2: any = await sp.web.lists.getByTitle('DMSFolderDeligationApprovalTask').items.select(
  //         "*",
  //         "Folderdetail"
  //         , "Folderdetail/SiteTitle"
  //         , "Folderdetail/DocumentLibraryName"
  //         , "Folderdetail/CurrentUser"
  //         , "Folderdetail/FolderPath"
  //         , "Folderdetail/FolderName"
  //         , "Folderdetail/ParentFolderId"
  //         , "Folderdetail/Department"
  //         , "Folderdetail/Devision"
  //         , "Folderdetail/RequestNo"
  //         , "FolderMeta"
  //         , "FolderMeta/SiteName"
  //         , "FolderMeta/DocumentLibraryName"
  //         , "FolderMeta/ColumnName",
  //         "Folderdetail/Processname",
  //         "Folderdetail/Status",
  //         "Approver"
  //       ).expand("Folderdetail", "FolderMeta")
  //         .filter(`Approver eq '${actingfor}' and Folderdetail/Status eq '${value}'`)();
  //       console.log("Item2 get from dmsfolderdeligationapprovaltasklist", Item2)
  //       const normalizeItem2 = async (item: any) => ({
  //         Log: item?.Log || '', // Replace with appropriate mappings
  //         CurrentUser: item?.Folderdetail?.CurrentUser || '',
  //         Remark: item?.Remark || '',
  //         LogHistory: item?.LogHistory || '',
  //         // ProcessName:  item?.Folderdetail?.Processname,
  //         RequestedByTitle: await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
  //         FileUID: {
  //           FileUID: item?.FolderMeta?.FileUID || item?.Folderdetail?.RequestNo,
  //           SiteName: item?.FolderMeta?.SiteName || '',
  //           DocumentLibraryName: item?.FolderMeta?.DocumentLibraryName || '',
  //           // FileName: item?.FolderMeta?.FolderName || '',
  //           FileName: item?.Folderdetail?.FolderName === null ? item?.Folderdetail?.DocumentLibraryName : item?.Folderdetail?.FolderName,
  //           RequestNo: item?.Folderdetail?.RequestNo || '',
  //           Status: item?.Folderdetail?.Status || '',
  //           FolderPath: item?.Folderdetail?.FolderPath || '',
  //           // RequestedBy: item?.RequestedBy || item?.Folderdetail?.CurrentUser || '',
  //           // RequestedByTitle:await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
  //           Created: item?.Created || '',
  //           ApproveAction: item?.ApproveAction || '',
  //           Processname: item?.Folderdetail?.Processname
  //         },
  //         MasterApproval: {
  //           ApprovalType: item?.ApprovalType || '',
  //           Level: item?.Level || '',
  //           DocumentLibraryName: item?.DocumentLibraryName || ''
  //         }
  //       });
  //       // const normalizeItem3 = Item2.map(normalizeItem2);
  //       const normalizeItem3 = await Promise.all(Item2.map(normalizeItem2));
  //       console.log("normalizeItem3", normalizeItem3);
  //       console.log(approvalData, "approvalData 666");
  //       const CombinedItems = [...approvalData, ...normalizeItem3];
  //       console.log(CombinedItems, "CombinedItems")
  //       setMylistdata(CombinedItems);
  //       // setMylistdata(updatedItems);


  //       // return arr = CombinedItems
  //     }
  //   } catch (error) {
  //     console.error("Error fetching list items:", error);
  //   }
  // };
  
  

const getApprovalmasterTasklist = async (value: any, actingfor?: any) => {
  try {
    const currentUserEmail = currentUserEmailRef.current;
    let allCombinedItems: any[] = [];

    // 1. Fetch all site URLs from the MasterSiteCollection list
    const masterSites = await sp.web.lists
      .getByTitle("MasterSiteCollection")
      .items.select("SiteURL")();

    // 2. Map through each site and fetch data
    const fetchPromises = masterSites.map(async (site: any) => {
      const siteUrl = site.SiteURL;
      
      try {
        // Create a dedicated web object for the remote site
        const remoteWeb = Web(siteUrl).using(AssignFrom(sp.web));

        // Determine who we are filtering for
        const targetUser = actingfor || currentUserEmail;
        // srs 18/3/26
        const isNotPending = value !== "Pending";
        const filterString = isNotPending 
        ? `CurrentUser eq '${targetUser}' and Log eq '${value}'` 
        : `CurrentUser eq '${targetUser}' and (Log eq null)`;

        // Fetch File Approvals from the remote site
        const fileItems = await remoteWeb.lists
          .getByTitle("DMSFileApprovalTaskList")
          .items.select(
            "Log", "CurrentUser", "Remark", "LogHistory",
            "FileUID/FileUID", "FileUID/SiteName", "FileUID/DocumentLibraryName",
            "FileUID/FileName", "FileUID/RequestNo", "FileUID/Processname",
            "FileUID/Status", "FileUID/FolderPath", "FileUID/RequestedBy",
            "FileUID/Created", "FileUID/ApproveAction",
            "MasterApproval/ApprovalType", "MasterApproval/Level", "MasterApproval/DocumentLibraryName"
          )
          .expand("FileUID", "MasterApproval")
          // .filter(`CurrentUser eq '${targetUser}' and FileUID/Status eq '${value}'`)
          // srs 16/3/26
          // .filter(`CurrentUser eq '${targetUser}' and (Log eq null)`)
          // srs 18/3/26
          .filter(filterString)
          .orderBy("Created", false)();

        // Fetch Folder Delegations from the remote site
        const folderItems = await remoteWeb.lists
          .getByTitle("DMSFolderDeligationApprovalTask")
          .items.select(
            "*", "Folderdetail/SiteTitle", "Folderdetail/DocumentLibraryName",
            "Folderdetail/CurrentUser", "Folderdetail/FolderPath", "Folderdetail/FolderName",
            "Folderdetail/RequestNo", "Folderdetail/Processname", "Folderdetail/Status",
            "FolderMeta/ID", "FolderMeta/SiteName", "FolderMeta/DocumentLibraryName",
            "Approver"
          )
          .expand("Folderdetail", "FolderMeta")
          .filter(`Approver eq '${targetUser}' and Folderdetail/Status eq '${value}'`)();

        // Normalize Folder Items (Same as your existing logic)
        const normalizedFolders = await Promise.all(folderItems.map(async (item) => ({
          Log: item?.Log || '',
          CurrentUser: item?.Folderdetail?.CurrentUser || '',
          Remark: item?.Remark || '',
          RequestedByTitle: await getUserTitleByEmail(item?.Folderdetail?.CurrentUser),
          FileUID: {
            FileUID: item?.FolderMeta?.FileUID || item?.Folderdetail?.RequestNo,
            SiteName: item?.FolderMeta?.SiteName || '',
            FileName: item?.Folderdetail?.FolderName || item?.Folderdetail?.DocumentLibraryName,
            // Status: item?.Folderdetail?.Status || '',
            // srs 18/3/26
            Status: item?.Folderdetail?.Log || item?.Folderdetail?.Status || '',
            Processname: item?.Folderdetail?.Processname
            // ... add other fields as per your normalizeItem2
          },
          MasterApproval: {
            ApprovalType: item?.ApprovalType || '',
            Level: item?.Level || ''
          }
        })));

        // Enrich File Items with User Titles
        const enrichedFiles = await Promise.all(fileItems.map(async (item) => {
          const requestedbyuserTitle = await getUserTitleByEmail(item?.FileUID?.RequestedBy);
          // return { ...item, RequestedByTitle: requestedbyuserTitle };
          // srs 18/3/26
          return { 
            ...item, 
            RequestedByTitle: requestedbyuserTitle,
            FileUID: {
               ...item.FileUID,
               // BIND LOG TO STATUS COLUMN
               Status: item.Log || item?.FileUID?.Status || "" 
            }};
        }));

        return [...enrichedFiles, ...normalizedFolders];

      } catch (siteErr) {
        console.error(`Error fetching data from site: ${siteUrl}`, siteErr);
        return []; // Return empty if one site fails to keep the loop going
      }
    });

    // 3. Wait for all sites to return data and flatten the array
    const results = await Promise.all(fetchPromises);
    allCombinedItems = results.flat();

    console.log("Final Aggregated Data:", allCombinedItems);
    setMylistdata(allCombinedItems);
    // srs 10/4/26
    if (activeTab === "DMS") {
      setMyApprovalsData(allCombinedItems);
      setLoading(false); // Stop the bird loader
    }
    return allCombinedItems;
  } catch (error) {
    console.error("Master Site Collection fetch failed:", error);
  }
};


  console.log(Mylistdata, "Mylistdata");
  const currentUserEmailRef = useRef("");
  const getCurrrentuser = async () => {
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    getApprovalmasterTasklist('Pending', '');
    myActingfordata()
  };
  const fetchUsers = async () => {
    try {
      const siteUsers = await sp.web.siteUsers();
      const userOptions: UserOption[] = siteUsers
        .filter((user: any) => user.Email && user.Title) // Filter out users without email or title
        .map((user: any) => ({
          value: user.Id.toString(),
          label: user.Title,
          email: user.Email
        }));
      setUsers(userOptions);
    } catch (error) {
      console.error("Error fetching users:", error);
      // Fallback to some default users if fetch fails
      setUsers([
        { value: '1', label: 'Current User', email: props.context.pageContext.user.email },
        { value: '2', label: 'Admin User', email: 'admin@contoso.com' }
      ]);
    }
  };
  React.useEffect(() => {
    getCurrrentuser();
    fetchUsers();
    setCurrentUserId(props.context.pageContext.legacyPageContext.userId);
  }, []);





  const truncateText = (text: string, maxLength?: any) => {
    if (text) {
      return text.length > maxLength
        ? text.substring(0, maxLength) + "..."
        : text;
    }
  };

  const getTaskItemsbyID = async (e: any, itemid: any, ProcessName: string) => {
    // currentItemID = itemid
    currentItemID = itemid;
    setActiveComponent("Approval Action");
    setFolderActionOrFileAction(ProcessName);
    console.log("itemid", itemid);
    // const items = await sp.web.lists
    //   .getByTitle("DMSFileApprovalTaskList")
    //   .items.select("CurrentUser", "FileUID/FileUID", "Log")
    //   .expand("FileUID")
    //   .filter(`FileUID/RequestNo eq '${itemid}'`)();
    // console.log(items, "items");
  };
  const getTaskItemsbyID2 = async (e: any, itemid: any) => {
    // alert("Folder")
    // currentItemID = itemid
    currentItemID = itemid
    setActiveComponent('DMS Folder Approval')
    console.log("itemid", itemid)
    // const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select("CurrentUser" , "FileUID/FileUID" , "Log").expand("FileUID").filter(`FileUID/RequestNo eq '${itemid}'`)();
    //    console.log(items , "items")
  }
  const toggleDropdown = () => {
    setIsOpen(!isOpen);
  };

  const [sortConfig, setSortConfig] = React.useState({
    key: "",

    direction: "ascending",
  });

  const [formData, setFormData] = React.useState({
    Remark: "",
  });
  const handleCancel = () => {
    window.location.href = `${siteUrl}/SitePages/MyApprovals.aspx`;
  };
  //#region OnchangeData

  const onChange = (name: string, value: string) => {
    debugger;

    setFormData((prevData) => ({
      ...prevData,

      [name]: value,
    }));
  };

  //#endregion

  // const [formData, setFormData] = React.useState({

  //   topic: "",

  //   category: "",

  //   entity: "",

  //   Type: "",

  //   GroupType: "",

  //   description: "",

  //   overview: "",

  //   FeaturedAnnouncement: false,

  // });

  const [approveData, setApproveData] = useState([]);

  const [isActivedata, setisActivedata] = useState(false);

  const [DiscussionData, setDiscussion] = useState([]);

  const [CategoryData, setCategoryData] = React.useState([]);
  const [Statusvalue, SetStatusvalue] = useState("Pending");
  const [showModal, setShowModal] = React.useState(false);
  const [StatusChange, setStatusChange] = React.useState(false);
  const [showDocTable, setShowDocTable] = React.useState(false);

  const [showImgModal, setShowImgTable] = React.useState(false);

  const [showBannerModal, setShowBannerTable] = React.useState(false);

  const [currentUser, setCurrentUser] = React.useState(null);

  const [editForm, setEditForm] = React.useState(false);

  const [richTextValues, setRichTextValues] = React.useState<{
    [key: string]: string;
  }>({});

  const [documentControllerId, setDocumentControllerId] = React.useState<number | null>(null);
  const [dccId, setDccId] = React.useState<number | null>(null);


  //const [activeTab, setActiveTab] = useState("home1");
  // const [activeTab, setActiveTab] = useState("Automation");
  // srs 10/4/26
  const [activeTab, setActiveTab] = useState("DMS");
  const handleTabClick = async (tab: React.SetStateAction<string>) => {
    // Aman 25/3/26
    if (activeTab === tab) return;
     // Aman 25/3/26 ended
    setActiveTab(tab);
    debugger
    // Aman 25/3/26
    // setMyApprovalsData([]);
    // Aman 25/3/26 ended
    setLoading(true);
    // console.log(
    //   "tab",
    //   tab,
    //   myApprovalsDataAutomation,
    //   myApprovalsDataAll,
    //   myApprovalsData
    // );

    // if (tab == "Intranet") {
    //   setMyApprovalsData(myApprovalsDataAll);
    // } else if (tab == "DMS") {
    //   setMyApprovalsData(Mylistdata);
    // } else if (tab == "Automation") {
    //   //ApiCall("Pending");
    //   setMyApprovalsData(myApprovalsDataAutomation);
    //   //setMyApprovalsDataAutomation(myApprovalsDataAutomation);
    // }
    let MyApprovaldata: any = [];
    let Automationdata: any = [];
    // let MyDMSAPPROVALDATA:any = await MyDMSAPPROVALDATASTATUS(sp, value)
    let MyDMSAPPROVALDATA: any = [];
    let ProjectWorkflowData: any = []; // NEW
    if (actingforuseremail === undefined || actingforuseremail === null || actingforuseremail === "") {
      MyApprovaldata = await getMyApproval(sp, Statusvalue);
      Automationdata = await getApprovalListsData(sp, Statusvalue);
      // let MyDMSAPPROVALDATA:any = await MyDMSAPPROVALDATASTATUS(sp, value)
      MyDMSAPPROVALDATA = await getApprovalmasterTasklist(Statusvalue)
      ProjectWorkflowData = await getProjectWorkflowApprovals(Statusvalue); // NEW
    } else {
      MyApprovaldata = await getMyApproval(sp, Statusvalue, actingforuseremail);
      Automationdata = await getApprovalListsData(sp, Statusvalue, actingforuseremail);
      // let MyDMSAPPROVALDATA:any = await MyDMSAPPROVALDATASTATUS(sp, value)
      MyDMSAPPROVALDATA = await getApprovalmasterTasklist(Statusvalue, actingforuseremail)
      ProjectWorkflowData = await getProjectWorkflowApprovals(Statusvalue, actingforuseremail); // NEW
    }
    console.log("MyDMSAPPROVALDATA", MyDMSAPPROVALDATA)
    setMyApprovalsDataAll(MyApprovaldata);
    setMyApprovalsDataAutomation(Automationdata);
    setProjectWorkflowdata(ProjectWorkflowData); // NEW
    if (tab == "Intranet") {
      setMyApprovalsData(MyApprovaldata);
      if (MyApprovaldata.length > 0) {
        setTimeout(() => {
          setLoading(false);
        }, 5000);
      } else {
        setLoading(false)
      }

    } else if (tab == "DMS") {
      // alert(value)
      // setMyApprovalsData(MyDMSAPPROVALDATA);
      setMyApprovalsData(Mylistdata);
      if (Mylistdata.length > 0) {
        setTimeout(() => {
          setLoading(false);
        }, 5000);
      } else {
        setLoading(false)
      }
    } else if (tab == "ProjectWorkflow") { // NEW SECTION
      setMyApprovalsData(ProjectWorkflowData);
      if (ProjectWorkflowData.length > 0) {
        setTimeout(() => {
          setLoading(false);
        }, 5000);
      } else {
        setLoading(false)
      }
    } else if (tab == "Automation") {
      setMyApprovalsData(Automationdata.sort((a: any, b: any) => b.Created - a.Created));
      if (Automationdata.length > 0) {
        setTimeout(() => {
          setLoading(false);
        }, 5000);
      } else {
        setLoading(false)
      }
      console.log("Automationdata", Automationdata);
    }

  };

  const getProjectConfiguration = async (): Promise<{ documentControllerId: number | null; dccId: number | null }> => {
    try {
      const items = await sp.web.lists
        .getByTitle("ProjectConfiguration")
        .items.select(
          "*",
          "DocumentController/ID",
          "DocumentController/Title",
          "DocumentController/EMail",
          "DCC/ID",
          "DCC/Title",
          "DCC/EMail"
        )
        .expand("DocumentController", "DCC")
        .orderBy("Created", false)();

      if (items.length > 0) {
        const documentControllerId = items[0]?.DocumentControllerId || null;
        const dccId = items[0]?.DCCId || null;

        console.log("Document Controller ID:", documentControllerId);
        console.log("DCC ID:", dccId);

        return { documentControllerId, dccId };
      }
      return { documentControllerId: null, dccId: null };
    } catch (error) {
      console.error("Error fetching Project Configuration:", error);
      return { documentControllerId: null, dccId: null };
    }
  };
  React.useEffect(() => {
    const initializeData = async () => {
      const { documentControllerId, dccId } = await getProjectConfiguration();
      setDocumentControllerId(documentControllerId);
      setDccId(dccId);
    };

    initializeData();
  }, []);


  // Add this function near your other action handlers
  const handleProjectWorkflowAction = async (e: any, item: any, mode: string) => {
    e.preventDefault();

    // For both view and approval modes, just show the form details
    await handleProjectViewClick(item);
  };

  // Handle project view click - just sets the data and shows form
  const handleProjectViewClick = async (task: any) => {
    setSelectedProjectTask(task);
    setShowProjectForm(true);
    setProjectNeedsFurtherApproval(task.DoYouNeedApproval || "Select");
    setProjectRemarks(task.Remarks || "");

    // You can keep document fetching if needed for display
    if (task.DeliverableId) {
      const docInfo = await fetchDocumentForDeliverable(task.DeliverableId, task.RevisionNumber);
      setProjectDocumentInfo(docInfo);
      // You might want to set document info state if needed for display
      // ✅ declare documentIds here 
      const documentIds: number[] = docInfo.map(d => d.id);
    }
    await getDocumentComments(task.ProjectId, task.DeliverableId, task.RevisionNumber);
    // Initialize projectHierarchy if needed
    if (task.ProjectId && task.DeliverableId) {
      await getProjectApprovalHierarchy(task.ProjectId, task.DeliverableId, task.DocType);
    } else {
      addNewProjectApprovalRow();
    }
  };

  const getProjectApprovalHierarchy = async (projectId: number, deliverableId: number, documentType: string) => {
    try {
      const items = await sp.web.lists.getByTitle("ApprovalHierarchy").items
        .select("*,DeliverablesDetailsId/ID,ProjectCreationListID/ID,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail")
        .expand("DeliverablesDetailsId,ProjectCreationListID,AssignedTo")
        .filter(`ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId}`)
        .orderBy("SerialNumber", true)();

      if (items.length > 0) {
        const hierarchyItems: ApprovalHierarchyItem[] = items.map((item: any, index: number) => {
          // AssignedTo will be an array (for multi-select people fields)
          const assignedArray = Array.isArray(item.AssignedTo) ? item.AssignedTo : [];

          return {
            id: item.Id,
            level: `Level ${index + 1}`,
            approverRole: item.ApproverRole || '',
            approver: assignedArray.map((a: any) => a.Title).join(", "),
            approvers: assignedArray.map((a: any) => a.Title),
            approvalCriteria: item.ApprovalCriteria || 'Anyone',
            serialNumber: item.SerialNumber,
            assignedTo: assignedArray.map((a: any) => ({
              ID: a.ID,
              Title: a.Title,
              EMail: a.EMail
            }))
          };
        });

        setProjectHierarchy(hierarchyItems);
      } else {
        await getProjectWorkflowConfiguration(documentType);
      }
    } catch (error) {
      console.error("Error fetching approval projectHierarchy:", error);
      await getProjectWorkflowConfiguration(documentType);
    }
  };

  // Get Project Workflow Configuration
  const getProjectWorkflowConfiguration = async (docType: string) => {
    try {
      const items = await sp.web.lists.getByTitle("ProjectWorflowConfiguration").items
        .select("*,DocumentType/ID,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail")
        .expand("DocumentType,AssignedTo")
        .filter(`DocumentType/DocumentType eq '${docType}'`)
        .orderBy("ID", true)();

      if (items.length > 0) {
        const hierarchyItems: ApprovalHierarchyItem[] = items.map((item: any, index: number) => ({
          level: `Level ${index + 1}`,
          approverRole: item.Role || '',
          approver: item.AssignedTo?.Title || '',
          approvers: item?.AssignedTo ? [item.AssignedTo.Title] : [],
          approvalCriteria: item.ApprovalCriteria || 'Anyone',
          assignedTo: item.AssignedTo ? [item.AssignedTo] : []
        }));
        setProjectHierarchy(hierarchyItems);
      } else {
        // Add one default row if no configuration found
        addNewProjectApprovalRow();
      }
    } catch (error) {
      console.error("Error fetching workflow configuration:", error);
      addNewProjectApprovalRow();
    }
  };

  // Update the handleProjectBackClick to reset all states
  const handleProjectBackClick = () => {
    setSelectedProjectTask(null);
    setShowProjectForm(false);
    setProjectHierarchy([]);
    setProjectNeedsFurtherApproval("Select");
    setProjectRemarks("");
    setProjectDocumentInfo(null);
  };

  // Simple project document open function (if needed)
  // Handle project document open with proper URL construction
  // const handleProjectOpenDocument = () => {
  //   if (!projectDocumentInfo) {
  //     console.error('No document available');
  //     alert('No document available to open.');
  //     return;
  //   }

  //   try {
  //     let documentUrl = '';

  //     // Priority 1: Use SharedLink if available
  //     if (projectDocumentInfo.sharedLink) {
  //       documentUrl = projectDocumentInfo.sharedLink;
  //       console.log('Opening document using SharedLink:', documentUrl);
  //     }
  //     // Priority 2: Use the document URL to construct the full URL
  //     else if (projectDocumentInfo.documentUrl) {
  //       const siteUrl = props.siteUrl; // Use your siteUrl from props
  //       documentUrl = `${siteUrl}${projectDocumentInfo.documentUrl}`;
  //       console.log('Opening document using ServerRelativeUrl:', documentUrl);
  //     }
  //     // Priority 3: Use fileRef if available
  //     else if (projectDocumentInfo.fileRef) {
  //       const siteUrl = props.siteUrl;
  //       documentUrl = `${siteUrl}${projectDocumentInfo.fileRef}`;
  //       console.log('Opening document using FileRef:', documentUrl);
  //     }
  //     else {
  //       throw new Error('No valid document URL found');
  //     }

  //     // Open the document in a new tab
  //     window.open(documentUrl, '_blank', 'noopener,noreferrer');

  //     console.log(`Opened document: ${projectDocumentInfo.fileName || projectDocumentInfo.fileLeafRef}`);

  //   } catch (error) {
  //     console.error('Error opening document:', error);
  //     alert('Error opening document. Please try again or contact administrator.');
  //   }
  // };
  // const handleProjectOpenDocument = () => {
  //   if (!projectDocumentInfo) {
  //     alert("No document available");
  //     return;
  //   }

  //   try {
  //     let documentUrl = "";

  //     // 1️⃣ Shared link – direct open
  //     if (projectDocumentInfo.sharedLink) {
  //       documentUrl = projectDocumentInfo.sharedLink;
  //     }

  //     // 2️⃣ ServerRelativeUrl or documentUrl
  //     else if (projectDocumentInfo.documentUrl || projectDocumentInfo.fileRef) {
  //       const serverRelativeUrl =
  //         projectDocumentInfo.documentUrl || projectDocumentInfo.fileRef;

  //       // 👉 tenant root extract karo
  //       const tenantRoot = props.siteUrl.split("/sites")[0];

  //       documentUrl = `${tenantRoot}${serverRelativeUrl}`;
  //     } else {
  //       throw new Error("No valid document URL found");
  //     }

  //     window.open(documentUrl, "_blank", "noopener,noreferrer");
  //   } catch (error) {
  //     console.error(error);
  //     alert("Error opening document");
  //   }
  // };
  // Function to fetch document for a specific deliverable
  // const fetchDocumentForDeliverable = async (deliverableId: number) => {
  //   try {
  //     const documents = await sp.web.lists.getByTitle("DeliverablesDocument").items
  //       .select("*,File/ServerRelativeUrl,FileLeafRef,FileRef,Author/ID,Author/Title,SharedLink")
  //       .expand("File", "Author")
  //       .filter(`DeliverablesDetailsId eq ${deliverableId}`)
  //       .orderBy("ID", false)
  //       .top(1)(); // Get the latest document

  //     if (documents.length > 0) {
  //       const latestDoc = documents[0];
  //       return {
  //         documentUrl: latestDoc.File.ServerRelativeUrl,
  //         fileName: latestDoc.FileLeafRef,
  //         fileLeafRef: latestDoc.FileLeafRef,
  //         fileRef: latestDoc.FileRef,
  //         sharedLink: latestDoc.SharedLink // Get the SharedLink column value
  //       };
  //     } else {
  //       console.warn(`No documents found for deliverable ID: ${deliverableId}`);
  //       return null;
  //     }
  //   } catch (error) {
  //     console.error(`Error fetching documents for deliverable ${deliverableId}:`, error);
  //     return null;
  //   }
  // };
  const fetchDocumentForDeliverable = async (deliverableId: number, revisionNo: number) => {
    try {
      const documents = await sp.web.lists.getByTitle("DeliverablesDocument").items
        .select("Id,Revision,File/ServerRelativeUrl,FileLeafRef,FileRef,Author/ID,Author/Title,SharedLink")
        .expand("File", "Author")
        .filter(`DeliverablesDetailsId eq ${deliverableId} and Revision eq '${revisionNo.toString()}'`)
        .orderBy("ID", false)(); // remove .top(1) to get all

      if (documents.length > 0) {
        // Map each document into a simplified object
        return documents.map(doc => ({
          documentUrl: doc.File?.ServerRelativeUrl,
          fileName: doc.FileLeafRef,
          fileLeafRef: doc.FileLeafRef,
          fileRef: doc.FileRef,
          sharedLink: doc.SharedLink,
          author: doc.Author?.Title,
          authorId: doc.Author?.ID,
          id: doc.Id
        }));
      } else {
        console.warn(`No documents found for deliverable ID: ${deliverableId}`);
        return [];
      }
    } catch (error) {
      console.error(`Error fetching documents for deliverable ${deliverableId}:`, error);
      return [];
    }
  };



  // Fetch document comment
  const getDocumentComments = async (projectCreationID: number, deliverableDetailsID: number, revision: string) => {
    try {
      const items = await sp.web.lists.getByTitle("DocumentComments").items
        .select("*")
        .filter(`ProjectID eq ${projectCreationID} and DeliverableDetailsID/ID eq ${deliverableDetailsID}`)
        .orderBy("ID", false)();
      if (items.length > 0) {
        const comments: DocumentComment[] = items.map((item: any) => ({
          id: item.Id,
          userName: item.UserName || "",
          commentDate: item.CommentDate ? new Date(item.CommentDate).toLocaleString('en-GB') : "",
          pageNumber: item.PageNumber || "",
          revision: item.Revision || "",
          comment: item.Comment || ""
        }));
        setAllDocumentComments(comments);
        console.log("All Document Comments:", comments);

        const filteredComments = comments.filter(comment => comment.revision === revision);
        setDocumentComments(filteredComments);

        const uniqueVersions = [...new Set(comments.map(comment => comment.revision))].sort();
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
      setShowDocumentComments(false);
    }
  };

  const onVersionChange = (version: string) => {
    setSelectedVersion(version);
    if (!version) {
      setDocumentComments(allDocumentComments);
    } else {
      const filteredComments = allDocumentComments.filter(comment => comment.revision === version);
      setDocumentComments(filteredComments);
    }
  };

  const exportCommentsToExcel = () => {
    const headers = ['Users', 'Comment Date', 'Page No.', 'Revision', 'Comments/Clarifications'];
    const csvContent = [
      headers.join(','),
      ...documentComments.map(comment => [
        `"${comment.userName}"`,
        `"${comment.commentDate}"`,
        `"${comment.pageNumber}"`,
        `"${comment.revision}"`,
        `"${comment.comment}"`
      ].join(','))
    ].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `DocumentComments_${selectedProjectTask?.DocNumber || 'export'}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  const refreshDocComment = () => {
    if (selectedProjectTask) {
      getDocumentComments(selectedProjectTask.ProjectId, selectedProjectTask.DeliverableId, selectedProjectTask.RevisionNumber);
    }
  };

  // document comment ends here 

  // Handle individual item selection
  // const handleProjectItemSelect = (requestId: string) => {
  //   setSelectedProjectItems(prev => {
  //     if (prev.includes(requestId)) {
  //       return prev.filter(id => id !== requestId);
  //     } else {
  //       return [...prev, requestId];
  //     }
  //   });
  // };
  const handleProjectItemSelect = async (id: number) => {
    // check if already selected
    const alreadySelected = selectedProjectItems.some(item => item.Id === id);
    if (alreadySelected) {
      setSelectedProjectItems(prev => prev.filter(item => item.Id !== id));
      return;
    }

    // find base details
    const details = currentData?.find(item => item.Id === id);
    if (!details) return;

    // fetch additional project creation data if ProjectCreationListID exists
    let creationItem: any = null;
    if (details.ProjectId) {
      creationItem = await sp.web.lists
        .getByTitle("ProjectCreationList")
        .items.getById(details.ProjectId)
        .select(
          "*",
          "PreparedBy/ID",
          "PreparedBy/Title",
          "ProjectType/ProjectType",
          "ProjectType/Id",
          "ClientName"
        )
        .expand("ProjectType", "PreparedBy")();
    }

    // fetch additional deliverables data if DeliverablesDetails exists
    let DeliverablesItem: any = null;
    if (details.DeliverableId) {
      console.log("DeliverableId value:", details.DeliverableId, typeof details.DeliverableId);
      DeliverablesItem = await sp.web.lists
        .getByTitle("DeliverablesDetails")
        .items.getById(details.DeliverableId)
        .select(
          "*",
          "AssignedTo/ID",
          "AssignedTo/Title"
        )
        .expand("AssignedTo")();
    }

    // build full object
    const fullDetails: ProjectItem = {
      Id: details.Id,
      Status: details.Status,
      Author: { Title: details.Author?.Title || "" },
      Created: details.Created,
      InitiatedBy: creationItem?.AuthorId ?? creationItem?.AuthorId,
      documentNumber: DeliverablesItem?.DocNumber ?? DeliverablesItem?.DocNumber,
      VendorName: DeliverablesItem?.AssignedTo.Title ?? DeliverablesItem?.AssignedTo.Title,
      docrevisionNumber: DeliverablesItem?.RevisionNumber ?? DeliverablesItem?.RevisionNumber,
      VendorId: DeliverablesItem?.AssignedToId ?? DeliverablesItem?.AssignedToId,
      ProjectName: creationItem.ProjectName || "",
      deliverablesDocumentIDs: DeliverablesItem?.DeliverablesDocumentIDId ?? DeliverablesItem?.DeliverablesDocumentIDId,
      ProjectId: creationItem?.ID,
      DeliverableId: DeliverablesItem?.ID,
      PreparedBy: creationItem?.PreparedBy?.Title || details?.PreparedBy?.Title || "",
      PreparedById: creationItem?.PreparedBy?.ID || details?.PreparedBy?.ID || "",
      DocType: DeliverablesItem.DocumentType || "",
      ApprovalRole: details.ApprovalRole || "",
      RevisionNumber: details.RevisionNumber || "0",
      Deliverable: DeliverablesItem.Deliverables || "",
      // Requester: {
      //   Title: details.Author?.Title || "",
      //   EMail: details.Author?.EMail || ""
      // },
      // RequestID: details.DocNumber || `PROJ-${details.Id}`,
      // Title: creationItem.ProjectName || "Project Approval",
      // ApprovalTitle: creationItem.ProjectName || "Project Approval",
      // ProcessName: "Project Workflow",
      // ProjectType: creationItem?.ProjectType?.ProjectType || details?.ProjectType?.ProjectType || "",
      // ClientName: creationItem?.ClientName || details?.ClientName || "",
      // Area: DeliverablesItem?.Area || "",
      // DocNumber: DeliverablesItem.DocNumber || "",
      // DoYouNeedApproval: details.Doyouneedapproval || "",
      // CurrentApprovalLevel: details.Level || "",
      // ApprovalSN: details.SerialNumber || 0,
      // ApprovalCriteria: details.ApprovalCriteria || "",
      // AssignedTo: DeliverablesItem?.AssignedTo?.Title || details.AssignedTo?.Title || "",
      // Org: details.DeliverablesDetailsId?.Organization || "",
      // DocumentNumber: details.DocNumber || "",
      // ProjectDate: details.Created ? new Date(details.Created).toLocaleDateString("en-GB") : "",
      // Remarks: details.Remarks || "",
      // RedirectionLink: details.RedirectionLink || "",
      // SNo: selectedProjectItems.length + 1
    };

    // finally update state
    setSelectedProjectItems(prev => [...prev, fullDetails]);
  };





  // Handle select all
  const handleSelectAllProjects = () => {
    if (currentData && currentData.length > 0) {
      if (selectedProjectItems.length === currentData.length) {
        // If all are selected, deselect all
        setSelectedProjectItems([]);
      } else {
        // Select all current page items
        const allIds = currentData.map(item => item.RequestID);
        setSelectedProjectItems(allIds);
      }
    }
  };


  // const handlebulkReworkAction = async () => {
  //   if (selectedProjectItems.length === 0) {
  //     alert("Please select at least one item before performing Rework.");
  //     return;
  //   }
  //   // 1. Show Loader
  //   setShowSubmitLoader(true);
  //   // 👉 Put your actual rework logic here
  //   console.log("Rework action performed on:", selectedProjectItems);
  //   let actionExecuted = false; // ✅ Flag to control alert

  //   await processVendorRework();

  //   actionExecuted = true;
  //   // if (actionExecuted) { // ✅ Only show alert if something executed
  //   //   alert(`Rework successfully!`);
  //   //   window.location.reload(); // optional refresh
  //   // } else {
  //   //   console.warn("⚠️ No matching action found for status:", status);
  //   // }
  //   // Example: call API, update state, etc.
  // };


  const handlebulkReworkAction = async () => {
    if (selectedProjectItems.length === 0) {
      alert("Please select at least one item before performing Rework.");
      return;
    }

    // 1. Show Loader
    setShowSubmitLoader(true);

    try {
      console.log("Rework action performed on:", selectedProjectItems);

      // 2. Execute the processing function
      await processVendorRework();

      // 3. Configure Success Popup (This will trigger after the loop in processVendorRework finishes)
      setPopupType("success");
      setPopupTitle("Bulk Rework Success");
      setPopupMessage(`Successfully processed rework for ${selectedProjectItems.length} deliverable(s) and notified the relevant parties.`);
      setPopupOpen(true);

    } catch (error) {
      console.error("Error in bulk rework:", error);
      setPopupType("error");
      setPopupTitle("Error");
      setPopupMessage("Failed to process bulk rework. Please try again.");
      setPopupOpen(true);
    } finally {
      // 4. Hide Loader
      setShowSubmitLoader(false);
    }
  };
  const processVendorRework = async () => {
    if (!selectedProjectItems || selectedProjectItems.length === 0) return;

    const currentUserId = props.context.pageContext.legacyPageContext.userId;

    for (const task of selectedProjectItems) {
      try {
        // 1️⃣ Update DeliverablesDetails status + mark reworked
        const deliverableUpdate: any = {
          Status: "Pending",
          IsReworked: "Yes"
        };

        await sp.web.lists
          .getByTitle("DeliverablesDetails")
          .items.getById(task.DeliverableId!)
          .update(deliverableUpdate);

        // 2️⃣ Update ProjectApprovals only for Rework
        if (task.ApprovalRole === "Document Controller") {
          const approvalUpdate: any = {
            Status: "Rework",
            Remarks: projectRemarks,
            ApprovalDate: new Date(),
            OutgoingDate: new Date(),
            MDRStatus: mdrStatus
            // MDRStatus: ["A", "B", "C"][Math.floor(Math.random() * 3)]
          };

          if (projectNeedsFurtherApproval !== "Select") {
            approvalUpdate.Doyouneedapproval = projectNeedsFurtherApproval;
          }

          await sp.web.lists
            .getByTitle("ProjectApprovals")
            .items.getById(task.Id)
            .update(approvalUpdate);
        }

        // 3️⃣ Create vendor approval task
        const approvalData = {
          ProjectCreationListIDId: task.ProjectId,
          DeliverablesDetailsIdId: task.DeliverableId,
          AssignedToId: task.VendorId,
          RequestedById: currentUserId,
          RequestedDate: new Date(),
          DocumentType: task.DocType,
          ApproverRole: "Vendor",
          ApprovalCriteria: "Anyone",
          Status: "pending",
          Level: "Level 0",
          SerialNumber: 0,
          RevisionNumber: task.RevisionNumber,
        };

        await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);

        // 4️⃣ Trigger vendor email
        const fileDetails: { name: string; link: string, docrevisionNo: string }[] = [];
        for (const id of task.deliverablesDocumentIDs) {
          const details = await waitForShareLink(id);
          // Clean the file name by removing the prefix before first underscore
          const originalName = details.fileName;
          const cleanedName = originalName.includes("_")
            ? originalName.substring(originalName.indexOf("_") + 1)
            : originalName;
          fileDetails.push({
            name: cleanedName,
            link: details.shareLink,
            docrevisionNo: details.revisionNo
          });
        }

        const actionType: ActionType = "REWORK_BY_DOCUMENT_CONTROLLER";
        //  to get the initiator name
        var CCUserName = Array.isArray(users)
          ? users.filter(u => String(u.value) === String(task.InitiatedBy))
          : [];
        var CCName = CCUserName.length > 0 ? CCUserName[0].label : null;

        const emailBody = buildbulkApprovalEmailBody(
          actionType,
          task.documentNumber,
          task.VendorName, // whoever is receiving
          CCName,   // whoever is sending
          task.ProjectName,
          fileDetails.map(d => ({
            deliverable: task?.Deliverable,
            fileName: d.name,
            sharedLink: d.link,
            docrevisionNo: d.docrevisionNo
          }))
        );

        await createbulkEmailTriggerDetails(
          task.VendorId,
          [task.InitiatedBy],
          // `The submitted document has been reviewed by the Document Controller and requires rework for the Deliverable <b>"${task.Deliverable}"</b>. Please review the feedback, make the necessary corrections, and resubmit the document at the earliest.`,
          emailBody,
          `Rework Action Assigned – ${task.Deliverable}`,
          task.deliverablesDocumentIDs,
          task.Deliverable
        );

      } catch (err) {
        console.error("Error processing vendor rework for task:", task.Id, err);
      }
    }
  };



  // Project Approval Code starts here ----------------
  const handleProjectApprovalAction = async (status: "Approved" | "Rejected" | "Rework") => {
    setShowSubmitLoader(true);

    if (!selectedProjectTask) {
      setShowSubmitLoader(false); // 🔴 safety close
      return;
    }

    try {
      // Update the current approval status
      await updateCurrentApprovalStatus(status);
      console.log("Status updated successfully, now checking status type:", status);

      let actionExecuted = false; // ✅ Flag to control alert

      if (status === "Approved") {
        await handleApprovedAction();
        actionExecuted = true;
      } else if (status === "Rework") {
        await handleReworkAction();
        actionExecuted = true;
      } else if (status === "Rejected") {
        await handleRejectedAction();
        actionExecuted = true;
      }

      if (actionExecuted) { // ✅ Only show alert if something executed
        //alert(`${status} successfully!`);
        // handleProjectBackClick();
        // window.location.reload(); // optional refresh
        setPopupType("success");
        setPopupTitle("Success");
        setPopupMessage(`Request ${status} successfully.`);
        setPopupOpen(true);
      } else {
        console.warn("⚠️ No matching action found for status:", status);
      }
    } catch (error) {
      console.error("Error in approval action:", error);
      //alert("Something went wrong while processing approval.");
    }
    finally {
      // ✅ ALWAYS close loader here
      setShowSubmitLoader(false);
    }
  };
  // Update current approval status
  const updateCurrentApprovalStatus = async (status: string) => {
    if (!selectedProjectTask) return;

    try {
      const data: any = {
        Status: status,
        Remarks: projectRemarks,
        ApprovalDate: new Date(),
      };

      if (projectNeedsFurtherApproval !== "Select") {
        data.Doyouneedapproval = projectNeedsFurtherApproval;
      }

      if (status === "Rework" && selectedProjectTask.ApprovalRole === "Document Controller") {
        data.OutgoingDate = new Date();
        const statuses = ["A", "B", "C"];
        // data.MDRStatus = statuses[Math.floor(Math.random() * statuses.length)];
        data.MDRStatus = projectOutgoingstatus;
      }

      await sp.web.lists.getByTitle("ProjectApprovals").items.getById(selectedProjectTask.Id).update(data);
      console.log("Status Updated");
    } catch (error) {
      console.error("Error in updateCurrentApprovalStatus:", error);
      throw error; // Re-throw to be caught in main handler
    }
  };

  // Handle Approved action
  const handleApprovedAction = async () => {
    if (!selectedProjectTask) return;

    const approvals = await getAllProjectApproval(selectedProjectTask.ProjectId!, selectedProjectTask.DeliverableId!);
    console.log('Array?', Array.isArray(approvals));
    await addMultipleApproval("Approved", approvals);
  };

  // Get all project approvals
  const getAllProjectApproval = async (projectId: number, deliverableId: number) => {
    try {
      const items = await sp.web.lists.getByTitle("ProjectApprovals").items
        .select("*,DeliverablesDetailsId/ID,ProjectCreationListID/ID,AssignedTo/ID,AssignedTo/Title")
        .expand("DeliverablesDetailsId,ProjectCreationListID,AssignedTo")
        .filter(`ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId}`)
        .orderBy("Created", false)();

      // setProjectApprovalArr(items);
      return items;
    } catch (error) {
      console.error("Error fetching project approvals:", error);
      return [];
    }
  };

  // Add Multiple Approval logic
  const addMultipleApproval = async (statusUpdate: string, allApprovals: any[]) => {
    if (!selectedProjectTask) return;
    console.log('Array?', Array.isArray(allApprovals), allApprovals.length);
    try {
      if (statusUpdate === "Approved") {
        const currentRole = selectedProjectTask.ApprovalRole;
        const currentCriteria = selectedProjectTask.ApprovalCriteria;

        if (currentRole === "Document Controller" && projectNeedsFurtherApproval === "Yes") {
          // Create projectHierarchy and approvals for multiple approvers
          for (let index = 0; index < projectHierarchy.length; index++) {
            const row = projectHierarchy[index];
            const approverAssignOwner = row.assignedTo?.map(user => user.ID) || [];

            // Update or create projectHierarchy
            const rowData = {
              ProjectCreationListIDId: selectedProjectTask.ProjectId,
              DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
              // AssignedToId: { 'results': approverAssignOwner },
              AssignedToId: approverAssignOwner,
              RequestedById: props.context.pageContext.legacyPageContext.userId,
              RequestedDate: new Date(),
              DocumentType: selectedProjectTask.DocType,
              ApproverRole: row.approverRole,
              ApprovalCriteria: row.approvalCriteria,
              Status: 'Pending',
              Doyouneedapproval: projectNeedsFurtherApproval,
              Level: row.level,
              SerialNumber: row.serialNumber || index + 1,
            };

            if (row.id) {
              await sp.web.lists.getByTitle("ApprovalHierarchy").items.getById(row.id).update(rowData);
            } else {
              await sp.web.lists.getByTitle("ApprovalHierarchy").items.add(rowData);
            }

            // Create approvals for Level 1
            if (row.level === "Level 1") {
              for (const userId of approverAssignOwner) {
                const approvalData = {
                  ProjectCreationListIDId: selectedProjectTask.ProjectId,
                  DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
                  AssignedToId: userId,
                  RequestedById: props.context.pageContext.legacyPageContext.userId,
                  RequestedDate: new Date(),
                  DocumentType: selectedProjectTask.DocType,
                  ApproverRole: row.approverRole,
                  ApprovalCriteria: row.approvalCriteria,
                  Doyouneedapproval: projectNeedsFurtherApproval,
                  Status: 'Pending',
                  Level: row.level,
                  SerialNumber: row.serialNumber || index + 1,
                  RevisionNumber: selectedProjectTask.RevisionNumber,
                };
                await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);

                const actionType: ActionType = "APPROVAL_ASSIGNED";
                // to get the approver name 
                var TouserUsers = Array.isArray(users)
                  ? users.filter(u => String(u.value) === String(userId))
                  : [];
                var TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
                //  to get the initiator name
                var CCUserName = Array.isArray(users)
                  ? users.filter(u => String(u.value) === String(selectedProjectTask.InitiatedBy))
                  : [];
                var CCName = CCUserName.length > 0 ? CCUserName[0].label : null;

                const emailBody = buildApprovalEmailBody(
                  actionType,
                  TouserName, // whoever is receiving
                  CCName,   // whoever is sending
                  projectDocumentInfo.map(d => ({
                    deliverable: selectedProjectTask.Deliverable,
                    fileName: d.fileName,
                    sharedLink: d.sharedLink
                  }))
                );
                await createEmailTriggerDetails(
                  userId,
                  [selectedProjectTask.InitiatedBy],
                  emailBody,
                  `${actionConfig[actionType].subject} – ${selectedProjectTask.Deliverable}`,
                  projectDocumentInfo.map(d => d.id)
                );

                // await createEmailTriggerDetails(
                //   userId,
                //   [selectedProjectTask.InitiatedBy],
                //   `An approval action has been assigned to you by the Document Controller for the Deliverable <b>"${selectedProjectTask.Deliverable}"</b>. Please review the request and proceed with the required actions at the earliest.`,
                //   `Approval Action Assigned – ${selectedProjectTask.Deliverable}`,
                //   projectDocumentInfo.map(d => d.id)
                // );
              }
            }
          }
        } else if (currentRole === "Document Controller" && projectNeedsFurtherApproval === "No") {
          // Create DCC approval
          const approvalData = {
            ProjectCreationListIDId: selectedProjectTask.ProjectId,
            DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
            AssignedToId: dccId,
            RequestedById: props.context.pageContext.legacyPageContext.userId,
            RequestedDate: new Date(),
            DocumentType: selectedProjectTask.DocType,
            ApproverRole: "DCC",
            ApprovalCriteria: "Anyone",
            Status: 'Pending',
            Level: `Level ${selectedProjectTask.ApprovalSN + 1}`,
            SerialNumber: selectedProjectTask.ApprovalSN + 1,
            RevisionNumber: selectedProjectTask.RevisionNumber,
          };
          await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
          const actionType: ActionType = "SENT_FOR_PUBLISH";  //publish assign

          // to get the approver name 
          var TouserUsers = Array.isArray(users)
            ? users.filter(u => String(u.value) === String(dccId))
            : [];
          var TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
          //  to get the initiator name
          var CCUserName = Array.isArray(users)
            ? users.filter(u => String(u.value) === String(selectedProjectTask.InitiatedBy))
            : [];
          var CCName = CCUserName.length > 0 ? CCUserName[0].label : null;
          // end 

          const emailBody = buildApprovalEmailBody(
            actionType,
            TouserName, // whoever is receiving
            CCName,   // whoever is sending
            projectDocumentInfo.map(d => ({
              deliverable: selectedProjectTask.Deliverable,
              fileName: d.fileName,
              sharedLink: d.sharedLink
            }))
          );
          await createEmailTriggerDetails(
            dccId,
            [selectedProjectTask.InitiatedBy],
            emailBody,
            `${actionConfig[actionType].subject} – ${selectedProjectTask.Deliverable}`,
            projectDocumentInfo.map(d => d.id)
          );
          //Satish added
          await updateDeliverableStatus(statusUpdate, "");
        } else {
          // Regular approval flow
          const currentApprovals = allApprovals.filter(x =>
            x.ApproverRole === currentRole &&
            x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
            x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId
          );

          const allApproved = currentApprovals.length > 0 &&
            currentApprovals.every(item => item.Status === "Approved" || item.Status === "Rework");

          // Update current approvals if criteria = "Anyone"
          if (currentCriteria === "Anyone" && allApprovals.length > 0) {
            for (const approvalItem of allApprovals) {
              if (approvalItem.Id && approvalItem.ApproverRole === currentRole) {
                const updateData = {
                  Status: statusUpdate,
                  ApprovalDate: new Date(),
                };
                await sp.web.lists.getByTitle("ProjectApprovals").items.getById(approvalItem.Id).update(updateData);
                const actionType: ActionType = "APPROVED";  //approved

                // to get the approver name 
                var TouserUsers = Array.isArray(users)
                  ? users.filter(u => String(u.value) === String(documentControllerId))
                  : [];
                var TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
                //  to get the initiator name
                var CCUserName = Array.isArray(users)
                  ? users.filter(u => String(u.value) === String(selectedProjectTask.InitiatedBy))
                  : [];
                var CCName = CCUserName.length > 0 ? CCUserName[0].label : null;
                // end 

                const emailBody = buildApprovalEmailBody(
                  actionType,
                  TouserName, // whoever is receiving
                  CCName,   // whoever is sending
                  projectDocumentInfo.map(d => ({
                    deliverable: selectedProjectTask.Deliverable,
                    fileName: d.fileName,
                    sharedLink: d.sharedLink
                  }))
                );
                await createEmailTriggerDetails(
                  documentControllerId,
                  [selectedProjectTask.InitiatedBy],
                  emailBody,
                  `${actionConfig[actionType].subject} – ${selectedProjectTask.Deliverable}`,
                  projectDocumentInfo.map(d => d.id)
                );
                //Satish added
              }
            }
            await updateApprovalHierarchyStatus(allApprovals);
          }

          const shouldProceed = (currentCriteria === "Everyone" && allApproved) ||
            (currentCriteria === "Anyone");

          if (shouldProceed) {
            await handleNextApprovers();
          }
        }
      }
    } catch (error) {
      console.error("Error in addMultipleApproval:", error);
      throw error;
    }
  };

  // const updateDeliverableStatus = async (status: string) => {
  //   if (!selectedProjectTask) return;

  //   await sp.web.lists.getByTitle("DeliverablesDetails")
  //     .items.getById(selectedProjectTask.DeliverableId!)
  //     .update({ Status: status });
  // };

  const updateDeliverableStatus = async (status: string, isRework?: string) => {
    if (!selectedProjectTask) return;

    // Build payload
    const data: any = { Status: status };

    if (isRework) {
      data.IsReworked = isRework;
    }

    await sp.web.lists
      .getByTitle("DeliverablesDetails")
      .items.getById(selectedProjectTask.DeliverableId!)
      .update(data);
  };

  const updateApprovalHierarchyStatus = async (allApprovals: any[]) => {
    if (!selectedProjectTask) return;

    try {
      const currentRole = selectedProjectTask.ApprovalRole;
      const currentCriteria = selectedProjectTask.ApprovalCriteria;

      const currentApprovals = allApprovals.filter(x =>
        x.ApproverRole === currentRole &&
        x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
        x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId
      );

      if (currentApprovals.length === 0) return;

      const allApproved = currentApprovals.every(item => item.Status === "Approved");
      let newStatus = "";

      if (currentCriteria === "Anyone") {
        newStatus = "Approved";
      } else if (currentCriteria === "Everyone") {
        newStatus = allApproved ? "Approved" : "In Progress";
      }

      const currentHierarchy = projectHierarchy.find(x =>
        x.approverRole === currentRole &&
        x.id &&
        x.id === selectedProjectTask!.deliverableId
      );

      if (currentHierarchy && currentHierarchy.id && newStatus !== "") {
        await sp.web.lists.getByTitle("ApprovalHierarchy")
          .items.getById(currentHierarchy.id)
          .update({ Status: newStatus });
      }
    } catch (error) {
      console.error("Error in updateApprovalHierarchyStatus:", error);
    }
  };

  // Handle next approvers
  const handleNextApprovers = async () => {
    if (!selectedProjectTask) return;

    const currentHierarchy = projectHierarchy.find(x =>
      x.approverRole === selectedProjectTask.ApprovalRole &&
      x.level === selectedProjectTask.CurrentApprovalLevel
    );

    const nextApprovers = currentHierarchy ?
      projectHierarchy.filter(x => x.serialNumber === (currentHierarchy.serialNumber! + 1)) : [];

    if (!nextApprovers || nextApprovers.length === 0) {
      // Final stage
      if (selectedProjectTask.ApprovalRole !== "DCC") {
        const approvalData = {
          ProjectCreationListIDId: selectedProjectTask.ProjectId,
          DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
          AssignedToId: dccId,
          RequestedById: props.context.pageContext.legacyPageContext.userId,
          RequestedDate: new Date(),
          DocumentType: selectedProjectTask.DocType,
          ApproverRole: "DCC",
          ApprovalCriteria: "Anyone",
          Status: 'Pending',
          Level: `Level ${selectedProjectTask.ApprovalSN + 1}`,
          SerialNumber: selectedProjectTask.ApprovalSN + 1,
          RevisionNumber: selectedProjectTask.RevisionNumber,
        };
        await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
        await updateDeliverableStatus("Approved", "");
        const actionType: ActionType = "SENT_FOR_PUBLISH";  //publish

        // to get the approver name 
        var TouserUsers = Array.isArray(users)
          ? users.filter(u => String(u.value) === String(dccId))
          : [];
        var TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
        //  to get the initiator name
        var CCUserName = Array.isArray(users)
          ? users.filter(u => String(u.value) === String(selectedProjectTask.InitiatedBy))
          : [];
        var CCName = CCUserName.length > 0 ? CCUserName[0].label : null;
        // end 


        const emailBody = buildApprovalEmailBody(
          actionType,
          TouserName, // whoever is receiving
          CCName,   // whoever is sending
          projectDocumentInfo.map(d => ({
            deliverable: selectedProjectTask.Deliverable,
            fileName: d.fileName,
            sharedLink: d.sharedLink
          }))
        );
        await createEmailTriggerDetails(
          dccId,
          [selectedProjectTask.InitiatedBy],
          emailBody,
          `${actionConfig[actionType].subject} – ${selectedProjectTask.Deliverable}`,
          projectDocumentInfo.map(d => d.id)
        );
        // await createEmailTriggerDetails(
        //   dccId,
        //   [selectedProjectTask.InitiatedBy],
        //   `A publish action has been assigned to you by the Document Controller for the Deliverable <b>"${selectedProjectTask.Deliverable}"</b>. Please review the document and proceed with the publishing process as required.`,
        //   `Publish Action Assigned – ${selectedProjectTask.Deliverable}`,
        //   projectDocumentInfo.map(d => d.id)
        // );
        //Satish added
      } else if (selectedProjectTask?.ApprovalRole === "DCC" && projectWantsToPublishInDossier !== "Select") {
        await updateDeliverablePublishStatus(selectedProjectTask.DeliverableId!, projectWantsToPublishInDossier, projectDocumentInfo.map(d => d.id));
      }
    } else {
      // Create approvals for next approvers
      // for (const approvalItem of nextApprovers) {
      //     if (approvalItem.assignedTo && approvalItem.assignedTo.length > 0) {
      //         for (const user of approvalItem.assignedTo) {
      //             const approvalData = {
      //                 ProjectCreationListIDId: selectedProjectTask.ProjectId,
      //                 DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
      //                 AssignedToId: user[0].ID,
      //                 RequestedById: context.pageContext.legacyPageContext.userId,
      //                 RequestedDate: new Date(),
      //                 DocumentType: selectedProjectTask.DocType,
      //                 ApproverRole: approvalItem.approverRole,
      //                 ApprovalCriteria: approvalItem.approvalCriteria,
      //                 Status: 'Pending',
      //                 Level: approvalItem.level,
      //                 SerialNumber: approvalItem.serialNumber,
      //             };
      //             await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
      //         }
      //     }
      // }

      for (const approvalItem of nextApprovers) {
        if (approvalItem.assignedTo && approvalItem.assignedTo.length > 0) {

          // Flatten in case it's an array of arrays
          const allUsers = ([] as any[]).concat.apply([], approvalItem.assignedTo);

          for (const user of allUsers) {
            const approvalData = {
              ProjectCreationListIDId: selectedProjectTask.ProjectId,
              DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
              AssignedToId: user.ID,
              RequestedById: props.context.pageContext.legacyPageContext.userId,
              RequestedDate: new Date(),
              DocumentType: selectedProjectTask.DocType,
              ApproverRole: approvalItem.approverRole,
              ApprovalCriteria: approvalItem.approvalCriteria,
              Status: 'Pending',
              Level: approvalItem.level,
              SerialNumber: approvalItem.serialNumber,
              RevisionNumber: selectedProjectTask.RevisionNumber,
            };

            await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
            const actionType: ActionType = "APPROVAL_ASSIGNED";  //approval

            // to get the approver name 
            var TouserUsers = Array.isArray(users)
              ? users.filter(u => String(u.value) === String(user.ID))
              : [];
            var TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
            //  to get the initiator name
            var CCUserName = Array.isArray(users)
              ? users.filter(u => String(u.value) === String(selectedProjectTask.InitiatedBy))
              : [];
            var CCName = CCUserName.length > 0 ? CCUserName[0].label : null;
            // end 


            const emailBody = buildApprovalEmailBody(
              actionType,
              TouserName, // whoever is receiving
              CCName,   // whoever is sending
              projectDocumentInfo.map(d => ({
                deliverable: selectedProjectTask.Deliverable,
                fileName: d.fileName,
                sharedLink: d.sharedLink
              }))
            );
            await createEmailTriggerDetails(
              user.ID,
              [selectedProjectTask.InitiatedBy, documentControllerId],
              emailBody,
              `${actionConfig[actionType].subject} – ${selectedProjectTask.Deliverable}`,
              projectDocumentInfo.map(d => d.id)
            );
            // await createEmailTriggerDetails(
            //   user.ID,
            //   [selectedProjectTask.InitiatedBy, documentControllerId],
            //   `An approval action has been assigned to you by the Document Controller for the Deliverable <b>"${selectedProjectTask.Deliverable}"</b>. Please review the document and proceed with the required approval actions.`,
            //   `Approval Action Assigned – ${selectedProjectTask.Deliverable}`,
            //   projectDocumentInfo.map(d => d.id)
            // );
            //Satish added
          }
        }
      }

    }
  };


  // Handle Rework action
  const handleReworkAction = async () => {
    if (!selectedProjectTask) return;

    if (selectedProjectTask.ApprovalRole !== "Document Controller") {
      const approvals = await getAllProjectApproval(selectedProjectTask.ProjectId!, selectedProjectTask.DeliverableId!);

      // Update other pending approvals in current level to Rework
      const pendingApprovals = approvals.filter(x =>
        x.Level === selectedProjectTask?.CurrentApprovalLevel &&
        x.ApproverRole === selectedProjectTask?.ApprovalRole &&
        x.Id !== selectedProjectTask.Id &&
        x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
        x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId
      );

      for (const item of pendingApprovals) {
        const data = {
          Status: "Rework",
          Remarks: "Auto Rework",
          ApprovalDate: new Date(),
        };
        await sp.web.lists.getByTitle("ProjectApprovals").items.getById(item.Id).update(data);
      }

      // Create new approval for Document Controller
      await createDocumentControllerApproval();
      await updateApprovalHierarchyStatusForRework("Rework");

    } else {
      // Document Controller Rework logic
      await updateDeliverableStatus("Pending", "Yes");
      await createVendorApproval();
      // await sendEmail(); // Uncomment if you have email functionality
    }
  };


  const createDocumentControllerApproval = async () => {
    if (!selectedProjectTask) return;

    //const documentControllerId = 6 // Implement this

    const approvalData = {
      ProjectCreationListIDId: selectedProjectTask.ProjectId,
      DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
      AssignedToId: documentControllerId,
      RequestedById: props.context.pageContext.legacyPageContext.userId,
      RequestedDate: new Date(),
      DocumentType: selectedProjectTask.DocType,
      ApproverRole: "Document Controller",
      ApprovalCriteria: "Everyone",
      Doyouneedapproval: "Yes",
      Status: 'Pending',
      Level: "Level 1",
      SerialNumber: 1,
      RevisionNumber: selectedProjectTask.RevisionNumber,
    };
    await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
    const actionType: ActionType = "REWORK_BY_APPROVER";  //rework by approver

    // to get the approver name 
    var TouserUsers = Array.isArray(users)
      ? users.filter(u => String(u.value) === String(documentControllerId))
      : [];
    var TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
    //  to get the initiator name
    var CCUserName = Array.isArray(users)
      ? users.filter(u => String(u.value) === String(selectedProjectTask.InitiatedBy))
      : [];
    var CCName = CCUserName.length > 0 ? CCUserName[0].label : null;
    // end 

    const emailBody = buildApprovalEmailBody(
      actionType,
      TouserName, // whoever is receiving
      CCName,   // whoever is sending
      projectDocumentInfo.map(d => ({
        deliverable: selectedProjectTask.Deliverable,
        fileName: d.fileName,
        sharedLink: d.sharedLink
      }))
    );
    await createEmailTriggerDetails(
      documentControllerId,
      [selectedProjectTask.InitiatedBy],
      emailBody,
      `${actionConfig[actionType].subject} – ${selectedProjectTask.Deliverable}`,
      projectDocumentInfo.map(d => d.id)
    );
    // await createEmailTriggerDetails(
    //   documentControllerId,
    //   [selectedProjectTask.InitiatedBy],
    //   `A document reworked by the Approver has been returned to you for the Deliverable <b>"${selectedProjectTask.Deliverable}"</b>. Please review the updated document and proceed with the necessary actions.`,
    //   `Reworked Document Returned – ${selectedProjectTask.Deliverable}`,
    //   projectDocumentInfo.map(d => d.id)
    // );
    //Satish added
  };

  const updateApprovalHierarchyStatusForRework = async (status: string) => {
    if (!selectedProjectTask) return;

    try {
      const filterQuery = `ProjectCreationListIDId eq ${selectedProjectTask.ProjectId} and DeliverablesDetailsIdId eq ${selectedProjectTask.DeliverableId} and Level eq '${selectedProjectTask.CurrentApprovalLevel}'`;

      const items = await sp.web.lists.getByTitle("ApprovalHierarchy")
        .items.filter(filterQuery)();

      if (items && items.length > 0) {
        for (const item of items) {
          await sp.web.lists.getByTitle("ApprovalHierarchy")
            .items.getById(item.Id)
            .update({ Status: status });
        }
      }
    } catch (error) {
      console.error("Error in updateApprovalHierarchyStatusForRework:", error);
    }
  };

  const createVendorApproval = async () => {
    if (!selectedProjectTask) return;

    const vendorId = 6 // Implement this

    const approvalData = {
      ProjectCreationListIDId: selectedProjectTask.ProjectId,
      DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
      AssignedToId: selectedProjectTask.VendorId,
      RequestedById: props.context.pageContext.legacyPageContext.userId,
      RequestedDate: new Date(),
      DocumentType: selectedProjectTask.DocType,
      ApproverRole: "Vendor",
      ApprovalCriteria: "Anyone",
      Status: 'pending',
      Level: "Level 0",
      SerialNumber: 0,
      RevisionNumber: selectedProjectTask.RevisionNumber,
    };
    await sp.web.lists.getByTitle("ProjectApprovals").items.add(approvalData);
    const actionType: ActionType = "REWORK_BY_DOCUMENT_CONTROLLER";  //rework by document controller

    // to get the approver name 
    var TouserUsers = Array.isArray(users)
      ? users.filter(u => String(u.value) === String(selectedProjectTask.VendorId))
      : [];
    var TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
    //  to get the initiator name
    var CCUserName = Array.isArray(users)
      ? users.filter(u => String(u.value) === String(selectedProjectTask.InitiatedBy))
      : [];
    var CCName = CCUserName.length > 0 ? CCUserName[0].label : null;
    // end 

    const emailBody = buildApprovalEmailBody(
      actionType,
      TouserName, // whoever is receiving
      CCName,   // whoever is sending
      projectDocumentInfo.map(d => ({
        deliverable: selectedProjectTask.Deliverable,
        fileName: d.fileName,
        sharedLink: d.sharedLink
      }))
    );
    await createEmailTriggerDetails(
      selectedProjectTask.VendorId,
      [selectedProjectTask.InitiatedBy],
      emailBody,
      `${actionConfig[actionType].subject} – ${selectedProjectTask.Deliverable}`,
      projectDocumentInfo.map(d => d.id)
    );
    // await createEmailTriggerDetails(
    //   selectedProjectTask.VendorId,
    //   [selectedProjectTask.InitiatedBy],
    //   `The submitted document has been reviewed by the Document Controller and requires rework for the Deliverable <b>"${selectedProjectTask.Deliverable}"</b>. Please review the feedback, make the necessary corrections, and resubmit the document at the earliest.`,
    //   `Rework Action Assigned – ${selectedProjectTask.Deliverable}`,
    //   projectDocumentInfo.map(d => d.id)
    // );
    //Satish added
  };


  // Handle Rejected action
  const handleRejectedAction = async () => {
    if (!selectedProjectTask) return;

    const approvals = await getAllProjectApproval(selectedProjectTask.ProjectId!, selectedProjectTask.DeliverableId!);

    // Update other pending approvals in current level to Rejected
    const pendingApprovals = approvals.filter(x =>
      x.Level === selectedProjectTask?.CurrentApprovalLevel &&
      x.ApproverRole === selectedProjectTask?.ApprovalRole &&
      x.Id !== selectedProjectTask.Id &&
      x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
      x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId
    );

    for (const item of pendingApprovals) {
      const data = {
        Status: "Rejected",
        Remarks: "Auto Reject",
        ApprovalDate: new Date(),
      };
      await sp.web.lists.getByTitle("ProjectApprovals").items.getById(item.Id).update(data);
    }

    await updateDeliverableStatus("Rejected", "");
    await updateApprovalHierarchyStatusForRework("Rejected");
  };


  // const updateDeliverablePublishStatus = async (deliverableId: number, publishStatus: string): Promise<void> => {
  //   try {
  //     await sp.web.lists.getByTitle("DeliverablesDetails")
  //       .items.getById(deliverableId)
  //       .update({
  //         IsPublished: publishStatus
  //       });

  //     console.log("Deliverable publish status updated successfully");
  //   } catch (error) {
  //     console.error("Error updating deliverable publish status:", error);
  //     throw error;
  //   }
  // };
  // Project Approval Code ends here ----------------

  const updateDeliverablePublishStatus = async (
    deliverableId: number,
    publishStatus: string,
    documentIds: number[]
  ): Promise<void> => {
    try {
      // 1. Update DeliverablesDetails list
      await sp.web.lists.getByTitle("DeliverablesDetails")
        .items.getById(deliverableId)
        .update({
          IsPublished: publishStatus
        });

      // 2. Update DeliverablesDocument library items
      const deliverablesDocList = sp.web.lists.getByTitle("DeliverablesDocument");
      for (const docId of documentIds) {
        await deliverablesDocList.items.getById(docId).update({
          IsPublish: publishStatus,
          CodeStatus: projectCodestatus
        });
      }

      console.log("Deliverable and related documents publish status updated successfully");
    } catch (error) {
      console.error("Error updating publish status:", error);
      throw error;
    }
  };

  React.useEffect(() => {
    sessionStorage.removeItem("announcementId");

    ApiCall("Pending");
  }, [useHide]);

  // const ApiCall = async (status: string) => {
  //   // if(activeTab == "Intranet"){
  //   setLoading(true);
  //   let MyApprovaldata = await getMyApproval(sp, status);
  //   let Automationdata1 = await getApprovalListsData(sp, status);
  //   let typedata = await getType(sp);
  //   //setMyApprovalsData(MyApprovaldata);
  //   setMyApprovalsDataAll(MyApprovaldata);
  //   //}
  //   //else if(activeTab == "Automation"){
  //   // let Automationdata = Automationdata1.sort((a, b) => {
  //   //   return a.Created === b.Created ? 0 : a.Created ? -1 : 1;
  //   // });
  //   let Automationdata = Automationdata1.sort((a, b) => {
  //     return b.Created - a.Created;
  //   });
  //   setMyApprovalsData(Automationdata);
  //   setMyApprovalsDataAutomation(Automationdata);
  //   setTimeout(() => {
  //     setLoading(false);;
  //   }, 15000);

  //   console.log("Automationdata", Automationdata);
  //   // }
  // };

  // const FilterDiscussionData = async (optionFilter: string) => {

  //   setAnnouncementData(await getDiscussionFilterAll(sp, optionFilter));

  // };

  // const ApiCall = async (status: string) => {
  //   // if(activeTab == "Intranet"){
  //   setLoading(true);
  //   let MyApprovaldata = await getMyApproval(sp, status);
  //   let Automationdata1 = await getApprovalListsData(sp, status);
  //   let typedata = await getType(sp);
  //   let ProjectWorkflowData = await getProjectWorkflowApprovals(status); // NEW
  //   let Automationdata: any;
  //   //setMyApprovalsData(MyApprovaldata);
  //   setMyApprovalsDataAll(MyApprovaldata);
  //   setProjectWorkflowdata(ProjectWorkflowData); // NEW
  //   //}
  //   //else if(activeTab == "Automation"){
  //   // let Automationdata = Automationdata1.sort((a, b) => {
  //   //   return a.Created === b.Created ? 0 : a.Created ? -1 : 1;
  //   // });
  //   if (Automationdata1.length > 0) {
  //     // Automationdata = Automationdata1.sort((a, b) => {
  //     //   return b.Created - a.Created;
  //     // });
  //     setMyApprovalsData(Automationdata1.sort((a, b) => b.Created - a.Created));
  //     setMyApprovalsDataAutomation(Automationdata1.sort((a, b) => b.Created - a.Created));
  //     //setLoading(false)
  //   } else {
  //     setMyApprovalsData([]);
  //     setMyApprovalsDataAutomation([]);
  //     setLoading(false)
  //   }
  //   setLoading(false)
  //   // setTimeout(() => {
  //   //   setLoading(false);;
  //   // }, 15000);

  //   console.log("Automationdata", Automationdata);
  //   // }
  // };

  // srs 10/4/26
const ApiCall = async (status: string) => {
    setLoading(true);
    let MyApprovaldata = await getMyApproval(sp, status);
    let Automationdata1 = await getApprovalListsData(sp, status);
    let ProjectWorkflowData = await getProjectWorkflowApprovals(status);

    setMyApprovalsDataAll(MyApprovaldata);
    setMyApprovalsDataAutomation(Automationdata1.sort((a, b) => b.Created - a.Created));
    setProjectWorkflowdata(ProjectWorkflowData);

    // If the tab is DMS, we don't want to overwrite with Automation data
    if (activeTab === "Intranet") {
        setMyApprovalsData(MyApprovaldata);
        setLoading(false);
    } else if (activeTab === "Automation") {
        setMyApprovalsData(Automationdata1.sort((a, b) => b.Created - a.Created));
        setLoading(false);
    } else if (activeTab === "ProjectWorkflow") {
        setMyApprovalsData(ProjectWorkflowData);
        setLoading(false);
    } else if (activeTab === "DMS") {
        // Data is being handled by getApprovalmasterTasklist
        // We just ensure the loader stays on until that finishes
    }
};

  //   // ================= EMAIL TRIGGER DETAILS =================
  // function buildApprovalEmailBody(
  //   approverName: string,
  //   senderName: string,
  //   deliverableName: string,
  //   projectDocumentInfo: { fileName: string; sharedLink: string; }[]
  // ): string {

  //   const rows = projectDocumentInfo.map(doc => `
  //     <tr>
  //       <td style="border:1px solid #ccc;padding:8px;">${deliverableName}</td>
  //       <td style="border:1px solid #ccc;padding:8px;">
  //         <a href="${doc.sharedLink}" target="_blank">${doc.fileName}</a>
  //       </td>
  //     </tr>
  //   `).join("");

  //   return `
  //     <div style="font-family:Segoe UI,Arial,sans-serif;font-size:14px;">
  //       Dear ${approverName},<br/><br/>

  //       An approval action has been assigned to you by the Document Controller for the deliverable(s) mentioned below. Please review and take the necessary action at the earliest.<br/><br/>

  //       <table style="border-collapse:collapse;width:100%;max-width:700px;">
  //         <thead>
  //           <tr>
  //             <th style="border:1px solid #ccc;padding:8px;background:#f4f4f4;">Deliverable</th>
  //             <th style="border:1px solid #ccc;padding:8px;background:#f4f4f4;">Document</th>
  //           </tr>
  //         </thead>
  //         <tbody>
  //           ${rows}
  //         </tbody>
  //       </table>

  //       <br/><br/>

  //       <a href="https://officeindia.sharepoint.com/sites/Intranetdemos/SitePages/MyApprovals.aspx" target="_blank">
  //         Click here to view your approvals
  //       </a>

  //       <br/><br/>

  //       Regards,<br/>
  //       ${senderName}
  //     </div>
  //   `;
  // }

  const waitForShareLink = async (fileItemId: number, maxAttempts = 20, delayMs = 3000) => {
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
      const item = await sp.web.lists
        .getByTitle("DeliverablesDocument")
        .items.getById(fileItemId)
        .select("File/Name", "SharedLink", "Revision")
        .expand("File")();

      if (item.SharedLink) {
        return {
          shareLink: item.SharedLink,
          fileName: item.File?.Name,
          revisionNo: item.Revision
        };
      }

      // wait before next attempt
      await new Promise(res => setTimeout(res, delayMs));
    }

    throw new Error(`ShareLink not generated for FileItemId: ${fileItemId}`);
  };
  // older one 
  // const buildApprovalEmailBody = (
  //   actionType: ActionType,
  //   approverName: string,
  //   senderName: string,
  //   docs: { deliverable: string; fileName: string; sharedLink: string; }[]
  // ) => {
  //   const { message } = actionConfig[actionType];

  //   const rows = docs.map(d => `
  //     <tr>
  //       <td style="border:1px solid #ccc;padding:8px;">${d.deliverable}</td>
  //       <td style="border:1px solid #ccc;padding:8px;">
  //         <a href="${d.sharedLink}" target="_blank">${d.fileName}</a>
  //       </td>
  //     </tr>
  //   `).join("");

  //   return `
  //     Dear ${approverName},<br/><br/>
  //     ${message}<br/><br/>

  //     <table style="border-collapse:collapse;width:100%;max-width:700px;">
  //       <thead>
  //         <tr>
  //           <th style="border:1px solid #ccc;padding:8px;background:#f4f4f4;">Deliverable</th>
  //           <th style="border:1px solid #ccc;padding:8px;background:#f4f4f4;">Document</th>
  //         </tr>
  //       </thead>
  //       <tbody>
  //         ${rows}
  //       </tbody>
  //     </table>

  //     <br/><br/>

  //     <a href="https://officeindia.sharepoint.com/sites/Intranetdemos/SitePages/MyApprovals.aspx" target="_blank">
  //       Click here to view in portal
  //     </a>

  //     <br/><br/>
  //     Regards,<br/>
  //     ${senderName}
  //   `;
  // };
  // new one
  const buildbulkApprovalEmailBody = (
    actionType: ActionType,
    documentNumber: string,
    approverName: string,
    senderName: string,
    projectName: string,
    docs: { deliverable: string; fileName: string; sharedLink: string; docrevisionNo: string }[]
  ) => {
    const { message } = actionConfig[actionType];
    // Format current date as dd MMM yyyy
    const today = new Date();
    const sentDate = today.toLocaleDateString("en-GB", {
      day: "numeric",
      month: "short",
      year: "numeric"
    });

    let rows = docs.map((d, i) => {
      const no = i + 1;
      const transmittal = `TR-${String(no).padStart(3, "0")}`;

      return `
        <tr>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${no}</td>
          <td style="border:1px solid #ccc;padding:6px;">
            <a href="${d.sharedLink}" target="_blank">${documentNumber}</a>
          </td>
          <td style="border:1px solid #ccc;padding:6px;">${d.deliverable}</td>

          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${d.docrevisionNo ?? "-"}</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${mdrStatus ?? "-"}</td>
          <td style="border:1px solid #ccc;padding:6px;">${transmittal}</td>
          <td style="border:1px solid #ccc;padding:6px;">${sentDate}</td>
          <td style="border:1px solid #ccc;padding:6px;">Sent for rework</td>
        </tr>
      `;
    }).join("");

    // --- Extra Dropbox Row ---
    const extraNo = docs.length + 1;
    const extraTransmittal = `TR-${String(extraNo).padStart(3, "0")}`;
    const dropboxUrl = `https://officeindia.sharepoint.com/sites/Intranetdemos/DeliverablesDoc/Forms/AllItems.aspx?id=%2Fsites%2FIntranetdemos%2FDeliverablesDoc%2FESSADeliverablesDoc&viewid=21fa2859%2Dc815%2D46d1%2D96bd%2Dbdabee445848&npsAction=createList`;

    rows += `
    <tr>
      <td style="border:1px solid #ccc;padding:6px;text-align:center;">${extraNo}</td>
      <td style="border:1px solid #ccc;padding:6px;"></td>
      <td style="border:1px solid #ccc;padding:6px;">${projectName}</td>
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

    <a href="https://officeindia.sharepoint.com/sites/Intranetdemos/SitePages/MyTask.aspx" target="_blank">
      Click here to view in portal
    </a>

    <br/><br/>
    Regards,<br/>
    ${senderName}
  `;
  };
  const buildApprovalEmailBody = (
    actionType: ActionType,
    approverName: string,
    senderName: string,
    docs: { deliverable: string; fileName: string; sharedLink: string; }[]
  ) => {
    const { message } = actionConfig[actionType];
    // Logic: Send to MyTask for Vendor rework, else MyApprovals for standard flows
    const redirectUrl = actionType === "REWORK_BY_DOCUMENT_CONTROLLER"
      ? "https://officeindia.sharepoint.com/sites/Intranetdemos/SitePages/MyTask.aspx"
      : "https://officeindia.sharepoint.com/sites/Intranetdemos/SitePages/MyApprovals.aspx";
    // Format current date as dd MMM yyyy
    const today = new Date();
    const sentDate = today.toLocaleDateString("en-GB", {
      day: "numeric",
      month: "short",
      year: "numeric"
    });

    let rows = docs.map((d, i) => {
      const no = i + 1;
      const transmittal = `TR-${String(no).padStart(3, "0")}`;

      // <td style="border:1px solid #ccc;padding:6px;">${selectedProjectTask?.DocNumber}</td>
      // <th style="border:1px solid #ccc;padding:6px;">Document Name</th>
      return `
        <tr>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${no}</td>
          <td style="border:1px solid #ccc;padding:6px;">
            <a href="${d.sharedLink}" target="_blank">${selectedProjectTask?.DocNumber}</a>
          </td>
          <td style="border:1px solid #ccc;padding:6px;">${d.deliverable}</td>

          <td style="border:1px solid #ccc;padding:6px;text-align:center;">${selectedProjectTask.RevisionNumber ?? "-"}</td>
          <td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td>
          <td style="border:1px solid #ccc;padding:6px;">${transmittal}</td>
          <td style="border:1px solid #ccc;padding:6px;">${sentDate}</td>
          <td style="border:1px solid #ccc;padding:6px;">Attached</td>
        </tr>
      `;
    }).join("");

    // --- Extra Dropbox Row ---
    const extraNo = docs.length + 1;
    const extraTransmittal = `TR-${String(extraNo).padStart(3, "0")}`;
    const dropboxUrl = `https://officeindia.sharepoint.com/sites/Intranetdemos/DeliverablesDoc/Forms/AllItems.aspx?id=%2Fsites%2FIntranetdemos%2FDeliverablesDoc%2FESSADeliverablesDoc&viewid=21fa2859%2Dc815%2D46d1%2D96bd%2Dbdabee445848&npsAction=createList`;

    rows += `
    <tr>
      <td style="border:1px solid #ccc;padding:6px;text-align:center;">${extraNo}</td>
      <td style="border:1px solid #ccc;padding:6px;"></td>
      <td style="border:1px solid #ccc;padding:6px;">${selectedProjectTask?.ProjectName}</td>
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

    <a href="${redirectUrl}" target="_blank">
      Click here to view in portal
    </a>

    <br/><br/>
    Regards,<br/>
    ${senderName}
  `;
  };


  const createEmailTriggerDetails = async (TouserId: number, CCuserIds: number[], message: string, SubjectMsg: string, documentIds: number[]) => {
    if (!selectedProjectTask) return;
    await sp.web.lists
      .getByTitle("EmailTriggerDetails")
      .items.add({
        Title: selectedProjectTask.ProjectName,
        Subject: SubjectMsg,
        Body: message,
        ToUserId: [TouserId],
        CCUserId: CCuserIds,
        // AttachmentIdId: documentIds
      });
  };
  const createbulkEmailTriggerDetails = async (TouserId: number, CCuserIds: number[], message: string, SubjectMsg: string, documentIds: number[], Deliverable: string) => {
    if (!selectedProjectItems) return;
    await sp.web.lists
      .getByTitle("EmailTriggerDetails")
      .items.add({
        Title: Deliverable,
        Subject: SubjectMsg,
        Body: message,
        ToUserId: [TouserId],
        CCUserId: CCuserIds,
        // AttachmentIdId: documentIds
      });
    // Reload after successful post
    // window.location.reload();
  };
  const handleStatusChange = async (name: string, value: string, actingfor: any) => {
    setLoading(true);
    setStatusChange(true);
    actingforuseremail = actingfor
    // alert(`Status value is ${value} is acting for ${actingfor}`)
    if (actingforuseremail === undefined || actingforuseremail === null || actingforuseremail === "") {
      // alert("acting for is undefined")
    }

    if (value === "") {
      // Show all records if no type is selected
      console.log("No status selected");
    } else {
      SetStatusvalue(name);
      // Filter records based on the selected type
      let MyApprovaldata = await getMyApproval(sp, value, actingfor);
      let Automationdata = await getApprovalListsData(sp, value, actingfor);
      // let MyDMSAPPROVALDATA:any = await MyDMSAPPROVALDATASTATUS(sp, value)
      let MyDMSAPPROVALDATA: any = await getApprovalmasterTasklist(value, actingfor);
      console.log("MyDMSAPPROVALDATA", MyDMSAPPROVALDATA)

      let ProjectWorkflowData = await getProjectWorkflowApprovals(value, actingfor); // NEW
      console.log("ProjectWorkflowData", ProjectWorkflowData);
      setProjectWorkflowdata(ProjectWorkflowData); // NEW
      setMyApprovalsDataAll(MyApprovaldata);
      setMyApprovalsDataAutomation(Automationdata);
      if (activeTab == "Intranet") {
        setMyApprovalsData(MyApprovaldata);
        if (MyApprovaldata.length > 0) {
          setTimeout(() => {
            setLoading(false);
          }, 5000);
        } else {
          setLoading(false)
        }
      } else if (activeTab == "DMS") {
        // alert(value)
        setMyApprovalsData(MyDMSAPPROVALDATA);
        if (MyDMSAPPROVALDATA.length > 0) {
          setTimeout(() => {
            setLoading(false);
          }, 3000);
        } else {
          setLoading(false)
        }
      } else if (activeTab == "Automation") {
        setMyApprovalsData(Automationdata.sort((a, b) => b.Created - a.Created));
        if (Automationdata.length > 0) {
          setTimeout(() => {
            setLoading(false);
          }, 3000);
        } else {
          setLoading(false)
        }
        console.log("Automationdata", Automationdata);
      } else if (activeTab == "ProjectWorkflow") { // NEW
        setMyApprovalsData(ProjectWorkflowData);
        if (ProjectWorkflowData.length > 0) {
          setTimeout(() => {
            setLoading(false);
          }, 3000);
        } else {
          setLoading(false)
        }
      }
      setStatusChange(false);
      // else if (activeTab == "Automation") {
      //   setMyApprovalsData(null);
      //   setMyApprovalsData(MyDMSAPPROVALDATA);
      //   setMyApprovalsData(Mylistdata);
      //   console.log("Automationdata", Automationdata);
      // }
    }

    setLoading(false);
  };
  const handleFilterChange = (
    e: React.ChangeEvent<HTMLInputElement>,

    field: string
  ) => {
    console.log("eee", e);

    setFilters({
      ...filters,

      [field]: e.target.value,
    });
  };

  // const applyFiltersAndSorting = (data: any[]) => {
  //   // Filter data

  //   console.log(
  //     "filters",

  //     data,

  //     filters,

  //     filters.ProcessName,

  //     filters.RequestID
  //   );

  //   const filteredData = data?.filter((item, index) => {
  //     return (
  //       (filters.SNo === "" || String(index + 1).includes(filters.SNo)) &&
  //       // (filters.Title === "" ||
  //       //   (activeTab == "Intranet" ? item.Title != undefined : item.ApprovalTitle != undefined) &&
  //       // (activeTab == "Intranet" ? item.Title : item.ApprovalTitle).toLowerCase().includes(filters.Title.toLowerCase()))
  //       (filters.ProcessName === "" ||
  //         (item.ProcessName != undefined &&
  //           item.ProcessName.toLowerCase().includes(
  //             filters.ProcessName.toLowerCase()
  //           ))) &&
  //       (filters.RequestID === "" ||
  //         (item.RequestID != undefined &&
  //           item.RequestID.toLowerCase().includes(
  //             filters.RequestID.toLowerCase()
  //           ))) &&
  //       // (filters.ProcessName === "" ||

  //       //   item.ProcessName.toLowerCase().includes(filters.ProcessName.toLowerCase())) &&

  //       (filters.Status === "" ||
  //         item.Status.toLowerCase().includes(filters.Status.toLowerCase())) &&
  //       (filters.RequestedDate === "" ||
  //         new Date(item.Created)
  //           .toLocaleDateString()
  //           .startsWith(filters.RequestedDate + "")) &&
  //       (filters.Title === "" ||
  //         (activeTab == "Automation"
  //           ? item?.ApprovalTitle?.toLowerCase().includes(
  //             filters.Title.toLowerCase()
  //           )
  //           : item?.Title?.toLowerCase().includes(
  //             filters.Title.toLowerCase()
  //           ))) &&
  //       (filters.RequestedBy === "" ||
  //         (activeTab == "Automation"
  //           ? item?.Author?.Title?.toLowerCase().includes(
  //             filters.RequestedBy.toLowerCase()
  //           )
  //           : item?.Requester?.Title?.toLowerCase().includes(
  //             filters.RequestedBy.toLowerCase()
  //           )))
  //     );
  //   });

  //   const sortedData = filteredData?.sort((a, b) => {
  //     if (sortConfig.key === "SNo") {
  //       // Sort by index

  //       const aIndex = data.indexOf(a);

  //       const bIndex = data.indexOf(b);

  //       return sortConfig.direction === "ascending"
  //         ? aIndex - bIndex
  //         : bIndex - aIndex;
  //     } else if (sortConfig.key == "RequestedDate") {
  //       // Sort by other keys

  //       const aValue = a["Created"] ? new Date(a["Created"]) : "";

  //       const bValue = b["Created"] ? new Date(b["Created"]) : "";

  //       if (aValue < bValue) {
  //         return sortConfig.direction === "ascending" ? -1 : 1;
  //       }

  //       if (aValue > bValue) {
  //         return sortConfig.direction === "ascending" ? 1 : -1;
  //       }
  //     } else if (sortConfig.key) {
  //       // Sort by other keys

  //       const aValue = a[sortConfig.key] ? a[sortConfig.key].toLowerCase() : "";

  //       const bValue = b[sortConfig.key] ? b[sortConfig.key].toLowerCase() : "";

  //       if (aValue < bValue) {
  //         return sortConfig.direction === "ascending" ? -1 : 1;
  //       }

  //       if (aValue > bValue) {
  //         return sortConfig.direction === "ascending" ? 1 : -1;
  //       }
  //     }

  //     return 0;
  //   });

  //   return sortedData;
  // };

// Aman Filter logic added 18/3/26
  const applyFiltersAndSorting = (data: any[]) => {
    const filteredData = data?.filter((item, index) => {
      const sNo = String(index + 1);
     
      // Data extraction based on Tab structure
      let requestId = "";
      let title = "";
      let process = "";
      let requestedBy = "";
      let requestedDate = "";
      let status = "";
 
      if (activeTab === "DMS") {
        requestId = item?.FileUID?.RequestNo || "";
        title = item?.FileUID?.FileName || "";
        process = item?.FileUID?.Processname || "";
        requestedBy = item?.RequestedByTitle || "";
        // aman 17/03/26 - Fetching date string exactly  "02/02/2026, 03:02 PM"
        requestedDate = item?.FileUID?.Created || "";
        status = item?.FileUID?.Status || "Pending";
      } else {
        requestId = item?.RequestID || "";
        title = activeTab === "Automation" ? (item?.ApprovalTitle || "") : (item?.Title || "");
        process = item?.ProcessName || "";
        requestedBy = activeTab === "Automation" ? (item?.Author?.Title || "") : (item?.Requester?.Title || "");
        requestedDate = item?.Created ? new Date(item.Created).toLocaleString() : "";
        status = item?.Status || "";
      }
 
   
      return (
        (filters.RequestID === "" || requestId.toLowerCase().includes(filters.RequestID.toLowerCase())) &&
        (filters.Title === "" || title.toLowerCase().includes(filters.Title.toLowerCase())) &&
        (filters.ProcessName === "" || process.toLowerCase().includes(filters.ProcessName.toLowerCase())) &&
        (filters.RequestedBy === "" || requestedBy.toLowerCase().includes(filters.RequestedBy.toLowerCase())) &&
        (filters.RequestedDate === "" || String(requestedDate).toLowerCase().includes(filters.RequestedDate.toLowerCase())) &&
        (filters.Status === "" || status.toLowerCase().includes(filters.Status.toLowerCase()))
      );
    });
 
    const sortedData = filteredData?.sort((a, b) => {
      if (sortConfig.key === "SNo") {
        const aIndex = data.indexOf(a);
        const bIndex = data.indexOf(b);
        return sortConfig.direction === "ascending" ? aIndex - bIndex : bIndex - aIndex;
      } else if (sortConfig.key) {
        const aValue = String(a[sortConfig.key] || "").toLowerCase();
        const bValue = String(b[sortConfig.key] || "").toLowerCase();
        if (aValue < bValue) return sortConfig.direction === "ascending" ? -1 : 1;
        if (aValue > bValue) return sortConfig.direction === "ascending" ? 1 : -1;
      }
      return 0;
    });
 
    return sortedData;
  };
  const filteredMyApprovalData = applyFiltersAndSorting(myApprovalsData);

  const filteredNewsData = applyFiltersAndSorting(newsData);

  const [currentPage, setCurrentPage] = React.useState(1);
  const [currentGroup, setCurrentGroup] = React.useState(1);
  const itemsPerPage = 10;
  const pagesPerGroup = 10;
  const totalPages = Math.ceil(filteredMyApprovalData?.length / itemsPerPage);
  const totalGroups = Math.ceil(totalPages / pagesPerGroup);
  const [ContentData, setContentData] = React.useState<any>([]);

  const [currentItem, setCurrentItem] = React.useState<any>([]);

  // const handlePageChange = (pageNumber: any) => {
  //   if (pageNumber > 0 && pageNumber <= totalPages) {
  //     setCurrentPage(pageNumber);
  //   }
  // };
  const handleGroupChange = (direction: "next" | "prev") => {
    const newGroup = currentGroup + (direction === "next" ? 1 : -1);
    if (newGroup > 0 && newGroup <= totalGroups) {
      setCurrentGroup(newGroup);
      setCurrentPage((newGroup - 1) * pagesPerGroup + 1); // Go to the first page of the new group
    }
  };

  const handlePageChange = (pageNumber: any) => {

    if (pageNumber > 0 && pageNumber <= totalPages) {

      setCurrentPage(pageNumber);
      const newGroup = Math.ceil(pageNumber / pagesPerGroup);
      setCurrentGroup(newGroup);
    }

  };
  const startIndex = (currentPage - 1) * itemsPerPage;

  const endIndex = startIndex + itemsPerPage;

  const currentData = filteredMyApprovalData?.slice(startIndex, endIndex);
  const startPage = (currentGroup - 1) * pagesPerGroup + 1;
  const endPage = Math.min(currentGroup * pagesPerGroup, totalPages);
  const newsCurrentData = filteredNewsData?.slice(startIndex, endIndex);

  const [editID, setEditID] = React.useState(null);

  const [ImagepostIdsArr, setImagepostIdsArr] = React.useState([]);

  const siteUrl = props.siteUrl;

  const Breadcrumb = [
    {
      MainComponent: "Home",

      MainComponentURl: `${siteUrl}/SitePages/Dashboard.aspx`,
    },

    {
      ChildComponent: "My Approvals",

      ChildComponentURl: `${siteUrl}/SitePages/MyApprovals.aspx`,
    },
  ];

  console.log(announcementData, "announcementData");

  const exportToExcel = (data: any[], fileName: string) => {
    const workbook = XLSX.utils.book_new();

    const worksheet = XLSX.utils.json_to_sheet(data);

    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    XLSX.writeFile(workbook, `${fileName}.xlsx`);
  };
  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(1);
      setCurrentGroup(1);
    }
  }, [filteredMyApprovalData, currentGroup]);
  // const fetchOptions = async () => {

  //   try {

  //     const items = await fetchUserInformationList(sp);

  //     console.log(items, "itemsitemsitems");

  //     const formattedOptions = items.map((item: { Title: any; Id: any }) => ({

  //       name: item.Title, // Adjust according to your list schema

  //       id: item.Id,

  //     }));

  //     setOpions(formattedOptions);

  //   } catch (error) {

  //     console.error("Error fetching options:", error);

  //   }

  // };

  const handleSortChange = (key: string) => {
    let direction = "ascending";

    if (sortConfig.key === key && sortConfig.direction === "ascending") {
      direction = "descending";
    }

    setSortConfig({ key, direction });
  };

  React.useEffect(() => {
    const showNavbar = (
      toggleId: string,

      navId: string,

      bodyId: string,

      headerId: string
    ) => {
      const toggle = document.getElementById(toggleId);

      const nav = document.getElementById(navId);

      const bodypd = document.getElementById(bodyId);

      const headerpd = document.getElementById(headerId);

      UserGet();

      if (toggle && nav && bodypd && headerpd) {
        toggle.addEventListener("click", () => {
          nav.classList.toggle("show");

          toggle.classList.toggle("bx-x");

          bodypd.classList.toggle("body-pd");

          headerpd.classList.toggle("body-pd");
        });
      }
    };

    showNavbar("header-toggle", "nav-bar", "body-pd", "header");

    const linkColor = document.querySelectorAll(".nav_link");

    function colorLink(this: HTMLElement) {
      if (linkColor) {
        linkColor.forEach((l) => l.classList.remove("active"));

        this.classList.add("active");
      }
    }

    linkColor.forEach((l) => l.addEventListener("click", colorLink));
  }, [useHide]);

  const UserGet = async () => {
    const users = await sp.web.siteUsers();

    console.log(users, "users");
  };

  const handleRedirect = async (
    e: React.MouseEvent<SVGElement, MouseEvent>,
    Item: any, mode: any
  ) => {
    e.preventDefault();

    let arr = [];

    setCurrentItem(Item);



    //setisActivedata(true);

    let sessionkey = "";
    let redirecturl = "";
    if (activeTab == "Automation") {
      window.open(Item.RedirectionLink, "_blank");
      //window.location.href = `${Item.RedirectionLink}`;
    } else if (activeTab == "Intranet") {
      setContentData(await getDataByID(sp, Item?.ContentId, Item?.ContentName));

      if (Item?.ProcessName !== "Blog") {
        setisActivedata(true);
      }
      if (Item?.ProcessName) {
        switch (Item?.ProcessName) {
          case "Announcement":
            sessionkey = "announcementId";
            redirecturl =
              `${siteUrl}/SitePages/AddAnnouncement.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" + mode + "&page=MyApproval";
            break;
          case "News":
            sessionkey = "announcementId";
            redirecturl =
              `${siteUrl}/SitePages/AddAnnouncement.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" + mode + "&page=MyApproval";
            break;
          case "Event":
            sessionkey = "EventId";
            redirecturl =
              `${siteUrl}/SitePages/EventMasterForm.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" + mode + "&page=MyApproval";
            break;
          case "Media":
            sessionkey = "mediaId";
            redirecturl =
              `${siteUrl}/SitePages/MediaGalleryForm.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" + mode + "&page=MyApproval";
            break;
          case "Blog":
            sessionkey = "blogId";
            redirecturl =
              `${siteUrl}/SitePages/BlogDetails.aspx?` + Item?.ContentId + "&page=MyApproval";
            break;
          default:
        }

        const encryptedId = encryptId(String(Item?.ContentId));
        sessionStorage.setItem(sessionkey, encryptedId);
        location.href = redirecturl;
      }
    }

    // const encryptedId = encryptId(String(Item?.ContentId));

    // sessionStorage.setItem("announcementId", encryptedId);
  };

  const handleFromSubmit = async (
    e: React.MouseEvent<HTMLButtonElement, MouseEvent>,
    Status: string
  ) => {
    e.preventDefault();

    const postPayload = {
      Remark: formData.Remark,

      Status: Status,
      TriggerUpdateFlow: true,
    };

    console.log(postPayload);

    const postResult = await updateItemApproval(
      postPayload,
      sp,
      currentItem.Id
    );

    if (postResult) {
      setTimeout(() => {
        window.location.reload();
      }, 1000);
    }
  };

  React.useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(1);
    }
  }, [filteredMyApprovalData]);
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
              src={loaderGif}
              alt="Please wait..."
              style={{ width: "90px", height: "90px" }}
            />
            <p className="text-white mt-3 fw-semibold">
              please wait...
            </p>
          </div>
        </div>
      )}
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
            <div className="row" style={{ paddingLeft: "0.5rem" }}>
              <div className="col-md-4">

                <CustomBreadcrumb Breadcrumb={Breadcrumb} _context={sp} />

              </div>

              <div className="col-md-8">
                <div className="row">
                  <div style={{ textAlign: "center" }} className="col-md-3 newtexleft">
                    <div className="mb-0">
                      <label htmlFor="Status" className="form-label mt-2">
                  {/* <img src={require('../assets/delegation.png')} className='me-1' alt="d" />   */}
                           Acting on behalf of
                      </label>
                    </div>
                  </div>
                  <div style={{ paddingLeft: '0px' }} className="col-md-4">
                    <select
                      id="Type"
                      name="Type"
                      onChange={(e) => handleStatusChange(e.target.name, 'Pending', e.target.value)}
                      className="form-select"
                      disabled={loading || actingForUser.length === 0}
                    >
                      <option value="">
                        {loading ? "Loading..." : actingForUser.length === 0 ? "No Delegation" : "Select an option"}
                      </option>
                      {actingForUser.map((item, index) => (
                        <option key={index} value={item.email}>
                          {item.name}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div style={{ textAlign: "center", padding: '0px' }} className="col-md-1 newtexleft ivon">
                    <div className="mb-0">
                      <label htmlFor="Status" className="form-label newfil mt-2 mb-0">
                        Status
                        {/* <svg fill="#3c3c3c" width="23px" height="36px" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg" stroke="#b3b3b3"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path d="M12,25l6.67,6.67a1,1,0,0,0,.7.29.91.91,0,0,0,.39-.08,1,1,0,0,0,.61-.92V13.08L31.71,1.71A1,1,0,0,0,31.92.62,1,1,0,0,0,31,0H1A1,1,0,0,0,.08.62,1,1,0,0,0,.29,1.71L11.67,13.08V24.33A1,1,0,0,0,12,25ZM3.41,2H28.59l-10,10a1,1,0,0,0-.3.71V28.59l-4.66-4.67V12.67a1,1,0,0,0-.3-.71Z"></path> </g></svg> */}
                      </label>
                    </div>
                  </div>
                  <div style={{ paddingLeft: '0px' }} className="col-md-4 ivon2">
                    <select
                      id="Type"
                      name="Type"
                      onChange={(e) =>
                        handleStatusChange(e.target.name, e.target.value, actingforuseremail)
                      }
                      className="form-select"
                    >
                      {/* <option value="">Pending</option> */}
                      {StatusTypeData.map((item, index) => (
                        <option key={index} value={item.name}>
                          {item.name}
                        </option>
                      ))}
                    </select>
                  </div>

                </div>


              </div>
            </div>

            <div className="row mt-4">
              <div className="col-12">
                <div className="card mb-0 cardcsss">
                  <div className="card-body">
                    <div className="d-flex flex-wrap align-items-center justify-content-center">
                      <ul
                        className="nav nav-pills navtab-bg float-end justify-content-center"
                        role="tablist"
                      >
                        <li className="nav-item myapprovalcomingsoon" role="presentation">
                          {/* <li className="nav-item" role="presentation"> */}
                          <a
                            //onClick={() => handleTabClick("Intranet")}
                            className={`nav-link ${activeTab === "Intranet" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "Intranet"}
                            role="tab"
                          >
                            <span className="lenbg1">Intranet</span>{" "}
                            <span className="lenbg comingsoone">
                              {" "}
                              {/* {myApprovalsDataAll.length}  */}
                              coming soon
                            </span>
                          </a>
                          {/* <a
                            onClick={() => handleTabClick("Intranet")}
                            className={`nav-link ${activeTab === "Intranet" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "Intranet"}
                            role="tab"
                          >
                            <span className="lenbg1">Intranet</span>{" "}
                            <span className="lenbg">
                              {" "}
                              {myApprovalsDataAll.length}
                            </span>
                          </a> */}
                        </li>

                        <li className="nav-item" role="presentation">
                          <a
                            onClick={() => handleTabClick("DMS")}
                            className={`nav-link ${activeTab === "DMS" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "DMS"}
                            role="tab"
                            tabIndex={-1}
                          >
                            <span className="lenbg1">DMS </span>
                            <span className="lenbg">
                              {Mylistdata.length}
                            </span>
                          </a>
                        </li>

                        <li className="nav-item" role="presentation">
                          <a
                            onClick={() => handleTabClick("Automation")}
                            className={`nav-link ${activeTab === "Automation" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "Automation"}
                            role="tab"
                            tabIndex={-1}
                          >
                            <span className="lenbg1">Automation </span>
                            <span className="lenbg">
                              {myApprovalsDataAutomation.length}
                            </span>
                          </a>
                        </li>

                        <li className="nav-item" role="presentation">
                          <a
                            onClick={() => handleTabClick("ProjectWorkflow")}
                            className={`nav-link ${activeTab === "ProjectWorkflow" ? "active" : ""
                              }`}
                            aria-selected={activeTab === "ProjectWorkflow"}
                            role="tab"
                            tabIndex={-1}
                          >
                            <span className="lenbg1">Project Workflow </span>
                            <span className="lenbg">
                              {ProjectWorkflowdata.length}
                            </span>
                          </a>
                        </li>

                        {/* <li className="nav-item" role="presentation">
                          {selectedProjectItems.length > 0 && (
                            <button
                              type="button"
                              className="btn btn-danger ms-2"
                              onClick={handleReworkAction}
                            >
                              Rework
                            </button>
                          )}
                        </li> */}

                        {currentUserId === 12 && (
                          <>

                            <li className="nav-item" role="presentation"> <select className="form-select me-2" style={{ display: "inline-block" }} value={mdrStatus} onChange={(e) => setMdrStatus(e.target.value)} > <option value="">Select MDR</option> <option value="A">A</option> <option value="B">B</option> <option value="C">C</option> </select>

                            </li>
                            <li className="nav-item" role="presentation"> <button type="button" className="btn btn-danger ms-2" onClick={handlebulkReworkAction} > Rework </button>
                            </li>
                          </>

                        )}

                      </ul>
                    </div>
                  </div>
                </div>
                <div></div>
              </div>
            </div>
            {(activeTab === "Intranet" ||
              activeTab === "Automation" ||
              activeTab === "DMS" ||
              activeTab === "ProjectWorkflow") && (
                <div>
                  {!isActivedata && (
                    <div className=" mt-2">
                      <div className="">
                        <div style={{width:'100%'}} id="cardCollpase4" className="collapse show">

                          <div className="table-responsive pt-0">
                            {activeTab === "Intranet" ||
                              activeTab === "Automation" ? (

                              <>
                               <div className="card card-body">
                                <table 
                                  className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0"
                                  style={{ position: "relative",width:'100%' }}
                                >
                                  <thead>
                                    <tr>
                                      <th
                                        style={{
                                          borderBottomLeftRadius: "0px",

                                          minWidth: "40px",

                                          maxWidth: "40px",

                                          borderTopLeftRadius: "0px",
                                        }}
                                      >
                                        <div
                                          className="d-flex pb-2"
                                          style={{ justifyContent: "space-evenly" }}
                                        >
                                          <span>S.No.</span>

                                          <span
                                            onClick={() => handleSortChange("SNo")}
                                          >
                                            <FontAwesomeIcon icon={faSort} />
                                          </span>
                                        </div>

                                        <div className="bd-highlight">
                                          <input
                                            type="text"
                                            placeholder="index"
                                            onChange={(e) =>
                                              handleFilterChange(e, "SNo")
                                            }
                                            onKeyDown={(e) => {
                                              if (e.key === 'Enter' && !e.shiftKey) {
                                                e.preventDefault(); // Prevents the new line in textarea
                                              }
                                            }}
                                            className="inputcss"
                                            style={{ width: "100%" }}
                                          />
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "80px",
                                          maxWidth: "80px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Request ID</span>

                                            <span
                                              onClick={() =>
                                                handleSortChange("RequestID")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Request ID"
                                              onChange={(e) =>
                                                handleFilterChange(e, "RequestID")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>
                                      {/* {activeTab == "Intranet" && ( */}
                                      <th
                                        style={{
                                          minWidth: "120px",
                                          maxWidth: "120px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Title</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("Title")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />{" "}
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Title"
                                              onChange={(e) =>
                                                handleFilterChange(e, "Title")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>
                                      {/* )} */}
                                      <th
                                        style={{
                                          minWidth: "120px",
                                          maxWidth: "120px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Process Name</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("ProcessName")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />{" "}
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Process Name"
                                              onChange={(e) =>
                                                handleFilterChange(e, "ProcessName")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "100px",
                                          maxWidth: "100px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Requested By</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("RequestedBy")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />{" "}
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested By"
                                              onChange={(e) =>
                                                handleFilterChange(e, "RequestedBy")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "130px",
                                          maxWidth: "130px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Requested Date</span>{" "}
                                            {/* <span
                                            onClick={() =>
                                              handleSortChange("RequestedDate")
                                            }
                                          >
                                            <FontAwesomeIcon icon={faSort} />{" "}
                                          </span> */}
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested Date"
                                              onChange={(e) =>
                                                handleFilterChange(
                                                  e,
                                                  "RequestedDate"
                                                )
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "80px",
                                          maxWidth: "80px",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Status</span>{" "}
                                            {/* <span
                                            onClick={() =>
                                              handleSortChange("Status")
                                            }
                                          >
                                            <FontAwesomeIcon icon={faSort} />{" "}
                                          </span> */}
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Status"
                                              onChange={(e) =>
                                                handleFilterChange(e, "Status")
                                              }
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault(); // Prevents the new line in textarea
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th
                                        style={{
                                          minWidth: "50px",

                                          maxWidth: "50px",

                                          borderBottomRightRadius: "0px",

                                          borderTopRightRadius: "0px",

                                          textAlign: "center",

                                          verticalAlign: "top",
                                        }}
                                      >
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div
                                            className="d-flex  pb-2"
                                            style={{ justifyContent: "space-evenly" }}
                                          >
                                            <span>Action</span>{" "}
                                          </div>
                                        </div>
                                      </th>
                                    </tr>
                                  </thead>
                                  {console.log(
                                    "currentData",
                                    currentData,
                                    isActivedata
                                  )}
                                  {/* <tbody>
                                    {((loading && currentData?.length == 0)
                                      ||
                                      (StatusChange)) && (
                                        <div className="loadernewadd">
                                          <div>
                                            <img
                                              src={require("../../../CustomAsset/birdloader.gif")}
                                              className="alignrightl"
                                              alt="Loading..."
                                            />
                                          </div>
                                          <div className="loadnewarg">
                                            <span>Loading </span>{" "}
                                            <span>
                                              <img
                                                src={require("../../myApproval/assets/argloader.gif")}
                                                className="alignrightbird"
                                                alt="Loading..."
                                              />
                                            </span>
                                          </div>
                                        </div>
                                      )} */}
{/* Ritik 19/3/26 - Added loader for my approval  */}
<tbody>
  {((loading && currentData?.length == 0)
    ||
    (StatusChange)) && (
      <tr>
        <td colSpan={8} style={{ textAlign: "center", padding: "100px 0", border: "none", height:'450px' }}>
          <div className="loadernewadd" style={{ position: "static", display: "flex", flexDirection: "column", alignItems: "center" }}>
            <div>
              <img
                src={require("../../../CustomAsset/birdloader.gif")}
                className="alignrightl"
                alt="Loading..."
              />
            </div>
            <div className="loadnewarg">
              <span>Loading </span>{" "}
              <span>
                <img
                  src={require("../../myApproval/assets/argloader.gif")}
                  className="alignrightbird"
                  alt="Loading..."
                />
              </span>
            </div>
          </div>
        </td>
      </tr>
    )}
                                    {!loading && currentData?.length === 0 ? (
<tr>
  <td colSpan={8} style={{ border: "none" }}>
    <div className="no-results card card-body align-items-center annusvg text-center" style={{
      display: "flex",
      justifyContent: "center",
      position: 'relative',
      marginTop: '10px',
      height: '500px'
    }}>
      <svg style={{ top: '0%' }} xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>
      <p className="font-14 text-muted text-center">No Approval found </p>
    </div>
  </td>
</tr>
// end of no approval found
                                    ) : (
                                      !StatusChange && currentData?.map(
                                        (item: any, index: number) => (
                                          <tr
                                            key={index}

                                          >
                                            <td
                                              style={{
                                                minWidth: "40px",
                                                maxWidth: "40px",
                                              }}
                                            >
                                       <div className="d-flex align-items-center justify-content-center">       <div
                                                style={{ marginLeft: "0px" }}
                                                className="indexdesign"
                                              >
                                                {" "}
                                                {startIndex + index + 1}
                                          </div>    </div>{" "}
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "80px",

                                                maxWidth: "80px",
                                                textAlign: 'center',

                                                textTransform: "capitalize",
                                              }}
                                              title={item.RequestID}
                                            >
                                              {item.RequestID}
                                            </td>
                                            {/* {activeTab == "Intranet" && ( */}
                                            <td
                                              style={{
                                                minWidth: "120px",
                                                maxWidth: "120px",
                                              }}
                                              title={activeTab == "Intranet" ? item.Title : item.ApprovalTitle}
                                            >
                                              {activeTab == "Intranet" ? item.Title : item.ApprovalTitle}
                                            </td>
                                            {/* )} */}
                                            <td
                                              style={{
                                                minWidth: "120px",
                                                maxWidth: "120px",
                                                textAlign: 'center'
                                              }}
                                            >
                                              <span className="badge font-12 bg-secondary">   {item.ProcessName}</span>
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "100px",
                                                maxWidth: "100px",
                                              }}
                                              title={activeTab == "Automation"
                                                ? item?.Author?.Title
                                                : item?.Requester?.Title}
                                            >
                                              {activeTab == "Automation"
                                                ? item?.Author?.Title
                                                : item?.Requester?.Title}
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "130px",
                                                maxWidth: "130px",
                                                textAlign: 'center'
                                              }}
                                            >
                                              <div className="btn btn-light1 mt-1" style={{width:'79%'}}>
                                                {/* {new Date(
                                                  item?.Created
                                                ).toLocaleDateString()} */}
                                                {new Date(item?.Created).toLocaleString('en-US', {
                                                  month: '2-digit',
                                                  day: '2-digit',
                                                  year: 'numeric',
                                                  hour: '2-digit',
                                                  minute: '2-digit',
                                                  //second: '2-digit',
                                                  hour12: true
                                                })}
                                              </div>
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "80px",
                                                maxWidth: "80px",
                                                textAlign: 'center'
                                              }}
                                            >
                                              <div className="btn btn-status">
                                                {item?.Status}
                                              </div>
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "50px",
                                                maxWidth: "50px",
                                                textAlign: 'center'
                                              }}
                                              className="fe-eye font-18"
                                            >
                                              {item?.Status.toLowerCase() == "approved" || item?.Status.toLowerCase() == "rejected"
                                                || item?.Status.toLowerCase() == "completed" ?


                                                <Eye onClick={(e) =>
                                                  handleRedirect(e, item, "view")
                                                }

                                                  style={{

                                                    minWidth: "20px",

                                                    maxWidth: "20px",



                                                    cursor: "pointer",

                                                  }} />

                                                :
                                                <Edit
                                                  onClick={(e) =>
                                                    handleRedirect(e, item, "approval")
                                                  }

                                                  style={{
                                                    marginLeft: "0px",

                                                    cursor: "pointer",
                                                  }}
                                                />
                                              }
                                            </td>
                                          </tr>
                                        )
                                      )
                                    )}
                                  </tbody>
                                </table>

                                {currentData?.length > 0 ? (
                                  <nav className="pagination-container">
                                    <ul className="pagination">
                                      {/* <li
        className={`page-item ${currentPage === 1 ? "disabled" : ""
          }`}
      > */}
                                      <li

                                        className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""
                                          }`}
                                        onClick={() => handleGroupChange("prev")}
                                      >

                                        <a
                                          className="page-link"
                                          // onClick={() =>
                                          //   handlePageChange(currentPage - 1)
                                          // }
                                          aria-label="Previous"
                                        >
                                          «
                                        </a>
                                      </li>
                                      {Array.from(

                                        { length: endPage - startPage + 1 },

                                        (_, num) => {
                                          const pageNum = startPage + num;
                                          return (

                                            <li

                                              key={pageNum}

                                              className={`page-item ${currentPage === pageNum ? "active" : ""

                                                }`}

                                            >

                                              <a

                                                className="page-link"

                                                onClick={() =>

                                                  handlePageChange(pageNum)

                                                }

                                              >

                                                {pageNum}

                                              </a>

                                            </li>

                                          )
                                        }

                                      )}

                                      <li

                                        className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""

                                          }`}
                                        onClick={() => handleGroupChange("next")}
                                      >

                                        <a

                                          className="page-link"

                                          onClick={() =>

                                            handlePageChange(currentPage + 1)

                                          }

                                          aria-label="Next"

                                        >

                                          »

                                        </a>

                                      </li>

                                    </ul>
                                  </nav>
                                ) : (
                                  <></>
                                )} </div>
                              </>
                            ) : null}

                            {/* // Aman added filter logic 18/3/26 */}

                           {activeTab === "DMS" && (
                              <div>
                                {!showNestedDMSTable ? (
                                  <div className="card card-body">
                                    <table
                                      className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0"
                                   style={{ position: "relative",width:'100%' }}
                                >
                                  <thead>
                                    <tr>
                                          {/* aman 17/03/26 - Added search inputs for all DMS fields */}
                                         <th style={{ borderBottomLeftRadius: "0px", minWidth: "40px", maxWidth: "40px", borderTopLeftRadius: "0px" }}>
                                            <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                              <span>S.No.</span>
                                              <span onClick={() => handleSortChange("SNo")}><FontAwesomeIcon icon={faSort} /></span>
                                            </div>
                                            <div style={{height:'34px'}}></div>
                                           
                                          </th>
 
                                          <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                            <div className="d-flex flex-column bd-highlight">
                                              <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                <span>Request ID</span>
                                                <span onClick={() => handleSortChange("RequestID")}><FontAwesomeIcon icon={faSort} /></span>
                                              </div>
                                              <div className="bd-highlight">
                                                <input type="text" className="inputcss" placeholder="Filter" value={filters.RequestID} onChange={(e) => handleFilterChange(e, "RequestID")} style={{ width: "100%" }} />
                                              </div>
                                            </div>
                                          </th>
 
                                          <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                            <div className="d-flex flex-column bd-highlight">
                                              <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                <span>Title</span>
                                                <span onClick={() => handleSortChange("Title")}><FontAwesomeIcon icon={faSort} /></span>
                                              </div>
                                              <div className="bd-highlight">
                                                <input type="text" className="inputcss" placeholder="Filter" value={filters.Title} onChange={(e) => handleFilterChange(e, "Title")} style={{ width: "100%" }} />
                                              </div>
                                            </div>
                                          </th>
 
                                          <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                            <div className="d-flex flex-column bd-highlight">
                                              <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                <span>Process Name</span>
                                                <span onClick={() => handleSortChange("ProcessName")}><FontAwesomeIcon icon={faSort} /></span>
                                              </div>
                                              <div className="bd-highlight">
                                                <input type="text" className="inputcss" placeholder="Filter" value={filters.ProcessName} onChange={(e) => handleFilterChange(e, "ProcessName")} style={{ width: "100%" }} />
                                              </div>
                                            </div>
                                          </th>
 
                                          <th style={{ minWidth: "100px", maxWidth: "100px" }}>
                                            <div className="d-flex flex-column bd-highlight">
                                              <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                <span>Requested By</span>
                                                <span onClick={() => handleSortChange("RequestedBy")}><FontAwesomeIcon icon={faSort} /></span>
                                              </div>
                                              <div className="bd-highlight">
                                                <input type="text" className="inputcss" placeholder="Filter" value={filters.RequestedBy} onChange={(e) => handleFilterChange(e, "RequestedBy")} style={{ width: "100%" }} />
                                              </div>
                                            </div>
                                          </th>
 
                                          <th style={{ minWidth: "130px", maxWidth: "130px" }}>
                                            <div className="d-flex flex-column bd-highlight">
                                              <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                <span>Requested Date</span>
                                              </div>
                                              <div className="bd-highlight">
                                                <input type="text" className="inputcss" placeholder="Filter" value={filters.RequestedDate} onChange={(e) => handleFilterChange(e, "RequestedDate")} style={{ width: "100%" }} />
                                              </div>
                                            </div>
                                          </th>
 
                                          <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                            <div className="d-flex flex-column bd-highlight">
                                              <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                <span>Status</span>
                                              </div>
                                              <div className="bd-highlight">
                                                <input type="text" className="inputcss" placeholder="Filter" value={filters.Status} onChange={(e) => handleFilterChange(e, "Status")} style={{ width: "100%" }} />
                                              </div>
                                            </div>
                                          </th>
 
                                          <th style={{ minWidth: "50px", maxWidth: "50px", textAlign: "center", verticalAlign: "top" }}>
                                            <div className="d-flex flex-column bd-highlight">
                                              <div className="d-flex pb-2" style={{ justifyContent: "space-between" }}>
                                                <span>Action</span>
                                              </div>
                                            </div>
                                          </th>
                                        </tr>
                                      </thead>
                                      {console.log(
                                        "currentData",
                                        currentData,
                                        isActivedata
                                      )}
                                      {/* <tbody>
                                        {((loading && currentData?.length == 0) ||
                                          (StatusChange)) && (
                                            <div className="loadernewadd">
                                              <div>
                                                <img
                                                  src={require("../../../CustomAsset/birdloader.gif")}
                                                  className="alignrightl"
                                                  alt="Loading..."
                                                />
                                              </div>
                                              <div className="loadnewarg">
                                                <span>Loading </span>{" "}
                                                <span>
                                                  <img
                                                    src={require("../../myApproval/assets/argloader.gif")}
                                                    className="alignrightbird"
                                                    alt="Loading..."
                                                  />
                                                </span>
                                              </div>
                                            </div>
                                          )} */}
{/* Ritik 19/3/26 - Added loader in middle of the table for my approval DMS */}
<tbody>
  {((loading && currentData?.length == 0) ||
    (StatusChange)) && (
      <tr>
        <td colSpan={8} style={{ textAlign: "center", padding: "100px 0", border: "none" , height:'450px'}}>
          <div className="loadernewadd" style={{ position: "static", display: "flex", flexDirection: "column", alignItems: "center" }}>
            <div>
              <img
                src={require("../../../CustomAsset/birdloader.gif")}
                className="alignrightl"
                alt="Loading..."
              />
            </div>
            <div className="loadnewarg">
              <span>Loading </span>{" "}
              <span>
                <img
                  src={require("../../myApproval/assets/argloader.gif")}
                  className="alignrightbird"
                  alt="Loading..."
                />
              </span>
            </div>
          </div>
        </td>
      </tr>
    )}
                                        {!loading && currentData?.length === 0 ? (
<tr>
  <td colSpan={8} style={{ border: "none" }}>
    <div className="no-results card card-body align-items-center annusvg text-center" style={{
      display: "flex",
      justifyContent: "center",
      position: 'relative',
      marginTop: '10px',
      height: '500px'
    }}>
      <svg style={{ top: '0%' }} xmlns="http://www.w3.org/2000/svg" width="28" height="28" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>
      <p className="font-14 text-muted text-center">No Approval found </p>
    </div>
  </td>
</tr>
// end of no approval found for DMS
                                        ) : (
                                          !StatusChange && currentData?.map(
                                            (item: any, index: number) => (
                                              <tr
                                                key={index}

                                              >
                                                <td
                                                  style={{
                                                    minWidth: "40px",
                                                    maxWidth: "40px",

                                                    

                                                  }}
                                                >
                                                <div className="d-flex align-items-center justify-content-center">   <div
                                                    style={{ marginLeft: "0px" }}
                                                    className="indexdesign"
                                                  >
                                                    {" "}
                                                    {startIndex + index + 1}
                                                </div>  </div>{" "}
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "80px",

                                                    maxWidth: "80px",
                                                    textAlign: 'center',

                                                    textTransform: "capitalize",
                                                  }}
                                                  title={item?.FileUID?.RequestNo}
                                                >
                                                  {item?.FileUID?.RequestNo}

                                                </td>
                                                <td
                                                  style={{
                                                    minWidth: "120px",

                                                    maxWidth: "120px",


                                                  }}
                                                  title={item?.FileUID?.FileName}
                                                >
                                                  {item?.FileUID?.FileName}
                                                </td>
                                                <td
                                                  style={{
                                                    minWidth: "120px",
                                                    maxWidth: "120px",
                                                    textAlign: 'center'
                                                  }}
                                                >
                                                  <span className="badge font-12 bg-secondary">  {item?.FileUID?.Processname} </span>
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "100px",
                                                    maxWidth: "100px",
                                                  }}
                                                >
                                                  {item?.RequestedByTitle}
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "130px", maxWidth: "130px",
                                                    // maxWidth: "100px",
                                                    textAlign: 'center'
                                                  }}
                                                >
                                                 <div>

                                                    {/* {item?.FileUID?.Created} */}
                                                    {new Date(item?.FileUID?.Created).toLocaleString('en-US', {
                                                      month: '2-digit',
                                                      day: '2-digit',
                                                      year: 'numeric',
                                                      hour: '2-digit',
                                                      minute: '2-digit',
                                                      //second: '2-digit',
                                                      hour12: true
                                                    })}
                                                  </div>
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "80px",
                                                    maxWidth: "80px",
                                                    textAlign: 'center'
                                                  }}
                                                >
                                                  <div className="btn btn-status">
                                                    {item?.FileUID?.Status}
                                                  </div>
                                                </td>

                                                <td
                                                  style={{
                                                    minWidth: "50px",
                                                    maxWidth: "50px",
                                                  }}
                                                  className="fe-eye font-18"
                                                >

                                                  <Edit
                                                    onClick={(e) => { getTaskItemsbyID(e, item?.FileUID?.FileUID, item?.FileUID?.Processname); handleShowNestedDMSTable() }}
                                                    style={{


                                                      marginLeft: "15px",

                                                      cursor: "pointer",
                                                    }}
                                                  />
                                                </td>
                                              </tr>
                                            )
                                          )
                                        )}
                                      </tbody>
                                    </table>
                                    {currentData?.length > 0 ? (
                                      <nav className="pagination-container">
                                        <ul className="pagination">
                                          {/* <li
        className={`page-item ${currentPage === 1 ? "disabled" : ""
          }`}
      > */}
                                          <li

                                            className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""
                                              }`}
                                            onClick={() => handleGroupChange("prev")}
                                          >

                                            <a
                                              className="page-link"
                                              // onClick={() =>
                                              //   handlePageChange(currentPage - 1)
                                              // }
                                              aria-label="Previous"
                                            >
                                              «
                                            </a>
                                          </li>
                                          {Array.from(

                                            { length: endPage - startPage + 1 },

                                            (_, num) => {
                                              const pageNum = startPage + num;
                                              return (

                                                <li

                                                  key={pageNum}

                                                  className={`page-item ${currentPage === pageNum ? "active" : ""

                                                    }`}

                                                >

                                                  <a

                                                    className="page-link"

                                                    onClick={() =>

                                                      handlePageChange(pageNum)

                                                    }

                                                  >

                                                    {pageNum}

                                                  </a>

                                                </li>

                                              )
                                            }

                                          )}

                                          <li

                                            className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""

                                              }`}
                                            onClick={() => handleGroupChange("next")}
                                          >

                                            <a

                                              className="page-link"

                                              onClick={() =>

                                                handlePageChange(currentPage + 1)

                                              }

                                              aria-label="Next"

                                            >

                                              »

                                            </a>

                                          </li>

                                        </ul>
                                      </nav>
                                    ) : (
                                      <></>
                                    )}

                                  </div>
                                ) : (
                                  <div>
                                    {folderActionOrFileAction === "New File Request" && (
                                      <>
                                        <DMSMyApprovalAction props={{ currentItemID, actingforuseremail }} />
                                        <div className="col-sm-12 text-center">
                                          {/* <button style={{ float: 'right' }} type="button" className="btn btn-secondary" onClick={() => setShowNestedDMSTable(false)}> Back </button> */}

                                          <button type="button" className="btn cancel-btn newp waves-effect waves-light m-3" style={{ fontSize: '0.875rem' }} onClick={() => setShowNestedDMSTable(false)}>
                                            <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                              className='me-1' alt="x" />
                                            Cancel
                                          </button>
                                        </div>
                                      </>
                                    )}
                                    {folderActionOrFileAction === "New Folder Request" && (
                                      <>
                                        <DMSMyFolderApprovalAction props={{ currentItemID, actingforuseremail }} />
                                        <div className="col-sm-12 text-center">
                                          {/* <button style={{ float: 'right' }} type="button" className="btn btn-secondary" onClick={() => setShowNestedDMSTable(false)}> Back </button> */}

                                          <button type="button" className="btn cancel-btn newp waves-effect waves-light m-3" style={{ fontSize: '0.875rem' }} onClick={() => setShowNestedDMSTable(false)}>
                                            <img src={require('../../../Assets/ExtraImage/xIcon.svg')} style={{ width: '1rem' }}
                                              className='me-1' alt="x" />
                                            Cancel
                                          </button>
                                        </div>
                                      </>
                                    )}

                                  </div>
                                )}
                              </div>
                            )}

                            {/* {activeTab === "ProjectWorkflow" && (
                              <div>
                                <table className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0">
                                  <thead>
                                    <tr>
                                      <th style={{ borderBottomLeftRadius: "0px", minWidth: "40px", maxWidth: "40px", borderTopLeftRadius: "0px" }}>
                                        <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                          <span>S.No.</span>
                                          <span onClick={() => handleSortChange("SNo")}>
                                            <FontAwesomeIcon icon={faSort} />
                                          </span>
                                        </div>
                                        <div className="bd-highlight">
                                          <input
                                            type="text"
                                            placeholder="index"
                                            onChange={(e) => handleFilterChange(e, "SNo")}
                                            onKeyDown={(e) => {
                                              if (e.key === 'Enter' && !e.shiftKey) {
                                                e.preventDefault();
                                              }
                                            }}
                                            className="inputcss"
                                            style={{ width: "100%" }}
                                          />
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Request ID</span>
                                            <span onClick={() => handleSortChange("RequestID")}>
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Request ID"
                                              onChange={(e) => handleFilterChange(e, "RequestID")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Project Name</span>
                                            <span onClick={() => handleSortChange("Title")}>
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Project Name"
                                              onChange={(e) => handleFilterChange(e, "Title")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Process Name</span>
                                            <span onClick={() => handleSortChange("ProcessName")}>
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Process Name"
                                              onChange={(e) => handleFilterChange(e, "ProcessName")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "100px", maxWidth: "100px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Requested By</span>
                                            <span onClick={() => handleSortChange("RequestedBy")}>
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested By"
                                              onChange={(e) => handleFilterChange(e, "RequestedBy")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "130px", maxWidth: "130px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Requested Date</span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested Date"
                                              onChange={(e) => handleFilterChange(e, "RequestedDate")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Status</span>
                                          </div>
                                          <div className="bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Status"
                                              onChange={(e) => handleFilterChange(e, "Status")}
                                              onKeyDown={(e) => {
                                                if (e.key === 'Enter' && !e.shiftKey) {
                                                  e.preventDefault();
                                                }
                                              }}
                                              className="inputcss"
                                              style={{ width: "100%" }}
                                            />
                                          </div>
                                        </div>
                                      </th>

                                      <th style={{ minWidth: "50px", maxWidth: "50px", borderBottomRightRadius: "0px", borderTopRightRadius: "0px", textAlign: "center", verticalAlign: "top" }}>
                                        <div className="d-flex flex-column bd-highlight ">
                                          <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                            <span>Action</span>
                                          </div>
                                        </div>
                                      </th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {((loading && currentData?.length == 0) || (StatusChange)) && (
                                      <div className="loadernewadd">
                                        <div>
                                          <img
                                            src={require("../../../CustomAsset/birdloader.gif")}
                                            className="alignrightl"
                                            alt="Loading..."
                                          />
                                        </div>
                                        <div className="loadnewarg">
                                          <span>Loading </span>{" "}
                                          <span>
                                            <img
                                              src={require("../../myApproval/assets/argloader.gif")}
                                              className="alignrightbird"
                                              alt="Loading..."
                                            />
                                          </span>
                                        </div>
                                      </div>
                                    )}
                                    {!loading && currentData?.length === 0 ? (
                                      <div className="no-results card card-body align-items-center annusvg text-center" style={{ display: "flex", justifyContent: "center", position: 'relative', marginTop: '10px', height: '500px' }}>
                                        <svg style={{ top: '0%' }} xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline></svg>
                                        <p className="font-14 text-muted text-center">No Project Approvals found</p>
                                      </div>
                                    ) : (
                                      !StatusChange && currentData?.map((item: any, index: number) => (
                                        <tr key={index}>
                                          <td style={{ minWidth: "40px", maxWidth: "40px" }}>
                                            <div style={{ marginLeft: "0px" }} className="indexdesign">
                                              {startIndex + index + 1}
                                            </div>
                                          </td>
                                          <td style={{ minWidth: "80px", maxWidth: "80px", textAlign: 'center', textTransform: "capitalize" }} title={item.RequestID}>
                                            {item.RequestID}
                                          </td>
                                          <td style={{ minWidth: "120px", maxWidth: "120px" }} title={item.ProjectName}>
                                            {item.ProjectName}
                                          </td>
                                          <td style={{ minWidth: "120px", maxWidth: "120px", textAlign: 'center' }}>
                                            <span className="badge font-12 bg-secondary">{item.ProcessName}</span>
                                          </td>
                                          <td style={{ minWidth: "100px", maxWidth: "100px" }} title={item.Requester?.Title}>
                                            {item.Requester?.Title}
                                          </td>
                                          <td style={{ minWidth: "130px", maxWidth: "130px", textAlign: 'center' }}>
                                            <div className="btn btn-light1">
                                              {new Date(item.Created).toLocaleString('en-US', {
                                                month: '2-digit',
                                                day: '2-digit',
                                                year: 'numeric',
                                                hour: '2-digit',
                                                minute: '2-digit',
                                                hour12: true
                                              })}
                                            </div>
                                          </td>
                                          <td style={{ minWidth: "80px", maxWidth: "80px", textAlign: 'center' }}>
                                            <div className="btn btn-status">
                                              {item.Status}
                                            </div>
                                          </td>
                                          <td style={{ minWidth: "50px", maxWidth: "50px", textAlign: 'center' }} className="fe-eye font-18">
                                            {item.Status.toLowerCase() == "approved" || item.Status.toLowerCase() == "rejected" || item.Status.toLowerCase() == "completed" ?
                                              <Eye onClick={(e) => handleProjectWorkflowAction(e, item, "view")} style={{ minWidth: "20px", maxWidth: "20px", cursor: "pointer" }} />
                                              :
                                              <Edit onClick={(e) => handleProjectWorkflowAction(e, item, "approval")} style={{ marginLeft: "0px", cursor: "pointer" }} />
                                            }
                                          </td>
                                        </tr>
                                      ))
                                    )}
                                  </tbody>
                                </table>

                                {currentData?.length > 0 ? (
                                  <nav className="pagination-container">
                                    <ul className="pagination">
                                      <li className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""}`} onClick={() => handleGroupChange("prev")}>
                                        <a className="page-link" aria-label="Previous">«</a>
                                      </li>
                                      {Array.from({ length: endPage - startPage + 1 }, (_, num) => {
                                        const pageNum = startPage + num;
                                        return (
                                          <li key={pageNum} className={`page-item ${currentPage === pageNum ? "active" : ""}`}>
                                            <a className="page-link" onClick={() => handlePageChange(pageNum)}>{pageNum}</a>
                                          </li>
                                        );
                                      })}
                                      <li className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""}`} onClick={() => handleGroupChange("next")}>
                                        <a className="page-link" onClick={() => handlePageChange(currentPage + 1)} aria-label="Next">»</a>
                                      </li>
                                    </ul>
                                  </nav>
                                ) : (
                                  <></>
                                )}
                              </div>
                            )} */}

                            {activeTab === "ProjectWorkflow" && (
                              <div>
                                {!showProjectForm ? (
                                  // TABLE VIEW
                                  <div className="card cardCss mt-2">
                                    <div className="card-body">
                                      <div id="cardCollpase4" className="collapse show">
                                        <div className="table-responsive pt-0">
                                          <table className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0" 
                                         style={{ position: "relative",width:'100%' }}
                                >
                                  <thead>
                                    <tr>
                                                <th style={{ borderBottomLeftRadius: "0px", minWidth: "40px", maxWidth: "40px", borderTopLeftRadius: "0px" }}>
                                                  <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                    <span>S.No.</span>
                                                    <span onClick={() => handleSortChange("SNo")}>
                                                      <FontAwesomeIcon icon={faSort} />
                                                    </span>
                                                  </div>
                                                  <div className="bd-highlight">
                                                    <input
                                                      type="text"
                                                      placeholder="index"
                                                      onChange={(e) => handleFilterChange(e, "SNo")}
                                                      onKeyDown={(e) => {
                                                        if (e.key === 'Enter' && !e.shiftKey) {
                                                          e.preventDefault();
                                                        }
                                                      }}
                                                      className="inputcss"
                                                      style={{ width: "100%" }}
                                                    />
                                                  </div>
                                                </th>
                                                {/* Added Checkbox column */}
                                                {currentUserId === 12 && (
                                                  <th style={{ minWidth: "40px", maxWidth: "40px" }}>
                                                    <div className="d-flex flex-column bd-highlight ">
                                                      <div className="d-flex pb-2" style={{ justifyContent: "center" }}>
                                                        <span>Select</span>
                                                      </div>
                                                      <div style={{ height: "33px" }}>        </div>                                             {/* <div className="bd-highlight d-flex justify-content-center">
                                                        <input
                                                          type="checkbox"
                                                          className="form-check-input"
                                                          checked={
                                                            currentData?.length > 0 &&
                                                            currentData.every(item =>
                                                              selectedProjectItems.some(sel => sel.RequestID === item.RequestID)
                                                            )
                                                          }
                                                          onChange={handleSelectAllProjects}
                                                          style={{ cursor: "pointer" }}
                                                        />
                                                      </div> */}
                                                    </div>
                                                  </th>
                                                )}

                                                <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Request ID</span>
                                                      <span onClick={() => handleSortChange("RequestID")}>
                                                        <FontAwesomeIcon icon={faSort} />
                                                      </span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Request ID"
                                                        onChange={(e) => handleFilterChange(e, "RequestID")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Project Name</span>
                                                      <span onClick={() => handleSortChange("Title")}>
                                                        <FontAwesomeIcon icon={faSort} />
                                                      </span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Project Name"
                                                        onChange={(e) => handleFilterChange(e, "Title")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "120px", maxWidth: "120px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Deliverable Name</span>
                                                      <span onClick={() => handleSortChange("DeliverableName")}>
                                                        <FontAwesomeIcon icon={faSort} />
                                                      </span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Deliverable Name"
                                                        onChange={(e) => handleFilterChange(e, "DeliverableName")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "100px", maxWidth: "100px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Requested By</span>
                                                      <span onClick={() => handleSortChange("RequestedBy")}>
                                                        <FontAwesomeIcon icon={faSort} />
                                                      </span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Requested By"
                                                        onChange={(e) => handleFilterChange(e, "RequestedBy")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "130px", maxWidth: "130px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Requested Date</span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Requested Date"
                                                        onChange={(e) => handleFilterChange(e, "RequestedDate")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{ minWidth: "80px", maxWidth: "80px" }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Status</span>
                                                    </div>
                                                    <div className="bd-highlight">
                                                      <input
                                                        type="text"
                                                        placeholder="Filter by Status"
                                                        onChange={(e) => handleFilterChange(e, "Status")}
                                                        onKeyDown={(e) => {
                                                          if (e.key === 'Enter' && !e.shiftKey) {
                                                            e.preventDefault();
                                                          }
                                                        }}
                                                        className="inputcss"
                                                        style={{ width: "100%" }}
                                                      />
                                                    </div>
                                                  </div>
                                                </th>

                                                <th style={{
                                                  minWidth: "50px",
                                                  maxWidth: "50px",
                                                  borderBottomRightRadius: "0px",
                                                  borderTopRightRadius: "0px",
                                                  textAlign: "center",
                                                  verticalAlign: "top"
                                                }}>
                                                  <div className="d-flex flex-column bd-highlight ">
                                                    <div className="d-flex pb-2" style={{ justifyContent: "space-evenly" }}>
                                                      <span>Action</span>
                                                    </div>
                                                  </div>
                                                </th>
                                              </tr>
                                            </thead>
                                            {/* <tbody>
                                              {((loading && currentData?.length == 0) || (StatusChange)) && (
                                                <div className="loadernewadd">
                                                  <div>
                                                    <img
                                                      src={require("../../../CustomAsset/birdloader.gif")}
                                                      className="alignrightl"
                                                      alt="Loading..."
                                                    />
                                                  </div>
                                                  <div className="loadnewarg">
                                                    <span>Loading </span>{" "}
                                                    <span>
                                                      <img
                                                        src={require("../../myApproval/assets/argloader.gif")}
                                                        className="alignrightbird"
                                                        alt="Loading..."
                                                      />
                                                    </span>
                                                  </div>
                                                </div>
                                              )} */}
{/* Ritik 19/3/26 - Added loader in middle of the table for ProjectWorkflow */}
<tbody>
  {((loading && currentData?.length == 0) || (StatusChange)) && (
    <tr>
      <td colSpan={9} style={{ textAlign: "center", padding: "100px 0", border: "none", height:'450px' }}>
        <div className="loadernewadd" style={{ position: "static", display: "flex", flexDirection: "column", alignItems: "center" }}>
          <div>
            <img
              src={require("../../../CustomAsset/birdloader.gif")}
              className="alignrightl"
              alt="Loading..."
            />
          </div>
          <div className="loadnewarg">
            <span>Loading </span>{" "}
            <span>
              <img
                src={require("../../myApproval/assets/argloader.gif")}
                className="alignrightbird"
                alt="Loading..."
              />
            </span>
          </div>
        </div>
      </td>
    </tr>
  )}
                                              {!loading && currentData?.length === 0 ? (
  <tr>
    <td colSpan={9} style={{ border: "none" }}>
      <div className="no-results card card-body align-items-center annusvg text-center" style={{
        display: "flex",
        justifyContent: "center",
        position: 'relative',
        marginTop: '10px',
        height: '500px'
      }}>
        <svg style={{ top: '0%' }} xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
          <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>
        </svg>
        <p className="font-14 text-muted text-center">No Project Approvals found</p>
      </div>
    </td>
  </tr>
  //end of no project approvals found
                                              ) : (
                                                !StatusChange &&
                                                currentData
                                                  ?.slice() // copy array
                                                  .sort((a, b) => {
                                                    const nameA = (a.ProjectName || a.Title || "").toLowerCase();
                                                    const nameB = (b.ProjectName || b.Title || "").toLowerCase();
                                                    return nameA.localeCompare(nameB);
                                                  }).map((item: any, index: number) => (
                                                    <tr key={index}>
                                                      <td style={{ minWidth: "40px", maxWidth: "40px" }}>
                                                      <div className="d-flex align-items-center justify-content-center">   <div style={{ marginLeft: "0px" }} className="indexdesign">
                                                          {startIndex + index + 1}
                                                        </div> </div>
                                                      </td>
                                                      {/* Added Checkbox cell */}
                                                      {currentUserId === 12 && (
                                                        <td style={{ minWidth: "30px", maxWidth: "30px", textAlign: "center" }}>
                                                          <input
                                                            type="checkbox"
                                                            className="form-check-input"
                                                            checked={selectedProjectItems.some(selected => selected.Id === item.Id)}
                                                            onChange={() => handleProjectItemSelect(item.Id)}
                                                            style={{ cursor: "pointer" }}
                                                          />
                                                        </td>
                                                      )}
                                                      <td style={{
                                                        minWidth: "80px",
                                                        maxWidth: "80px",
                                                        textAlign: 'center',
                                                        textTransform: "capitalize"
                                                      }} title={item.RequestID}>
                                                        {item.RequestID}
                                                      </td>
                                                      <td style={{ minWidth: "120px", maxWidth: "120px" }} title={item.ProjectName || item.Title}>
                                                        {item.ProjectName || item.Title}
                                                      </td>
                                                      <td style={{ minWidth: "120px", maxWidth: "120px", textAlign: 'center' }}>
                                                        <span
                                                          className="badge font-12 bg-secondary"
                                                          style={{ whiteSpace: 'normal', height: 'auto', display: 'inline-block' }}
                                                          title={item.Deliverable} // Keep this for accessibility
                                                        >{item.Deliverable}</span>
                                                      </td>
                                                      <td style={{ minWidth: "100px", maxWidth: "100px" }} title={item.Requester?.Title}>
                                                        {item.Requester?.Title}
                                                      </td>
                                                      <td style={{ minWidth: "130px", maxWidth: "130px", textAlign: 'center' }}>
                                                       <div className="btn btn-light1 mt-1" style={{width:'79%'}}>
                                                          {new Date(item.Created).toLocaleString('en-US', {
                                                            month: '2-digit',
                                                            day: '2-digit',
                                                            year: 'numeric',
                                                            hour: '2-digit',
                                                            minute: '2-digit',
                                                            hour12: true
                                                          })}
                                                        </div>
                                                      </td>
                                                      <td style={{ minWidth: "80px", maxWidth: "80px", textAlign: 'center' }}>
                                                        <div className="btn btn-status">
                                                          {item.Status}
                                                        </div>
                                                      </td>
                                                      <td style={{ minWidth: "50px", maxWidth: "50px", textAlign: 'center' }} className="fe-eye font-18">
                                                        <Edit
                                                          onClick={(e) => handleProjectWorkflowAction(e, item, "view")}
                                                          style={{ cursor: "pointer" }}
                                                        />
                                                      </td>
                                                    </tr>
                                                  ))
                                              )}
                                            </tbody>
                                          </table>

                                          {currentData?.length > 0 ? (
                                            <nav className="pagination-container">
                                              <ul className="pagination">
                                                <li className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""}`} onClick={() => handleGroupChange("prev")}>
                                                  <a className="page-link" aria-label="Previous">«</a>
                                                </li>
                                                {Array.from({ length: endPage - startPage + 1 }, (_, num) => {
                                                  const pageNum = startPage + num;
                                                  return (
                                                    <li key={pageNum} className={`page-item ${currentPage === pageNum ? "active" : ""}`}>
                                                      <a className="page-link" onClick={() => handlePageChange(pageNum)}>{pageNum}</a>
                                                    </li>
                                                  );
                                                })}
                                                <li className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""}`} onClick={() => handleGroupChange("next")}>
                                                  <a className="page-link" onClick={() => handlePageChange(currentPage + 1)} aria-label="Next">»</a>
                                                </li>
                                              </ul>
                                            </nav>
                                          ) : (
                                            <></>
                                          )}
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                ) : (
                                  // FORM VIEW - This is where your form goes
                                  <div className="card mt-2">
                                    <div className="card-body">
                                      <div className="form-header d-flex justify-content-between align-items-center mb-3">
                                        <h4 className="text-dark font-16 fw-bold m-0">Project Approval Details</h4>
                                        {/* <button
                                          className="btn btn-secondary"
                                          onClick={handleProjectBackClick}
                                        >
                                          Back
                                        </button> */}
                                      </div>

                                      <div className="row">
                                        <div className="col-md-4">
                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Request ID:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.RequestID || ''}
                                              disabled
                                            />
                                          </div> */}
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Project Name:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ProjectName || selectedProjectTask?.Title || ''}
                                              disabled
                                            />
                                          </div>



                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Process Name:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ProcessName || ''}
                                              disabled
                                            />
                                          </div> */}
                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Requested By:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Requester?.Title || ''}
                                              disabled
                                            />
                                          </div> */}
                                        </div>
                                        <div className="col-md-4">
                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Requested Date:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={new Date(selectedProjectTask?.Created).toLocaleString() || ''}
                                              disabled
                                            />
                                          </div> */}
                                          {/* <div className="mb-3">
                                            <label className="form-label"><strong>Status:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Status || ''}
                                              disabled
                                            />
                                          </div> */}
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Project Type:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ProjectType || 'N/A'}
                                              disabled
                                            />
                                          </div>

                                        </div>
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Deliverable:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Deliverable || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>


                                        {/* Additional fields from your data */}

                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Document Type:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.DocType || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Document Number:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.DocNumber || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Area:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Area || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Organization:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.Org || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>


                                        {/* Client and Prepared By fields */}

                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Client Name:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ClientName || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Prepared By:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.PreparedBy || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div>


                                        {/* Approval Role and Revision Number */}

                                        {/* <div className="col-md-6">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Approval Role:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.ApprovalRole || 'N/A'}
                                              disabled
                                            />
                                          </div>
                                        </div> */}
                                        <div className="col-md-4">
                                          <div className="mb-3">
                                            <label className="form-label"><strong>Revision Number:</strong></label>
                                            <input
                                              type="text"
                                              className="form-control"
                                              value={selectedProjectTask?.RevisionNumber || '0'}
                                              disabled
                                            />
                                          </div>
                                        </div>
                                      </div>

                                      {/* Document section */}
                                      {/* {projectDocumentInfo?.documentUrl && (
                                        <div className="row">
                                          <div className="col-12">
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Document:</strong></label>
                                              <div className="document-container">
                                                {projectDocumentInfo ? (
                                                  <div
                                                    className="document-link-container p-3 border rounded bg-light"
                                                    style={{ cursor: 'pointer' }}
                                                    onClick={handleProjectOpenDocument}
                                                    title="Click to open document"
                                                  >
                                                    <div className="d-flex align-items-center">
                                                      <span className="document-icon me-2" style={{ fontSize: '1.5rem' }}>📄</span>
                                                      <div>
                                                        <div className="document-name fw-bold">
                                                          {projectDocumentInfo.fileName || projectDocumentInfo.fileLeafRef}
                                                        </div>
                                                        <div className="document-hint text-muted small">
                                                          Click to open document in new tab
                                                          {projectDocumentInfo.sharedLink && (
                                                            <span className="ms-2">🔗 Shared Link Available</span>
                                                          )}
                                                        </div>
                                                      </div>
                                                    </div>
                                                  </div>
                                                ) : projectDocumentInfo?.documentUrl ? (
                                                  <div
                                                    className="document-link-container p-3 border rounded bg-light"
                                                    style={{ cursor: 'pointer' }}
                                                    onClick={() => {
                                                      if (projectDocumentInfo?.documentUrl) {
                                                        window.open(projectDocumentInfo?.documentUrl, '_blank', 'noopener,noreferrer');
                                                      }
                                                    }}
                                                    title="Click to open document"
                                                  >
                                                    <div className="d-flex align-items-center">
                                                      <span className="document-icon me-2" style={{ fontSize: '1.5rem' }}>📄</span>
                                                      <div>
                                                        <div className="document-name fw-bold">Document Available</div>
                                                        <div className="document-hint text-muted small">Click to open document</div>
                                                      </div>
                                                    </div>
                                                  </div>
                                                ) : (
                                                  <div className="alert alert-info mb-0">
                                                    <small>No document available for this deliverable</small>
                                                  </div>
                                                )}
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      )} */}

                                      {projectDocumentInfo && projectDocumentInfo.length > 0 ? (
                                        <div className="row">
                                          <div className="col-12">
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Documents:</strong></label>
                                              <div className="document-container">
                                                {projectDocumentInfo.map((doc, index) => (
                                                  <div
                                                    key={index}
                                                    className="document-link-container p-3 mb-2 border rounded bg-light"
                                                    style={{ cursor: 'pointer' }}
                                                    onClick={() => {
                                                      if (doc.sharedLink) {
                                                        window.open(doc.sharedLink, '_blank', 'noopener,noreferrer');
                                                      }
                                                    }}
                                                    title="Click to open document"
                                                  >
                                                    <div className="d-flex align-items-center">
                                                      <span className="document-icon me-2" style={{ fontSize: '1.5rem' }}>📄</span>
                                                      <div>
                                                        <div className="document-name fw-bold">
                                                          {doc.fileName || doc.fileLeafRef || "Document Available"}
                                                        </div>
                                                        <div className="document-hint text-muted small">
                                                          Click to open document in new tab
                                                          {doc.sharedLink && (
                                                            <span className="ms-2"></span>
                                                          )}
                                                        </div>
                                                      </div>
                                                    </div>
                                                  </div>
                                                ))}
                                              </div>
                                            </div>
                                          </div>
                                        </div>
                                      ) : (
                                        <div className="alert alert-info mb-0">
                                          <small>No documents available for this deliverable</small>
                                        </div>
                                      )}


                                      {/* Remarks field (if available) */}
                                      {selectedProjectTask?.Remarks && (
                                        <div className="row">
                                          <div className="col-12">
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Remarks:</strong></label>
                                              <textarea
                                                className="form-control"
                                                value={selectedProjectTask?.Remarks || ''}
                                                disabled
                                                rows={3}
                                              />
                                            </div>
                                          </div>
                                        </div>
                                      )}

                                      {/* <div className="row mt-4">
                                        <div className="col-12 text-center">
                                          <p className="text-muted">
                                            <em>Approval actions will be implemented in the next phase</em>
                                          </p>
                                        </div>
                                      </div> */}
                                      {/* Conditional Fields based on Approver Role */}
                                      <div className="row">
                                        <div className="col-12">
                                          {selectedProjectTask?.ApprovalRole === "Document Controller" && (
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Do you need further approval?</strong></label>
                                              <select
                                                value={projectNeedsFurtherApproval}
                                                onChange={(e) => setProjectNeedsFurtherApproval(e.target.value)}
                                                className="form-select"
                                              >
                                                <option value="Select" disabled>Select</option>
                                                <option value="Yes">Yes</option>
                                                <option value="No">No</option>
                                              </select>
                                            </div>
                                          )}
                                          {selectedProjectTask?.ApprovalRole === "Document Controller" && (
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Outgoing status</strong></label>
                                              <select
                                                value={projectOutgoingstatus}
                                                onChange={(e) => setprojectOutgoingstatus(e.target.value)}
                                                className="form-select"
                                              >
                                                <option value="Select" disabled>Select</option>
                                                <option value="A">A</option>
                                                <option value="B">B</option>
                                                <option value="C">C</option>
                                              </select>
                                            </div>
                                          )}

                                          {selectedProjectTask?.ApprovalRole === "DCC" && (
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Want to publish in Dossier?</strong></label>
                                              <select
                                                value={projectWantsToPublishInDossier}
                                                onChange={(e) => setProjectWantsToPublishInDossier(e.target.value)}
                                                className="form-select"
                                              >
                                                <option value="Select" disabled>Select</option>
                                                <option value="Yes">Yes</option>
                                                <option value="No">No</option>
                                              </select>
                                            </div>
                                          )}
                                          {selectedProjectTask?.ApprovalRole === "DCC" && (
                                            <div className="mb-3">
                                              <label className="form-label"><strong>Code status</strong></label>
                                              <select
                                                value={projectCodestatus}
                                                onChange={(e) => setprojectCodestatus(e.target.value)}
                                                className="form-select"
                                              >
                                                <option value="Select" disabled>Select</option>
                                                <option value="A">A</option>
                                                <option value="B">B</option>
                                                <option value="C">C</option>
                                              </select>
                                            </div>
                                          )}

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
                                                  <div className={styles.documentCommentsHeader}>
                                                    <select
                                                      className={styles.formSelect}
                                                      value={selectedVersion}
                                                      onChange={(e) => onVersionChange(e.target.value)}
                                                    >
                                                      <option value="">-- Select Revision --</option>
                                                      {versionList.map(version => (
                                                        <option key={version} value={version}>{version}</option>
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
                                                          {/* <th style={{ minWidth: '80px', maxWidth: '80px' }}>Document Name</th> */}
                                                          <th style={{ minWidth: '80px', maxWidth: '80px' }}>Users</th>
                                                          <th style={{ minWidth: '100px', maxWidth: '100px' }}>Comment Date</th>
                                                          <th style={{ minWidth: '80px', maxWidth: '80px' }}>Page No.</th>
                                                          <th style={{ minWidth: '80px', maxWidth: '80px' }}>Revision</th>
                                                          <th style={{ minWidth: '200px', maxWidth: '200px' }}>Comments</th>
                                                        </tr>
                                                      </thead>
                                                      <tbody style={{ maxHeight: '250px', overflowY: 'auto' }}>
                                                        {documentComments.map((commentItem) => (
                                                          <tr key={commentItem.id}>
                                                            {/* <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
                                                              {commentItem.DocumentName}
                                                            </td> */}
                                                            <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
                                                              {commentItem.userName}
                                                            </td>
                                                            <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '100px', maxWidth: '100px' }}>
                                                              {commentItem.commentDate}
                                                            </td>
                                                            <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
                                                              {commentItem.pageNumber}
                                                            </td>
                                                            <td style={{ padding: '10px', verticalAlign: 'top', minWidth: '80px', maxWidth: '80px' }}>
                                                              {commentItem.revision}
                                                            </td>
                                                            <td style={{ padding: '15px', verticalAlign: 'top', minWidth: '200px', maxWidth: '200px' }}>
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

                                          <div className="mb-3">
                                            <label className="form-label"><strong>Remarks</strong></label>
                                            <textarea
                                              value={projectRemarks}
                                              onChange={(e) => setProjectRemarks(e.target.value)}
                                              placeholder="Enter your remarks here..."
                                              rows={4}
                                              className="form-control"
                                            />
                                          </div>
                                        </div>
                                      </div>

                                      {/* Dynamic Approval Hierarchy Table */}
                                      {selectedProjectTask?.ApprovalRole === "Document Controller" && (
                                        <div className="card">
                                          <div className="card-body">
                                            <div className="approval-projectHierarchy mt-4">
                                              <div className="d-flex justify-content-between align-items-center mb-3">
                                                <h5 style={{ margin: 'inherit' }}>Approval Hierarchy</h5>
                                                <button
                                                  type="button"
                                                  className="btn btn-primary btn-sm"
                                                  onClick={() => addNewProjectApprovalRow()}
                                                >
                                                  + Add New Row
                                                </button>
                                              </div>
                                              <div className="table-responsive">
                                                <table className="table table-bordered">
                                                  <thead className="table-light">
                                                    <tr>
                                                      <th style={{ width: '15%' }}>Level</th>
                                                      <th style={{ width: '25%' }}>Approver Role</th>
                                                      <th style={{ width: '30%' }}>Approver</th>
                                                      <th style={{ width: '20%' }}>Approval Criteria</th>
                                                      <th style={{ width: '10%' }}>Action</th>
                                                    </tr>
                                                  </thead>
                                                  <tbody>
                                                    {projectHierarchy.map((row, index) => (
                                                      <tr key={index}>
                                                        <td>
                                                          <input
                                                            type="text"
                                                            className="form-control"
                                                            value={row.level}
                                                            disabled
                                                          />
                                                        </td>
                                                        <td>
                                                          <select
                                                            value={row.approverRole}
                                                            onChange={(e) => updateProjectApprovalRow(index, 'approverRole', e.target.value)}
                                                            className="form-select"
                                                          >
                                                            <option value="" disabled>Select Role</option>
                                                            <option value="Project Coordinator">Project Coordinator</option>
                                                            <option value="Project Team">Project Team</option>
                                                            <option value="Project Manager">Project Manager</option>
                                                          </select>
                                                        </td>
                                                        <td>
                                                          <Select
                                                            isMulti
                                                            options={users}
                                                            value={row.assignedTo ? users.filter(user =>
                                                              row.assignedTo?.some((assigned: any) => {
                                                                const assignedId = assigned.ID ? assigned.ID.toString() : assigned.toString();
                                                                return assignedId === user.value;
                                                              })
                                                            ) : []}
                                                            onChange={(selectedOptions: any) => handleProjectApproverChange(index, selectedOptions)}
                                                            placeholder="Select Approver(s)"
                                                            className="people-picker"
                                                            classNamePrefix="react-select"
                                                            closeMenuOnSelect={false}
                                                            isClearable={false}
                                                          />
                                                        </td>
                                                        <td>
                                                          <select
                                                            value={row.approvalCriteria}
                                                            onChange={(e) => updateProjectApprovalRow(index, 'approvalCriteria', e.target.value)}
                                                            className="form-select"
                                                          >
                                                            <option value="" disabled>Select Criteria</option>
                                                            <option value="Everyone">Everyone</option>
                                                            <option value="Anyone">Anyone</option>
                                                          </select>
                                                        </td>
                                                        <td className="text-center">
                                                          <button
                                                            type="button"
                                                            className="btn btn-outline-danger btn-sm"
                                                            onClick={() => deleteProjectApprovalRow(index)}
                                                            title="Delete Row"
                                                          >
                                                            🗑️
                                                          </button>
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

                                      {/* Approval Action Buttons */}
                                      <div className="row mt-4">
                                        <div className="col-12">
                                          <div className="d-flex justify-content-center gap-3">
                                            {selectedProjectTask?.Status === "Pending" && (
                                              <>
                                                {selectedProjectTask?.ApprovalRole === "DCC" ? (
                                                  <>
                                                    <button
                                                      type="button"
                                                      className="btn btn-success"
                                                      onClick={() => handleProjectApprovalAction("Approved")}
                                                      disabled={showSubmitLoader}
                                                    >
                                                      Submit
                                                    </button>
                                                    <button
                                                      type="button"
                                                      className="btn btn-secondary"
                                                      onClick={handleProjectBackClick}
                                                    >
                                                      Cancel
                                                    </button>
                                                  </>
                                                ) : (
                                                  <>
                                                    <button
                                                      type="button"
                                                      className="btn btn-success"
                                                      disabled={showSubmitLoader}
                                                      onClick={() => {
                                                        setPendingAction("Approved");
                                                        setPopupType("confirmation");
                                                        setPopupTitle("Confirm Approval");
                                                        setPopupMessage("Are you sure you want to approve this request?");
                                                        setPopupOpen(true);
                                                      }}
                                                    >
                                                      Approve
                                                    </button>

                                                    <button
                                                      type="button"
                                                      className="btn btn-danger"
                                                      disabled={showSubmitLoader}
                                                      onClick={() => {
                                                        setPendingAction("Rejected");
                                                        setPopupType("confirmation");
                                                        setPopupTitle("Confirm Rejection");
                                                        setPopupMessage("Are you sure you want to reject this request?");
                                                        setPopupOpen(true);
                                                      }}
                                                    >
                                                      Reject
                                                    </button>

                                                    <button
                                                      type="button"
                                                      className="btn btn-warning"
                                                      disabled={showSubmitLoader}
                                                      onClick={() => {
                                                        setPendingAction("Rework");
                                                        setPopupType("confirmation");
                                                        setPopupTitle("Send for Rework");
                                                        setPopupMessage("Do you want to send this item for rework?");
                                                        setPopupOpen(true);
                                                      }}
                                                    >
                                                      Rework
                                                    </button>

                                                    <button
                                                      type="button"
                                                      className="btn btn-secondary"
                                                      onClick={handleProjectBackClick}
                                                    >
                                                      Cancel
                                                    </button>
                                                  </>
                                                )}
                                              </>
                                            )}

                                            {(selectedProjectTask?.Status === "Approved" ||
                                              selectedProjectTask?.Status === "Rejected" ||
                                              selectedProjectTask?.Status === "Rework") && (
                                                <button type="button" className="btn btn-secondary" onClick={handleProjectBackClick}>
                                                  Back
                                                </button>
                                              )}
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                )}
                              </div>
                            )}
                            {/* {activeTab === "DMS" ?
               (
               <div>
                    <div className="DMSMasterContainer">
                <h4 className="page-title fw-bold mb-1 font-20">My Approvals 1</h4>
                <div className="" style={{ backgroundColor: 'white', border:'1px solid #54ade0', marginTop:'20px', borderRadius:'20px', padding: '15px'}}>
                <table style={{width:'100%'}} className="mtbalenew">
    <thead>
      <tr>
        <th
          style={{
            minWidth: '40px',
            maxWidth: '40px',
         
          }}
        >
          S.No
        </th>
        <th>Request ID</th>
        <th>Process Name</th>
        <th>Requested By</th>
        <th >Requested Date</th>
        <th style={{ minWidth: '80px', maxWidth: '80px' }}>Status</th>
        <th
          style={{
            minWidth: '70px',
            maxWidth: '70px',
         
          }}
        >
          Action
        </th>
      </tr>
    </thead>
    <tbody style={{ maxHeight: '8007px' }}>
       
      {Mylistdata.length > 0  ? Mylistdata.map((item, index) => {
      return(
        <tr>
<td style={{ minWidth: '40px', maxWidth: '40px'}}>
  <span style={{marginLeft:'5px'}} className="indexdesign">{index}</span>
  </td>
<td >{(truncateText(item.FileUID.FileUID, 22))}</td>
<td >{item?.ProcessName}</td>
<td >{(truncateText(item.FileUID.RequestedBy, 22))}</td> 
<td >
<div
  style={{
    padding: '5px',
    border: '1px solid #efefef',
    background: '#fff', fontSize:'14px',
    borderRadius: '30px',
  
  }}
  className="btn btn-light"
>
 {item.FileUID.Created}
</div>
</td>
<td style={{ minWidth: '80px', maxWidth: '80px', textAlign:'center' }}>
<div className="finish mb-0">Pending</div>
</td>
<td style={{ minWidth: '70px', maxWidth: '70px' }}>
  {item?.ProcessName === "DMS Folder Approval" ?
    (<a onClick={(e )=>getTaskItemsbyID2(e , item.FileUID.FileUID)}>
    <FontAwesomeIcon icon={faEye} />
   </a>
   ) : item?.ProcessName === "" || item?.ProcessName === null || item?.ProcessName === undefined ?    (
      <a onClick={(e )=>getTaskItemsbyID(e , item.FileUID.FileUID)}>
 <FontAwesomeIcon icon={faEye} />
</a>
    ) : null
  }

</td>
</tr>
      )

       })
       :""

}

      
   

     
    </tbody>
  </table>
        </div>
              </div>
               </div>
               ) : (
                <div>
                  {activeComponent === 'Approval Action' ? (
                    <div>
                   <button style={{float:'right'}} type="button" className="btn btn-secondary" onClick={()=>handleReturnToMain('')}> Back </button>
                  <DMSMyApprovalAction props={currentItemID}/>
                    </div>
               
                  ) : activeComponent === 'DMS Folder Approval' ? (
                    <div>
<button style={{float:'right'}} type="button" className="btn btn-secondary" onClick={()=>handleReturnToMain('')}> Back </button>
<DMSMyFolderApprovalAction props={currentItemID}/>
                    </div>
                                       
                  ) :null
                  
                  } 
             
                </div>
               
            
               )
               } */}
                          </div>


                        </div>
                      </div>
                    </div>
                  )}

                  {isActivedata == true &&
                    ContentData.length > 0 &&
                    currentItem != null && (
                      <div className="row mt-4">
                        <div className="col-12">
                          <div className="card">
                            <div className="card-body">
                              <h4 className="header-title mb-0">
                                {ContentData[0].Title}
                              </h4>

                              <p className="sub-header">
                                {currentItem.EntityName}
                              </p>

                              <div className="row">
                                <div className="col-lg-4">
                                  <div className="mb-3">
                                    <label className="form-label text-dark font-14">
                                      Company / Department:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {currentItem.EntityName}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="col-lg-4">
                                  <div className="mb-3">
                                    <label className="form-label text-dark font-14">
                                      Date of Request:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {currentItem.Created}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="col-lg-4">
                                  <div className="mb-3">
                                    <label className="form-label text-dark font-14">
                                      Status:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {currentItem.Status}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="col-lg-6">
                                  <div className="mb-0">
                                    <label className="form-label text-dark font-14">
                                      Content:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {ContentData[0].Title ||
                                          ContentData[0].EventName}
                                      </span>
                                    </div>
                                  </div>
                                </div>

                                <div className="col-lg-6">
                                  <div className="mb-0">
                                    <label className="form-label text-dark font-14">
                                      Overview:
                                    </label>

                                    <div>
                                      <span className="text-muted font-14">
                                        {ContentData[0].Overview}
                                      </span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>

                          {ContentData[0]?.Description != null && (
                            <div className="card">
                              <div className="card-body">
                                <h4 className="header-title mb-0">Description</h4>

                                <p className="sub-header">
                                  <div
                                    dangerouslySetInnerHTML={{
                                      __html: ContentData[0].Description,
                                    }}
                                  ></div>
                                </p>
                              </div>
                            </div>
                          )}

                          {currentItem.Status == "Submitted" && (
                            <div className="card">
                              <div className="card-body">
                                <div className="row">
                                  {currentItem.Status == "Submitted" && (
                                    <div className="col-lg-12">
                                      <div className="mb-0">
                                        <label
                                          htmlFor="example-textarea"
                                          className="form-label text-dark font-14"
                                        >
                                          Remarks:
                                        </label>

                                        <textarea
                                          className="form-control"
                                          id="example-textarea"
                                          rows={5}
                                          name="Remark"
                                          value={formData.Remark}
                                          onChange={(e) =>
                                            onChange(
                                              e.target.name,
                                              e.target.value
                                            )
                                          }
                                        ></textarea>
                                      </div>
                                    </div>
                                  )}
                                </div>

                                {currentItem.Status == "Submitted" && (
                                  <div className="row mt-3">
                                    <div className="col-12 text-center">
                                      <a href="my-approval.html">
                                        <button
                                          type="button"
                                          className="btn btn-success waves-effect waves-light m-1"
                                          onClick={(e) =>
                                            handleFromSubmit(e, "Approved")
                                          }
                                        >
                                          <i className="fe-check-circle me-1"></i>{" "}
                                          Approve
                                        </button>
                                      </a>

                                      <a href="my-approval.html">
                                        <button
                                          type="button"
                                          className="btn btn-warning waves-effect waves-light m-1"
                                          onClick={(e) =>
                                            handleFromSubmit(e, "Rework")
                                          }
                                        >
                                          <i className="fe-corner-up-left me-1"></i>{" "}
                                          Rework
                                        </button>
                                      </a>

                                      <a href="my-approval.html">
                                        <button
                                          type="button"
                                          className="btn btn-danger waves-effect waves-light m-1"
                                          onClick={(e) =>
                                            handleFromSubmit(e, "Reject")
                                          }
                                        >
                                          <i className="fe-x-circle me-1"></i>{" "}
                                          Reject
                                        </button>
                                      </a>

                                      <button
                                        type="button"
                                        className="btn cancel-btn waves-effect waves-light m-1"
                                        onClick={(e) => handleCancel()}
                                      >
                                        <i className="fe-x me-1"></i> Cancel
                                      </button>
                                    </div>
                                  </div>
                                )}
                              </div>
                            </div>
                          )}
                        </div>
                      </div>
                    )}
                </div>
              )}
          </div>
        </div>
      </div>
      <CustomPopup
        isOpen={popupOpen}
        type={popupType}
        title={popupTitle}
        message={popupMessage}
        onConfirm={async () => {
          if (!pendingAction) return;

          setPopupOpen(false);
          await handleProjectApprovalAction(pendingAction);
        }}
        onCancel={() => {
          setPopupOpen(false);
          setPendingAction(null);
        }}
        onClose={() => {
          setPopupOpen(false);
          setPendingAction(null);
        }}
        onSuccessOk={() => {
          setPopupOpen(false);
          window.location.reload();
        }}
      />


    </div>
  );
};

const MyApproval: React.FC<IMyApprovalProps> = (props) => (
  <Provider>
    <MyApprovalContext props={props} />
  </Provider>
);

export default MyApproval;