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
import Select from "react-select";
import loaderGif from "../assets/Loder.gif"; //priyanshu
import CustomPopup from "../../myProject/components/CustomPopup";
import { ErrorLogger } from "../../../utils/ErrorLogger";

interface ApprovalHierarchyItem {
  id?: number;
  level: string;
  approverRole: string;
  approver: string;
  approvers: string[];
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
  isDeleted: Boolean;
}

interface UserOption {
  value: string;
  label: string;
  email: string;
}

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
  ApprovalRole: string;
}

let actingforuseremail: any;
const MyApprovalContext = ({ props }: any) => {
  const sp: SPFI = getSP(props.context);
  const [activeComponent, setActiveComponent] = useState<string>("");
  const { useHide }: any = React.useContext(UserContext);
  const [showNestedDMSTable, setShowNestedDMSTable] = useState(false);
  const [announcementData, setAnnouncementData] = React.useState([]);

  const [myApprovalsDataAll, setMyApprovalsDataAll] = React.useState([]);
  const [myApprovalsDataAutomation, setMyApprovalsDataAutomation] =
    React.useState([]);
  const [Mylistdata, setMylistdata] = useState<any[]>([]);

  const handleShowNestedDMSTable = () => {
    setShowNestedDMSTable(true); // Show nested table within DMS
  };
  const elementRef = React.useRef<HTMLDivElement>(null);

  const [folderActionOrFileAction, setFolderActionOrFileAction] = useState("");
  const [newsData, setNewsData] = React.useState([]);
  const [isOpen, setIsOpen] = React.useState(false);

  const [IsinvideHide, setIsinvideHide] = React.useState(false);

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

  // project workflow states and variables

  const SiteUrl = props.siteUrl;

  const [myApprovalsData, setMyApprovalsData] = React.useState<any[]>([]);

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

  const [ProjectWorkflowdata, setProjectWorkflowdata] = useState([]);
  const [selectedProjectTask, setSelectedProjectTask] =
    React.useState<any>(null);
  const [showProjectForm, setShowProjectForm] = React.useState(false);
  const [projectNeedsFurtherApproval, setProjectNeedsFurtherApproval] =
    React.useState<string>("Select");
  const [projectOutgoingstatus, setprojectOutgoingstatus] =
    React.useState<string>("Select");
  const [projectCodestatus, setprojectCodestatus] =
    React.useState<string>("Select");
  const [projectWantsToPublishInDossier, setProjectWantsToPublishInDossier] =
    React.useState<string>("Select");
  const [projectRemarks, setProjectRemarks] = React.useState<string>("");
  const [projectHierarchy, setProjectHierarchy] = React.useState<
    ApprovalHierarchyItem[]
  >([]);
  const [projectDocumentInfo, setProjectDocumentInfo] = React.useState<
    {
      documentUrl: string;
      fileName: string;
      fileLeafRef: string;
      fileRef: string;
      sharedLink: string;
      id: number;
      isConsolidatorCopy?: string;
      isMyTask?: string;
    }[]
  >([]);

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
  const [users, setUsers] = React.useState<UserOption[]>([]);
  const [selectedProjectItems, setSelectedProjectItems] = useState<
    ProjectItem[]
  >([]);
  const [outgoingStatus, setOutgoingStatus] = React.useState<string>("");
  const [mdrStatus, setMdrStatus] = React.useState<string>("");
  const [currentUserId, setCurrentUserId] = React.useState<number | null>(null);

  // ========== UI CONTROL STATE ==========

  const [showSubmitLoader, setShowSubmitLoader] =
    React.useState<boolean>(false);
  const [popupOpen, setPopupOpen] = React.useState(false);
  const [popupType, setPopupType] = React.useState<
    "confirmation" | "validation" | "success" | "error"
  >("confirmation");

  const [popupTitle, setPopupTitle] = React.useState("");
  const [popupMessage, setPopupMessage] = React.useState("");
  const [pendingAction, setPendingAction] = React.useState<
    "Approved" | "Rejected" | "Rework" | null
  >(null);

  // Additional state variables for Project Workflow
  const [approverRoles, setApproverRoles] = useState<
    { value: string; label: string }[]
  >([]);
  const [fieldErrors, setFieldErrors] = React.useState<{
    [key: string]: boolean;
  }>({});
  const [projectConfigurationsData, setProjectConfigurationsData] =
    React.useState<any[]>([]);

  const [documentControllerId, setDocumentControllerId] = React.useState<
    number | null
  >(null);
  const [dccId, setDccId] = React.useState<number | null>(null);
  const [sortConfig, setSortConfig] = React.useState({
    key: "",
    direction: "ascending",
  });

  const [isReadOnly, setIsReadOnly] = useState<boolean>(false);

  type ActionType =
    | "APPROVAL_ASSIGNED"
    | "APPROVED"
    | "REWORK_REQUESTED"
    | "SENT_FOR_PUBLISH"
    | "PUBLISHED"
    | "REWORK_BY_APPROVER"
    | "REWORK_BY_DOCUMENT_CONTROLLER";

  const actionConfig: Record<ActionType, { subject: string; message: string }> =
    {
      APPROVAL_ASSIGNED: {
        subject: "Approval Action Assigned",
        message:
          "An approval action has been assigned to you for the following deliverable(s).",
      },
      APPROVED: {
        subject: "Deliverable Approved",
        message: "The deliverable(s) mentioned below have been approved.",
      },
      REWORK_REQUESTED: {
        subject: "Rework Requested",
        message: "Rework has been requested for the following deliverable(s).",
      },
      SENT_FOR_PUBLISH: {
        subject: "Deliverable Pending Publication",
        message: "The deliverable(s) mentioned below are pending publication.",
      },

      PUBLISHED: {
        subject: "Deliverable Published",
        message: "The deliverable(s) mentioned below have been published.",
      },
      REWORK_BY_APPROVER: {
        subject: "Rework Requested by Approver",
        message:
          "The Approver has requested rework for the following deliverable(s).",
      },
      REWORK_BY_DOCUMENT_CONTROLLER: {
        subject: "Rework Requested by Document Controller",
        message:
          "The Document Controller has requested rework for the following deliverable(s).",
      },
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
  const [selectedBehalfEmail, setSelectedBehalfEmail] = useState("");

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
              "Log",
              "CurrentUser",
              "Remark",
              "LogHistory",
              "FileUID/FileUID",
              "FileUID/SiteName",
              "FileUID/DocumentLibraryName",
              "FileUID/FileName",
              "FileUID/RequestNo",
              "FileUID/Processname",
              "FileUID/Status",
              "FileUID/FolderPath",
              "FileUID/RequestedBy",
              "FileUID/Created",
              "FileUID/ApproveAction",
              "MasterApproval/ApprovalType",
              "MasterApproval/Level",
              "MasterApproval/DocumentLibraryName",
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
              "*",
              "Folderdetail/SiteTitle",
              "Folderdetail/DocumentLibraryName",
              "Folderdetail/CurrentUser",
              "Folderdetail/FolderPath",
              "Folderdetail/FolderName",
              "Folderdetail/RequestNo",
              "Folderdetail/Processname",
              "Folderdetail/Status",
              "FolderMeta/ID",
              "FolderMeta/SiteName",
              "FolderMeta/DocumentLibraryName",
              "Approver",
            )
            .expand("Folderdetail", "FolderMeta")
            .filter(
              `Approver eq '${targetUser}' and Folderdetail/Status eq '${value}'`,
            )();

          // Normalize Folder Items (Same as your existing logic)
          const normalizedFolders = await Promise.all(
            folderItems.map(async (item) => ({
              Log: item?.Log || "",
              CurrentUser: item?.Folderdetail?.CurrentUser || "",
              Remark: item?.Remark || "",
              RequestedByTitle: await getUserTitleByEmail(
                item?.Folderdetail?.CurrentUser,
              ),
              FileUID: {
                FileUID:
                  item?.FolderMeta?.FileUID || item?.Folderdetail?.RequestNo,
                SiteName: item?.FolderMeta?.SiteName || "",
                FileName:
                  item?.Folderdetail?.FolderName ||
                  item?.Folderdetail?.DocumentLibraryName,
                // Status: item?.Folderdetail?.Status || '',
                // srs 18/3/26
                Status:
                  item?.Folderdetail?.Log || item?.Folderdetail?.Status || "",
                Processname: item?.Folderdetail?.Processname,
                // ... add other fields as per your normalizeItem2
              },
              MasterApproval: {
                ApprovalType: item?.ApprovalType || "",
                Level: item?.Level || "",
              },
            })),
          );

          // Enrich File Items with User Titles
          const enrichedFiles = await Promise.all(
            fileItems.map(async (item) => {
              const requestedbyuserTitle = await getUserTitleByEmail(
                item?.FileUID?.RequestedBy,
              );
              // return { ...item, RequestedByTitle: requestedbyuserTitle };
              // srs 18/3/26
              return {
                ...item,
                RequestedByTitle: requestedbyuserTitle,
                FileUID: {
                  ...item.FileUID,
                  // BIND LOG TO STATUS COLUMN
                  Status: item.Log || item?.FileUID?.Status || "",
                },
              };
            }),
          );

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
      // setMylistdata(allCombinedItems);
      // // srs 10/4/26
      // if (activeTab === "DMS") {
      //   setMyApprovalsData(allCombinedItems);
      //   setLoading(false); // Stop the bird loader
      // }
      return allCombinedItems;
    } catch (error) {
      console.error("Master Site Collection fetch failed:", error);
    }
  };

  console.log(Mylistdata, "Mylistdata");
  const currentUserEmailRef = useRef("");

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
    currentItemID = itemid;
    setActiveComponent("DMS Folder Approval");
    console.log("itemid", itemid);
    // const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select("CurrentUser" , "FileUID/FileUID" , "Log").expand("FileUID").filter(`FileUID/RequestNo eq '${itemid}'`)();
    //    console.log(items , "items")
  };
  const toggleDropdown = () => {
    setIsOpen(!isOpen);
  };

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

  //const [activeTab, setActiveTab] = useState("home1");
  // const [activeTab, setActiveTab] = useState("Automation");
  // srs 10/4/26
  const [activeTab, setActiveTab] = useState("DMS");
  const handleTabClick = async (tab: React.SetStateAction<string>) => {
    if (activeTab === tab) return;
    setActiveTab(tab);
    setSelectedBehalfEmail("");
    actingforuseremail = "";
    setLoading(true);
    setMyApprovalsData([]);
    setStatusChange(true);

    try {
      let MyApprovaldata: any = [];
      let Automationdata: any = [];
      let MyDMSAPPROVALDATA: any = [];
      let ProjectWorkflowData: any = [];

      if (
        actingforuseremail === undefined ||
        actingforuseremail === null ||
        actingforuseremail === ""
      ) {
        [MyApprovaldata, Automationdata, ProjectWorkflowData] =
          await Promise.all([
            getMyApproval(sp, Statusvalue),
            getApprovalListsData(sp, Statusvalue),
            getProjectWorkflowApprovals(Statusvalue),
          ]);

        if (tab === "DMS") {
          MyDMSAPPROVALDATA = await getApprovalmasterTasklist(Statusvalue);
        }
      } else {
        [MyApprovaldata, Automationdata, ProjectWorkflowData] =
          await Promise.all([
            getMyApproval(sp, Statusvalue, actingforuseremail),

            getApprovalListsData(sp, Statusvalue, actingforuseremail),

            getProjectWorkflowApprovals(Statusvalue, actingforuseremail),
          ]);

        if (tab === "DMS") {
          MyDMSAPPROVALDATA = await getApprovalmasterTasklist(
            Statusvalue,
            actingforuseremail,
          );
        }
      }

      setMyApprovalsDataAll(MyApprovaldata);
      setMyApprovalsDataAutomation(Automationdata);
      setProjectWorkflowdata(ProjectWorkflowData);
      if (tab === "DMS") {
        setMylistdata(MyDMSAPPROVALDATA);
      }

      if (tab === "Intranet") {
        setMyApprovalsData(MyApprovaldata);
      } else if (tab === "DMS") {
        setMyApprovalsData(MyDMSAPPROVALDATA);
        setMylistdata(MyDMSAPPROVALDATA);
      } else if (tab === "ProjectWorkflow") {
        setMyApprovalsData(ProjectWorkflowData);
      } else if (tab === "Automation") {
        const sorted = [...Automationdata].sort(
          (a: any, b: any) => b.Created - a.Created,
        );
        setMyApprovalsData(sorted);
        setMyApprovalsDataAutomation(sorted);
      }
    } catch (error) {
      console.error("Tab click error:", error);
      setMyApprovalsData([]);
    } finally {
      setLoading(false);
      setStatusChange(false);
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
    try {
      const [MyApprovaldata, Automationdata1, ProjectWorkflowData] =
        await Promise.all([
          getMyApproval(sp, status),
          getApprovalListsData(sp, status),
          getProjectWorkflowApprovals(status),
        ]);

      let MyDMSAPPROVALDATA: any[] = [];

      if (activeTab === "DMS") {
        MyDMSAPPROVALDATA = await getApprovalmasterTasklist(status);
      }

      setMyApprovalsDataAll(MyApprovaldata);
      setMyApprovalsDataAutomation(
        Automationdata1.sort((a, b) => b.Created - a.Created),
      );
      setProjectWorkflowdata(ProjectWorkflowData);
      if (activeTab === "DMS") {
        setMylistdata(MyDMSAPPROVALDATA);
      }

      if (activeTab === "Intranet") {
        setMyApprovalsData(MyApprovaldata);
      } else if (activeTab === "Automation") {
        setMyApprovalsData(
          Automationdata1.sort((a, b) => b.Created - a.Created),
        );
      } else if (activeTab === "ProjectWorkflow") {
        setMyApprovalsData(ProjectWorkflowData);
      } else if (activeTab === "DMS") {
        setMyApprovalsData(MyDMSAPPROVALDATA); // ✅ Set DMS data
      }
    } catch (err) {
      console.error("ApiCall error:", err);
    } finally {
      setLoading(false); // ✅ Always stop loader
    }
  };

  const handleStatusChange = async (
    name: string,
    value: string,
    actingfor: any,
  ) => {
    if (value === "") return;
    setLoading(true);
    setStatusChange(true);
    actingforuseremail = actingfor;
    SetStatusvalue(name);

    try {
      const [MyApprovaldata, Automationdata, ProjectWorkflowData] =
        await Promise.all([
          getMyApproval(sp, value, actingfor),
          getApprovalListsData(sp, value, actingfor),
          getProjectWorkflowApprovals(value, actingfor),
        ]);

      let MyDMSAPPROVALDATA: any[] = [];

      if (activeTab === "DMS") {
        MyDMSAPPROVALDATA = await getApprovalmasterTasklist(value, actingfor);
      }
      setMyApprovalsDataAll(MyApprovaldata);
      setMyApprovalsDataAutomation(Automationdata);
      setProjectWorkflowdata(ProjectWorkflowData);

      if (activeTab === "DMS") {
        setMylistdata(MyDMSAPPROVALDATA);
      }

      if (activeTab === "Intranet") {
        setMyApprovalsData(MyApprovaldata);
      } else if (activeTab === "DMS") {
        setMyApprovalsData(MyDMSAPPROVALDATA);
      } else if (activeTab === "Automation") {
        const sorted = [...Automationdata].sort(
          (a, b) => b.Created - a.Created,
        );
        setMyApprovalsData(sorted);
      } else if (activeTab === "ProjectWorkflow") {
        setMyApprovalsData(ProjectWorkflowData);
      }
    } catch (error) {
      console.error("Status change error:", error);
    } finally {
      setLoading(false);
      setStatusChange(false);
    }
  };
  const handleFilterChange = (
    e: React.ChangeEvent<HTMLInputElement>,

    field: string,
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
    if (!data) return [];

    const filteredData = data.filter((item, index) => {
      let requestId = "",
        title = "",
        process = "",
        requestedBy = "",
        requestedDate = "",
        status = "";

      if (activeTab === "DMS") {
        requestId = item?.FileUID?.RequestNo || "";
        title = item?.FileUID?.FileName || "";
        process = item?.FileUID?.Processname || "";
        requestedBy = item?.RequestedByTitle || "";
        requestedDate = item?.FileUID?.Created
          ? new Date(item.FileUID.Created).toLocaleString()
          : "";
        status = item?.FileUID?.Status || "Pending";
      } else {
        requestId = item?.RequestID || "";
        title =
          activeTab === "Automation"
            ? item?.ApprovalTitle || ""
            : item?.Title || "";
        process = item?.ProcessName || "";
        requestedBy =
          activeTab === "Automation"
            ? item?.Author?.Title || ""
            : item?.Requester?.Title || "";
        requestedDate = item?.Created
          ? new Date(item.Created).toLocaleString()
          : "";
        status = item?.Status || "";
      }

      return (
        (filters.RequestID === "" ||
          requestId.toLowerCase().includes(filters.RequestID.toLowerCase())) &&
        (filters.Title === "" ||
          title.toLowerCase().includes(filters.Title.toLowerCase())) &&
        (filters.ProcessName === "" ||
          process.toLowerCase().includes(filters.ProcessName.toLowerCase())) &&
        (filters.RequestedBy === "" ||
          requestedBy
            .toLowerCase()
            .includes(filters.RequestedBy.toLowerCase())) &&
        (filters.RequestedDate === "" ||
          requestedDate
            .toLowerCase()
            .includes(filters.RequestedDate.toLowerCase())) &&
        (filters.Status === "" ||
          status.toLowerCase().includes(filters.Status.toLowerCase()))
      );
    });

    // Sorting with DMS field mapping
    let sortedData = [...filteredData];
    if (sortConfig.key) {
      sortedData.sort((a, b) => {
        let aVal = "",
          bVal = "";
        if (activeTab === "DMS") {
          if (sortConfig.key === "RequestID") {
            aVal = a?.FileUID?.RequestNo || "";
            bVal = b?.FileUID?.RequestNo || "";
          } else if (sortConfig.key === "Title") {
            aVal = a?.FileUID?.FileName || "";
            bVal = b?.FileUID?.FileName || "";
          } else if (sortConfig.key === "ProcessName") {
            aVal = a?.FileUID?.Processname || "";
            bVal = b?.FileUID?.Processname || "";
          } else if (sortConfig.key === "RequestedBy") {
            aVal = a?.RequestedByTitle || "";
            bVal = b?.RequestedByTitle || "";
          } else {
            aVal = a[sortConfig.key]
              ? String(a[sortConfig.key]).toLowerCase()
              : "";
            bVal = b[sortConfig.key]
              ? String(b[sortConfig.key]).toLowerCase()
              : "";
          }
        } else {
          aVal = a[sortConfig.key]
            ? String(a[sortConfig.key]).toLowerCase()
            : "";
          bVal = b[sortConfig.key]
            ? String(b[sortConfig.key]).toLowerCase()
            : "";
        }
        if (aVal < bVal) return sortConfig.direction === "ascending" ? -1 : 1;
        if (aVal > bVal) return sortConfig.direction === "ascending" ? 1 : -1;
        return 0;
      });
    }
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

      headerId: string,
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
    Item: any,
    mode: any,
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
              "&mode=" +
              mode +
              "&page=MyApproval";
            break;
          case "News":
            sessionkey = "announcementId";
            redirecturl =
              `${siteUrl}/SitePages/AddAnnouncement.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" +
              mode +
              "&page=MyApproval";
            break;
          case "Event":
            sessionkey = "EventId";
            redirecturl =
              `${siteUrl}/SitePages/EventMasterForm.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" +
              mode +
              "&page=MyApproval";
            break;
          case "Media":
            sessionkey = "mediaId";
            redirecturl =
              `${siteUrl}/SitePages/MediaGalleryForm.aspx` +
              "?requestid=" +
              Item?.Id +
              "&mode=" +
              mode +
              "&page=MyApproval";
            break;
          case "Blog":
            sessionkey = "blogId";
            redirecturl =
              `${siteUrl}/SitePages/BlogDetails.aspx?` +
              Item?.ContentId +
              "&page=MyApproval";
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
    Status: string,
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
      currentItem.Id,
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

  /**
   * below is all project workflow function added by priyanshu
   */

  const myActingfordata = async () => {
    try {
      const currentUserEmail = currentUserEmailRef.current;
      const today = new Date().toISOString();
      const delegateListItems = await sp.web.lists
        .getByTitle("DelegateList")
        .items.select(
          "DelegateName/EMail",
          "ActingFor/EMail",
          "ActingFor/Title",
          "DelegateName/Title",
          "StartDate",
          "EndDate",
          "Status",
        )
        .expand("DelegateName", "ActingFor")
        .filter(
          `ActingFor/EMail eq '${currentUserEmail}' and Status eq 'Active' and StartDate le '${today}' and EndDate ge '${today}'`,
        )();

      const uniqueTitlesAndEmails = [
        ...new Map(
          delegateListItems.map((item: any) => [
            item.DelegateName?.Title,
            {
              title: item.DelegateName?.Title,
              email: item.DelegateName?.EMail,
            },
          ]),
        ).values(),
      ];

      setSetActingForUser(
        uniqueTitlesAndEmails.map((item: any, index: number) => ({
          id: index.toString(),
          name: item.title,
          email: item.email,
        })),
      );
    } catch (error) {
      console.error("Error fetching data:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "myActingfordata",
        "MyApprovalWebPart",
      );
    }
  };

  const getProjectWorkflowApprovals = async (
    status: string,
    actingfor?: any,
  ) => {
    try {
      let currentUser;
      let currentUserId;

      if (actingfor) {
        // Get the acting user's ID by email
        const actingUser = await sp.web.siteUsers.getByEmail(actingfor)();
        currentUserId = actingUser.Id;
      } else {
        currentUser = await sp.web.currentUser();
        currentUserId = currentUser.Id;
      }

      // Remove .getAll() - use proper PnP JS syntax
      const items = await sp.web.lists
        .getByTitle("ProjectApprovals")
        .items.select(
          "*",
          "DeliverablesDetailsId/ID",
          "DeliverablesDetailsId/DocNumber",
          "ProjectCreationListID/ProjectName",
          "ProjectCreationListID/ID",
          "AssignedTo/ID",
          "AssignedTo/Title",
          "AssignedTo/EMail",
          "Author/Title",
          "Author/EMail",
        )
        .expand(
          "DeliverablesDetailsId",
          "ProjectCreationListID",
          "AssignedTo",
          "Author",
        )
        .filter(
          `AssignedTo/ID eq ${currentUserId} and Status eq '${status}' and ApproverRole ne 'Vendor'`,
        )
        .orderBy("Created", false)(); // Remove .getAll(), use () instead

      const approvals = await Promise.all(
        items.map(async (item: any, index: number) => {
          try {
            let creationItem: any = null;
            let DeliverablesItem: any = null;

            if (item.ProjectCreationListID?.ID) {
              creationItem = await sp.web.lists
                .getByTitle("ProjectCreationList")
                .items.getById(item.ProjectCreationListID.ID)
                .select(
                  "*",
                  "PreparedBy/Title",
                  "ProjectType/ProjectType",
                  "ProjectType/Id",
                  "ClientName",
                )
                .expand("ProjectType", "PreparedBy")();
            }

            if (item.DeliverablesDetailsId?.ID) {
              DeliverablesItem = await sp.web.lists
                .getByTitle("DeliverablesDetails")
                .items.getById(item.DeliverablesDetailsId.ID)
                .select(
                  "*",
                  "AssignedTo/ID",
                  "AssignedTo/Title",
                  "Organisation/Organisation",
                  "Area/Area",
                  "Deliverables/ID",
                  "Deliverables/Deliverables",
                )
                .expand("AssignedTo", "Organisation", "Area", "Deliverables")();
            }

            return {
              Id: item.Id,
              RequestID: item.DocNumber || `PROJ-${item.Id}`,
              Title:
                item.ProjectCreationListID?.ProjectName || "Project Approval",
              ProcessName: "Project Workflow",
              Status: item.Status,
              Requester: {
                Title: item.Author?.Title || "",
                EMail: item.Author?.EMail || "",
              },
              Author: { Title: item.Author?.Title || "" },
              Created: item.Created,
              InitiatedBy: creationItem?.AuthorId,
              VendorId: DeliverablesItem?.AssignedToId,
              ProjectName: item.ProjectCreationListID?.ProjectName || "",
              ProjectType: creationItem?.ProjectType?.ProjectType || "",
              ClientName: creationItem?.ClientName || "",
              PreparedBy: creationItem?.PreparedBy?.Title || "",
              Deliverable: DeliverablesItem?.Deliverables?.Deliverables || "",
              Area: DeliverablesItem?.Area?.Area || "",
              DocType: item.DocumentType || "",
              DocNumber: item.DeliverablesDetailsId?.DocNumber || "",
              DoYouNeedApproval: item.Doyouneedapproval || "",
              CurrentApprovalLevel: item.Level || "",
              ApprovalRole: item.ApproverRole || "",
              ApprovalSN: item.SerialNumber || 0,
              ApprovalCriteria: item.ApprovalCriteria || "",
              AssignedTo: item.AssignedTo?.Title || "",
              Org: DeliverablesItem?.Organisation?.Organisation || "N/A",
              RevisionNumber: item.RevisionNumber || "0",
              DocumentNumber: item.DocNumber || "",
              ProjectDate: item.Created
                ? new Date(item.Created).toLocaleDateString("en-GB")
                : "",
              ProjectId: item.ProjectCreationListID?.ID,
              DeliverableId: item.DeliverablesDetailsId?.ID,
              Remarks: item.Remarks || "",
              SNo: index + 1,
            };
          } catch (error) {
            console.error(`Error fetching creation item:`, error);
            await ErrorLogger.logError(
              sp,
              error,
              "getProjectWorkflowApprovals.item",
              "MyApprovalWebPart",
            );
            return {
              Id: item.Id,
              RequestID: item.DocNumber || `PROJ-${item.Id}`,
              Title:
                item.ProjectCreationListID?.ProjectName || "Project Approval",
              ProcessName: "Project Workflow",
              Status: item.Status,
              Requester: {
                Title: item.Author?.Title || "",
                EMail: item.Author?.EMail || "",
              },
              Author: { Title: item.Author?.Title || "" },
              Created: item.Created,
              ProjectName: item.ProjectCreationListID?.ProjectName || "",
              ProjectType: "",
              ClientName: "",
              PreparedBy: "",
              Deliverable: "",
              Area: "",
              DocType: item.DocumentType || "",
              DocNumber: item.DeliverablesDetailsId?.DocNumber || "",
              DoYouNeedApproval: item.Doyouneedapproval || "",
              CurrentApprovalLevel: item.Level || "",
              ApprovalRole: item.ApproverRole || "",
              ApprovalSN: item.SerialNumber || 0,
              ApprovalCriteria: item.ApprovalCriteria || "",
              AssignedTo: item.AssignedTo?.Title || "",
              Org: "",
              RevisionNumber: item.RevisionNumber || "0",
              DocumentNumber: item.DocNumber || "",
              ProjectDate: item.Created
                ? new Date(item.Created).toLocaleDateString("en-GB")
                : "",
              ProjectId: item.ProjectCreationListID?.ID,
              DeliverableId: item.DeliverablesDetailsId?.ID,
              Remarks: item.Remarks || "",
              SNo: index + 1,
            };
          }
        }),
      );

      return approvals;
    } catch (error) {
      console.error("Error fetching Project Workflow items:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "getProjectWorkflowApprovals",
        "MyApprovalWebPart",
      );
      return [];
    }
  };

  const addNewProjectApprovalRow = (item?: any) => {
    const newRow: ApprovalHierarchyItem = {
      level: `Level ${projectHierarchy.length + 1}`,
      approverRole: item?.ApproverRole || "",
      approver: item?.AssignedTo?.Title || "",
      approvers: item?.AssignedTo ? [item.AssignedTo.Title] : [],
      approvalCriteria: item?.ApprovalCriteria || "Anyone",
      assignedTo: item?.AssignedTo ? [item.AssignedTo] : [],
    };
    setProjectHierarchy((prev) => [...prev, newRow]);
  };

  const deleteProjectApprovalRow = async (index: number) => {
    // Maintain at least one approval level
    if (projectHierarchy.length <= 1) {
      setPopupType("validation");
      setPopupTitle("Cannot Delete");
      setPopupMessage(
        "At least one approval level is required in the hierarchy.",
      );
      setPopupOpen(true);
      return;
    }

    const rowToDelete = projectHierarchy[index];

    // ===== REVISION CHECK: DELETE ONLY IF REVISION > 0 =====
    // Handle both string and number formats
    const revisionValue = selectedProjectTask?.RevisionNumber;
    let currentRevision = 0;

    if (typeof revisionValue === "string") {
      currentRevision = parseInt(revisionValue, 10);
    } else if (typeof revisionValue === "number") {
      currentRevision = revisionValue;
    }

    // Also check if revision exists in selectedProjectTask
    if (!revisionValue && selectedProjectTask?.RevisionNumber) {
      currentRevision = parseInt(
        selectedProjectTask.RevisionNumber.toString(),
        10,
      );
    }

    console.log(
      `Delete row - Current Revision: ${currentRevision}, Row ID: ${rowToDelete.id}`,
    );

    if (isNaN(currentRevision)) {
      currentRevision = 0;
    }

    if (currentRevision > 0) {
      // DELETE FROM SHAREPOINT LIST ONLY IF REVISION > 0
      if (rowToDelete.id) {
        try {
          await sp.web.lists
            .getByTitle("ApprovalHierarchy")
            .items.getById(rowToDelete.id)
            .delete();
          console.log(
            `✅ Deleted hierarchy row with ID: ${rowToDelete.id} from SharePoint (Revision: ${currentRevision})`,
          );
        } catch (error) {
          console.error("Error deleting hierarchy row from SharePoint:", error);
          await ErrorLogger.logError(
            sp,
            error,
            "deleteProjectApprovalRow",
            "MyApprovalWebPart",
          );
          // Stop execution if delete fails
          return;
        }
      }
    } else {
      console.log(
        `ℹ️ Revision is ${currentRevision}, skipping SharePoint delete`,
      );
    }

    // Update UI state after successful delete
    setProjectHierarchy((prev) => {
      const updatedHierarchy = prev.filter((_, i) => i !== index);
      return updatedHierarchy.map((row, i) => ({
        ...row,
        level: `Level ${i + 1}`,
      }));
    });
  };
  const updateProjectApprovalRow = (
    index: number,
    field: keyof ApprovalHierarchyItem,
    value: string,
  ) => {
    setProjectHierarchy((prev) =>
      prev.map((row, i) => (i === index ? { ...row, [field]: value } : row)),
    );
  };
  // Add this function to clear old approval tasks
  const clearOldProjectApprovalTasks = async () => {
    if (!selectedProjectTask) return;

    try {
      // Get all existing tasks for this deliverable
      const existingTasks = await sp.web.lists
        .getByTitle("ProjectApprovals")
        .items.select("Id")
        .filter(
          `ProjectCreationListID/ID eq ${selectedProjectTask.ProjectId} and DeliverablesDetailsId/ID eq ${selectedProjectTask.DeliverableId} and Status eq 'Pending'`,
        )
        .expand("ProjectCreationListID", "DeliverablesDetailsId")();

      // Delete all pending tasks
      for (const task of existingTasks) {
        if (task && task.Id) {
          await sp.web.lists
            .getByTitle("ProjectApprovals")
            .items.getById(task.Id)
            .delete();
          console.log(`Deleted old task ID: ${task.Id}`);
        }
      }
    } catch (error) {
      console.error("Error clearing old tasks:", error);
    }
  };
  const handleProjectApproverChange = (index: number, selectedOptions: any) => {
    setProjectHierarchy((prev) =>
      prev.map((row, i) => {
        if (i === index) {
          const selectedUsers = selectedOptions || [];
          const approverNames = selectedUsers
            .map((user: any) => user.label)
            .join(", ");
          const assignedToArray = selectedUsers.map((user: any) => ({
            ID: parseInt(user.value),
            Title: user.label,
            EMail: user.email,
          }));
          return {
            ...row,
            approver: approverNames,
            assignedTo: assignedToArray,
          };
        }
        return row;
      }),
    );
  };

  const getCurrrentuser = async () => {
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    myActingfordata();
  };

  /**
   * Fetch and cache all SharePoint site users
   * Used for approver assignment dropdowns
   */
  const fetchUsers = async () => {
    try {
      const siteUsers = await sp.web.siteUsers();
      const userOptions: UserOption[] = siteUsers
        .filter((user: any) => user.Email && user.Title)
        .map((user: any) => ({
          value: user.Id.toString(),
          label: user.Title,
          email: user.Email,
        }));
      setUsers(userOptions);
    } catch (error) {
      console.error("Error fetching users:", error);
      await ErrorLogger.logError(sp, error, "fetchUsers", "MyApprovalWebPart");
      setUsers([
        {
          value: "1",
          label: "Current User",
          email: props.context.pageContext.user.email,
        },
        { value: "2", label: "Admin User", email: "admin@contoso.com" },
      ]);
    }
  };

  const isCurrentUserDocumentController = (docType: string): boolean => {
    const config = getConfigurationByDocType(docType);
    return currentUserId === config.documentControllerId;
  };

  React.useEffect(() => {
    getCurrrentuser();
    fetchUsers();
    fetchApproverRoles();
    setCurrentUserId(props.context.pageContext.legacyPageContext.userId);
  }, []);

  const getProjectConfiguration = async (): Promise<{
    documentControllerId: number | null;
    dccId: number | null;
  }> => {
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
          "DCC/EMail",
          "Consolidator/ID",
          "Consolidator/Title",
          "Consolidator/EMail",
          "DocumentType/ID",
          "DocumentType/DocumentType",
        )
        .expand("DocumentController", "DCC", "Consolidator", "DocumentType")
        .orderBy("Created", false)();

      // Store all configuration records for dynamic filtering
      console.log("All ProjectConfiguration Data:", items);
      setProjectConfigurationsData(items);

      // For backward compatibility, return first record if available
      if (items.length > 0) {
        return {
          documentControllerId: items[0]?.DocumentControllerId || null,
          dccId: items[0]?.DCCId || null,
        };
      }
      return { documentControllerId: null, dccId: null };
    } catch (error) {
      console.error("Error fetching Project Configuration:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "getProjectConfiguration",
        "MyApprovalWebPart",
      );
      return { documentControllerId: null, dccId: null };
    }
  };

  const getConfigurationByDocType = (
    docType: string,
  ): {
    documentControllerId: number | null;
    dccId: number | null;
  } => {
    console.log(`Filtering configuration for DocumentType: ${docType}`);
    const config = projectConfigurationsData.find(
      (item) =>
        item.DocumentType?.DocumentType === docType ||
        item.DocumentType === docType,
    );

    if (config) {
      console.log(`Found configuration for ${docType}:`, {
        documentControllerId: config.DocumentControllerId,
        dccId: config.DCCId,
      });
      return {
        documentControllerId: config.DocumentControllerId || null,
        dccId: config.DCCId || null,
      };
    }

    console.log(
      `No configuration found for ${docType}, using default (first record)`,
    );
    return {
      documentControllerId:
        projectConfigurationsData[0]?.DocumentControllerId || null,
      dccId: projectConfigurationsData[0]?.DCCId || null,
    };
  };

  React.useEffect(() => {
    const initializeData = async () => {
      const { documentControllerId: dcId, dccId: dId } =
        await getProjectConfiguration();
      setDocumentControllerId(dcId);
      setDccId(dId);
    };
    initializeData();
  }, []);

  const handleProjectWorkflowAction = async (e: any, item: any) => {
    e.preventDefault();
    await handleProjectViewClick(item);
  };

  const handleProjectViewClick = async (task: any) => {
    // Dynamically get configuration based on task's document type
    const docTypeConfig = getConfigurationByDocType(task.DocType);
    setDocumentControllerId(docTypeConfig.documentControllerId);
    setDccId(docTypeConfig.dccId);

    console.log(
      `Loaded configuration for task with DocType: ${task.DocType}`,
      docTypeConfig,
    );

    setSelectedProjectTask(task);
    setShowProjectForm(true);
    setProjectNeedsFurtherApproval(task.DoYouNeedApproval || "Select");
    setProjectRemarks(task.Remarks || "");

    setIsReadOnly(task.Status !== "Pending");

    if (task.DeliverableId) {
      const docInfo = await fetchDocumentForDeliverable(
        task.DeliverableId,
        task.RevisionNumber,
      );
      setProjectDocumentInfo(docInfo);
    }
    await getDocumentComments(
      task.ProjectId,
      task.DeliverableId,
      task.RevisionNumber,
    );
    if (task.ProjectId && task.DeliverableId) {
      await getProjectApprovalHierarchy(
        task.ProjectId,
        task.DeliverableId,
        task.DocType,
      );
    } else {
      addNewProjectApprovalRow();
    }
  };

  const getProjectApprovalHierarchy = async (
    projectId: number,
    deliverableId: number,
    documentType: string,
  ) => {
    try {
      const items = await sp.web.lists
        .getByTitle("ApprovalHierarchy")
        .items.select(
          "*,DeliverablesDetailsId/ID,ProjectCreationListID/ID,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail",
        )
        .expand("DeliverablesDetailsId,ProjectCreationListID,AssignedTo")
        .filter(
          `ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId}`,
        )
        .orderBy("SerialNumber", true)();

      if (items.length > 0) {
        const hierarchyItems: ApprovalHierarchyItem[] = items.map(
          (item: any, index: number) => {
            const assignedArray = Array.isArray(item.AssignedTo)
              ? item.AssignedTo
              : [];
            return {
              id: item.Id,
              level: `Level ${index + 1}`,
              approverRole: item.ApproverRole || "",
              approver: assignedArray.map((a: any) => a.Title).join(", "),
              approvers: assignedArray.map((a: any) => a.Title),
              approvalCriteria: item.ApprovalCriteria || "Anyone",
              serialNumber: item.SerialNumber,
              assignedTo: assignedArray.map((a: any) => ({
                ID: a.ID,
                Title: a.Title,
                EMail: a.EMail,
              })),
            };
          },
        );
        setProjectHierarchy(hierarchyItems);
      } else {
        await getProjectWorkflowConfiguration(documentType);
      }
    } catch (error) {
      console.error("Error fetching approval hierarchy:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "getProjectApprovalHierarchy",
        "MyApprovalWebPart",
      );
      await getProjectWorkflowConfiguration(documentType);
    }
  };

  const getProjectWorkflowConfiguration = async (docType: string) => {
    try {
      const items = await sp.web.lists
        .getByTitle("ProjectWorkflowConfigurationMaster")
        .items.select(
          "*,DocumentType/ID,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail",
        )
        .expand("DocumentType,AssignedTo")
        .filter(`DocumentType/DocumentType eq '${docType}'`)
        .orderBy("ID", true)();

      if (items.length > 0) {
        const hierarchyItems: ApprovalHierarchyItem[] = items.map(
          (item: any, index: number) => ({
            level: `Level ${index + 1}`,
            approverRole: item.Role || "",
            approver: item.AssignedTo?.Title || "",
            approvers: item?.AssignedTo ? [item.AssignedTo.Title] : [],
            approvalCriteria: item.ApprovalCriteria || "Anyone",
            assignedTo: item.AssignedTo ? [item.AssignedTo] : [],
          }),
        );
        setProjectHierarchy(hierarchyItems);
      } else {
        addNewProjectApprovalRow();
      }
    } catch (error) {
      console.error("Error fetching workflow configuration:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "getProjectWorkflowConfiguration",
        "MyApprovalWebPart",
      );
      addNewProjectApprovalRow();
    }
  };

  const handleProjectBackClick = () => {
    setSelectedProjectTask(null);
    setShowProjectForm(false);
    setProjectHierarchy([]);
    setProjectNeedsFurtherApproval("Select");
    setProjectRemarks("");
    setProjectDocumentInfo([]);
  };

  const fetchDocumentForDeliverable = async (
    deliverableId: number,
    revisionNo: number,
  ) => {
    try {
      const documents = await sp.web.lists
        .getByTitle("DeliverablesDocument")
        .items.select(
          "Id,Revision,File/ServerRelativeUrl,FileLeafRef,FileRef,Author/ID,Author/Title,SharedLink,IsConsolidatorCopy,IsMyTask,ProjectID,DeliverablesDetailsId",
        )
        .expand("File", "Author")
        .filter(
          `DeliverablesDetailsId eq ${deliverableId} and Revision eq '${revisionNo.toString()}'`,
        )
        .orderBy("ID", false)();

      if (documents.length > 0) {
        return documents.map((doc: any) => ({
          documentUrl: doc.File?.ServerRelativeUrl,
          fileName: doc.FileLeafRef,
          fileLeafRef: doc.FileLeafRef,
          fileRef: doc.FileRef,
          sharedLink: doc.SharedLink,
          author: doc.Author?.Title,
          authorId: doc.Author?.ID,
          id: doc.Id,
          isConsolidatorCopy: doc.IsConsolidatorCopy,
          isMyTask: doc.IsMyTask,
        }));
      }
      return [];
    } catch (error) {
      console.error(`Error fetching documents:`, error);
      await ErrorLogger.logError(
        sp,
        error,
        "fetchDocumentForDeliverable",
        "MyApprovalWebPart",
      );
      return [];
    }
  };

  const getDocumentComments = async (
    projectCreationID: number,
    deliverableDetailsID: number,
    revision: string,
  ) => {
    try {
      const items = await sp.web.lists
        .getByTitle("DocumentComments")
        .items.select("*", "IsDeleted")
        .filter(
          `ProjectID eq ${projectCreationID} and DeliverablesDetailsId eq ${deliverableDetailsID}`,
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
          isDeleted: item.IsDeleted === true || item.IsDeleted === "Yes",
        }));
        const activeComments = comments.filter((comment) => !comment.isDeleted);
        setAllDocumentComments(activeComments);
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
      } else {
        setShowDocumentComments(false);
        setDocumentComments([]);
        setAllDocumentComments([]);
        setVersionList([]);
      }
    } catch (error) {
      console.error("Error fetching document comments:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "getDocumentComments",
        "MyApprovalWebPart",
      );
      setShowDocumentComments(false);
    }
  };

  const onVersionChange = (version: string) => {
    setSelectedVersion(version);
    if (!version) {
      setDocumentComments(allDocumentComments);
    } else {
      const filteredComments = allDocumentComments.filter(
        (comment) => comment.revision === version,
      );
      setDocumentComments(filteredComments);
    }
  };

  const exportCommentsToExcel = () => {
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
    a.download = `DocumentComments_${selectedProjectTask?.DocNumber || "export"}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    window.URL.revokeObjectURL(url);
  };

  const refreshDocComment = () => {
    if (selectedProjectTask) {
      getDocumentComments(
        selectedProjectTask.ProjectId,
        selectedProjectTask.DeliverableId,
        selectedProjectTask.RevisionNumber,
      );
    }
  };

  const handleProjectItemSelect = async (id: number) => {
    const alreadySelected = selectedProjectItems.some((item) => item.Id === id);
    if (alreadySelected) {
      setSelectedProjectItems((prev) => prev.filter((item) => item.Id !== id));
      return;
    }

    const details = currentData?.find((item) => item.Id === id);
    if (!details) return;

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
          "ClientName",
        )
        .expand("ProjectType", "PreparedBy")();
    }

    let DeliverablesItem: any = null;
    if (details.DeliverableId) {
      DeliverablesItem = await sp.web.lists
        .getByTitle("DeliverablesDetails")
        .items.getById(details.DeliverableId)
        .select(
          "*",
          "AssignedTo/ID",
          "AssignedTo/Title",
          "Organisation/Organisation",
          "Area/Area",
          "Deliverables/ID",
          "Deliverables/Deliverables",
        )
        .expand("AssignedTo", "Organisation", "Area", "Deliverables")();
    }

    const fullDetails: ProjectItem = {
      Id: details.Id,
      Status: details.Status,
      Author: { Title: details.Author?.Title || "" },
      Created: details.Created,
      InitiatedBy: creationItem?.AuthorId ?? creationItem?.AuthorId,
      documentNumber: DeliverablesItem?.DocNumber ?? "",
      VendorName: DeliverablesItem?.AssignedTo?.Title ?? "",
      docrevisionNumber: DeliverablesItem?.RevisionNumber ?? "",
      VendorId: DeliverablesItem?.AssignedToId ?? 0,
      ProjectName: creationItem?.ProjectName || "",
      deliverablesDocumentIDs: DeliverablesItem?.DeliverablesDocumentIDId ?? [],
      ProjectId: creationItem?.ID,
      DeliverableId: DeliverablesItem?.ID,
      PreparedBy: creationItem?.PreparedBy?.Title || "",
      PreparedById: creationItem?.PreparedBy?.ID || 0,
      DocType: DeliverablesItem?.DocumentType || "",
      ApprovalRole: details.ApprovalRole || "",
      RevisionNumber: details.RevisionNumber || "0",
      Deliverable:
        DeliverablesItem?.Deliverables?.Deliverables ||
        details?.Deliverable ||
        "",
    };

    setSelectedProjectItems((prev) => [...prev, fullDetails]);
  };

  const handlebulkReworkAction = async () => {
    if (selectedProjectItems.length === 0) {
      alert("Please select at least one item before performing Rework.");
      return;
    }
    setShowSubmitLoader(true);
    try {
      await processVendorRework();
      setPopupType("success");
      setPopupTitle("Bulk Rework Success");
      setPopupMessage(
        `Successfully processed rework for ${selectedProjectItems.length} deliverable(s).`,
      );
      setPopupOpen(true);
    } catch (error) {
      console.error("Error in bulk rework:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "handlebulkReworkAction",
        "MyApprovalWebPart",
      );
      setPopupType("error");
      setPopupTitle("Error");
      setPopupMessage("Failed to process bulk rework. Please try again.");
      setPopupOpen(true);
    } finally {
      setShowSubmitLoader(false);
    }
  };

  const processVendorRework = async () => {
    if (!selectedProjectItems || selectedProjectItems.length === 0) return;
    const currentUserId = props.context.pageContext.legacyPageContext.userId;

    for (const task of selectedProjectItems) {
      try {
        await sp.web.lists
          .getByTitle("DeliverablesDetails")
          .items.getById(task.DeliverableId!)
          .update({ Status: "Pending", IsReworked: "Yes" });

        if (task.ApprovalRole === "Document Controller") {
          const approvalUpdate: any = {
            Status: "Rework",
            Remarks: projectRemarks,
            ApprovalDate: new Date(),
            OutgoingDate: new Date(),
            OutgoingStatus: outgoingStatus,
          };
          if (projectNeedsFurtherApproval !== "Select")
            approvalUpdate.Doyouneedapproval = projectNeedsFurtherApproval;
          await sp.web.lists
            .getByTitle("ProjectApprovals")
            .items.getById(task.Id)
            .update(approvalUpdate);
        }

        await sp.web.lists.getByTitle("ProjectApprovals").items.add({
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
        });

        const fileDetails: {
          name: string;
          link: string;
          docrevisionNo: string;
          id: string;
        }[] = [];

        const CCUserName = users.filter(
          (u) => String(u.value) === String(task.InitiatedBy),
        );
        const CCName = CCUserName.length > 0 ? CCUserName[0].label : null;

        const emailBody = buildbulkApprovalEmailBody(
          "REWORK_BY_DOCUMENT_CONTROLLER",
          task.documentNumber,
          task.VendorName,
          CCName,
          task.ProjectName,
          fileDetails.map((d) => ({
            deliverable: task?.Deliverable,
            fileName: d.name,
            sharedLink: d.link,
            docrevisionNo: d.docrevisionNo,
            id: Number(d.id),
          })),
        );

        await createbulkEmailTriggerDetails(
          task.VendorId,
          [task.InitiatedBy],
          emailBody,
          `Rework Action Assigned – ${task.Deliverable}`,
          task.deliverablesDocumentIDs,
          task.Deliverable,
        );
      } catch (err) {
        console.error("Error processing vendor rework:", err);
        await ErrorLogger.logError(
          sp,
          err,
          "processVendorRework",
          "MyApprovalWebPart",
        );
      }
    }
  };

  const handleProjectApprovalAction = async (
    status: "Approved" | "Rejected" | "Rework",
  ) => {
    if (!validateApprovalAction(status)) return;
    setShowSubmitLoader(true);

    if (!selectedProjectTask) {
      setShowSubmitLoader(false);
      return;
    }

    try {
      // Log action details including filtered configuration
      console.log(`Processing action: ${status}`);
      console.log(
        `Selected task ApprovalRole: ${selectedProjectTask.ApprovalRole}`,
      );
      console.log(`Selected task DocType: ${selectedProjectTask.DocType}`);
      const docTypeConfig = getConfigurationByDocType(
        selectedProjectTask.DocType,
      );
      console.log(
        `Filtered configuration for ${selectedProjectTask.DocType}:`,
        {
          DocumentControllerId: docTypeConfig.documentControllerId,
          DCCId: docTypeConfig.dccId,
        },
      );

      await updateCurrentApprovalStatus(status);
      let actionExecuted = false;

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

      if (actionExecuted) {
        if (selectedProjectTask?.ApprovalRole === "DCC") {
          // DCC success - reload page after a short delay
          setShowSubmitLoader(false);
          setPopupType("success");
          setPopupTitle("Success");
          setPopupMessage("Document submitted successfully");
          setPopupOpen(true);
        } else {
          setPopupType("success");
          setPopupTitle("Success");
          setPopupMessage(`Request ${status} successfully.`);
          setPopupOpen(true);
        }
      }
    } catch (error) {
      console.error("Error in approval action:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "handleProjectApprovalAction",
        "MyApprovalWebPart",
      );
      setShowSubmitLoader(false);
      setPopupType("error");
      setPopupTitle("Error");
      setPopupMessage(
        "An error occurred while processing your request. Please try again.",
      );
      setPopupOpen(true);
    } finally {
      if (selectedProjectTask?.ApprovalRole !== "DCC") {
        setShowSubmitLoader(false);
      }
    }
  };

  const validateApprovalAction = (
    status: "Approved" | "Rejected" | "Rework",
  ): boolean => {
    const errors: { [key: string]: boolean } = {};
    let isValid = true;
    let errorMessages: string[] = [];

    if (!selectedProjectTask) {
      setPopupType("validation");
      setPopupTitle("Validation Error");
      setPopupMessage("No project task selected.");
      setPopupOpen(true);
      return false;
    }

    // Document Controller specific validation for Approve
    if (
      status === "Approved" &&
      selectedProjectTask.ApprovalRole === "Document Controller"
    ) {
      if (
        !projectNeedsFurtherApproval ||
        projectNeedsFurtherApproval === "Select"
      ) {
        errors.needsFurtherApproval = true;
        errorMessages.push("Do you need further approval?");
        isValid = false;
      }

      // Validate hierarchy if further approval is needed
      if (projectNeedsFurtherApproval === "Yes") {
        if (projectHierarchy.length === 0) {
          errors.hierarchy = true;
          errorMessages.push("At least one approval level is required");
          isValid = false;
        } else {
          // Validate each hierarchy row
          projectHierarchy.forEach((row, index) => {
            if (!row.approverRole || row.approverRole.trim() === "") {
              errors[`hierarchy_role_${index}`] = true;
              errorMessages.push(`Role in hierarchy row ${index + 1}`);
              isValid = false;
            }
            if (
              !row.approver ||
              row.approver.trim() === "" ||
              !row.assignedTo ||
              row.assignedTo.length === 0
            ) {
              errors[`hierarchy_approver_${index}`] = true;
              errorMessages.push(`Approver in hierarchy row ${index + 1}`);
              isValid = false;
            }
            if (!row.approvalCriteria || row.approvalCriteria.trim() === "") {
              errors[`hierarchy_criteria_${index}`] = true;
              errorMessages.push(
                `Approval Criteria in hierarchy row ${index + 1}`,
              );
              isValid = false;
            }
          });
        }
      }
    }

    // DCC specific validation
    if (selectedProjectTask.ApprovalRole === "DCC") {
      if (
        !projectWantsToPublishInDossier ||
        projectWantsToPublishInDossier === "Select"
      ) {
        errors.wantsToPublishInDossier = true;
        errorMessages.push("Want to publish in Dossier?");
        isValid = false;
      }
      if (!projectCodestatus || projectCodestatus === "Select") {
        errors.codeStatus = true;
        errorMessages.push("Code status");
        isValid = false;
      }
    }

    // Document Controller specific validation for Rework
    if (
      status === "Rework" &&
      selectedProjectTask.ApprovalRole === "Document Controller"
    ) {
      if (!projectOutgoingstatus || projectOutgoingstatus === "Select") {
        errors.outgoingStatus = true;
        errorMessages.push("Outgoing status");
        isValid = false;
      }
    }

    // Rework and Rejection always require remarks
    if (status === "Rework" || status === "Rejected") {
      if (!projectRemarks || projectRemarks.trim() === "") {
        errors.remarks = true;
        errorMessages.push("Remarks");
        isValid = false;
      }
    }

    setFieldErrors(errors);

    if (!isValid) {
      const fieldList = errorMessages
        .map((field) => `• ${field}`)
        .join("<br/>");
      setPopupType("validation");
      setPopupTitle("Validation Error");
      setPopupMessage(`Please fill all mandatory fields`);
      setPopupOpen(true);
    }

    return isValid;
  };

  const updateCurrentApprovalStatus = async (status: string) => {
    if (!selectedProjectTask) return;
    try {
      const data: any = {
        Status: status,
        Remarks: projectRemarks,
        ApprovalDate: new Date(),
      };
      if (projectNeedsFurtherApproval !== "Select")
        data.Doyouneedapproval = projectNeedsFurtherApproval;
      if (
        status === "Rework" &&
        selectedProjectTask.ApprovalRole === "Document Controller"
      ) {
        data.OutgoingDate = new Date();
        data.OutgoingStatus = projectOutgoingstatus;
      }
      await sp.web.lists
        .getByTitle("ProjectApprovals")
        .items.getById(selectedProjectTask.Id)
        .update(data);
    } catch (error) {
      console.error("Error in updateCurrentApprovalStatus:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "updateCurrentApprovalStatus",
        "MyApprovalWebPart",
      );
      throw error;
    }
  };

  const handleApprovedAction = async () => {
    if (!selectedProjectTask) return;
    const approvals = await getAllProjectApproval(
      selectedProjectTask.ProjectId!,
      selectedProjectTask.DeliverableId!,
    );
    await addMultipleApproval("Approved", approvals);
  };

  const getAllProjectApproval = async (
    projectId: number,
    deliverableId: number,
  ) => {
    try {
      return await sp.web.lists
        .getByTitle("ProjectApprovals")
        .items.select(
          "*,DeliverablesDetailsId/ID,ProjectCreationListID/ID,AssignedTo/ID,AssignedTo/Title",
        )
        .expand("DeliverablesDetailsId,ProjectCreationListID,AssignedTo")
        .filter(
          `ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId}`,
        )
        .orderBy("Created", false)();
    } catch (error) {
      console.error("Error fetching project approvals:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "getAllProjectApproval",
        "MyApprovalWebPart",
      );
      return [];
    }
  };

  const addMultipleApproval = async (
    statusUpdate: string,
    allApprovals: any[],
  ) => {
    if (!selectedProjectTask) return;
    try {
      if (statusUpdate === "Approved") {
        const currentRole = selectedProjectTask.ApprovalRole;
        const currentCriteria = selectedProjectTask.ApprovalCriteria;

        if (
          currentRole === "Document Controller" &&
          projectNeedsFurtherApproval === "Yes"
        ) {
          // ===== SIMPLE APPROACH: DELETE ALL BY DELIVERABLE ID =====
          const deliverableId = selectedProjectTask.DeliverableId;
          const projectId = selectedProjectTask.ProjectId;

          console.log(
            `Deleting all records for Deliverable ID: ${deliverableId}, Project ID: ${projectId}`,
          );

          // ===== STEP 1: DELETE ALL FROM APPROVALHIERARCHY =====
          try {
            const hierarchyItems = await sp.web.lists
              .getByTitle("ApprovalHierarchy")
              .items.select(
                "Id",
                "ProjectCreationListID/ID",
                "DeliverablesDetailsId/ID",
              )
              .filter(
                `ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId}`,
              )
              .expand("ProjectCreationListID", "DeliverablesDetailsId")();

            for (const item of hierarchyItems) {
              if (item && item.Id) {
                await sp.web.lists
                  .getByTitle("ApprovalHierarchy")
                  .items.getById(item.Id)
                  .delete();
                console.log(`Deleted hierarchy entry: ${item.Id}`);
              }
            }
          } catch (error) {
            console.error("Error deleting hierarchy entries:", error);
          }

          // ===== STEP 2: DELETE ALL FROM PROJECTAPPROVALS =====
          try {
            const approvalItems = await sp.web.lists
              .getByTitle("ProjectApprovals")
              .items.select(
                "Id",
                "ProjectCreationListID/ID",
                "DeliverablesDetailsId/ID",
              )
              .filter(
                `ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId} and Status eq 'Pending'`,
              )
              .expand("ProjectCreationListID", "DeliverablesDetailsId")();

            for (const item of approvalItems) {
              if (item && item.Id) {
                await sp.web.lists
                  .getByTitle("ProjectApprovals")
                  .items.getById(item.Id)
                  .delete();
                console.log(`Deleted approval entry: ${item.Id}`);
              }
            }
          } catch (error) {
            console.error("Error deleting approval entries:", error);
          }

          // ===== STEP 3: RESET DELIVERABLE STATUS =====
          await sp.web.lists
            .getByTitle("DeliverablesDetails")
            .items.getById(deliverableId)
            .update({ Status: "Pending" });

          // ===== STEP 4: ADD NEW HIERARCHY ENTRIES =====
          for (let index = 0; index < projectHierarchy.length; index++) {
            const row = projectHierarchy[index];
            const approverAssignOwner =
              row.assignedTo?.map((user: any) => user.ID) || [];

            const rowData = {
              ProjectCreationListIDId: projectId,
              DeliverablesDetailsIdId: deliverableId,
              AssignedToId: approverAssignOwner,
              RequestedById: props.context.pageContext.legacyPageContext.userId,
              RequestedDate: new Date(),
              DocumentType: selectedProjectTask.DocType,
              ApproverRole: row.approverRole,
              ApprovalCriteria: row.approvalCriteria,
              Status: "Pending",
              Doyouneedapproval: projectNeedsFurtherApproval,
              Level: row.level,
              SerialNumber: index + 1,
            };

            try {
              const newItem = await sp.web.lists
                .getByTitle("ApprovalHierarchy")
                .items.add(rowData);

              const newItemId = newItem?.Id || newItem?.data?.Id || newItem?.ID;
              console.log(`Added new hierarchy row with ID: ${newItemId}`);
            } catch (addError) {
              console.error(`Error adding hierarchy row ${index}:`, addError);
            }
          }

          // ===== STEP 5: CREATE NEW TASKS FOR LEVEL 1 =====
          for (let index = 0; index < projectHierarchy.length; index++) {
            const row = projectHierarchy[index];
            const approverAssignOwner =
              row.assignedTo?.map((user: any) => user.ID) || [];

            if (row.level === "Level 1" && approverAssignOwner.length > 0) {
              for (const userId of approverAssignOwner) {
                try {
                  await sp.web.lists.getByTitle("ProjectApprovals").items.add({
                    ProjectCreationListIDId: projectId,
                    DeliverablesDetailsIdId: deliverableId,
                    AssignedToId: userId,
                    RequestedById:
                      props.context.pageContext.legacyPageContext.userId,
                    RequestedDate: new Date(),
                    DocumentType: selectedProjectTask.DocType,
                    ApproverRole: row.approverRole,
                    ApprovalCriteria: row.approvalCriteria,
                    Doyouneedapproval: projectNeedsFurtherApproval,
                    Status: "Pending",
                    Level: row.level,
                    SerialNumber: index + 1,
                    RevisionNumber: selectedProjectTask.RevisionNumber,
                  });
                  console.log(`Created new task for user ID: ${userId}`);

                  // Send email notification
                  const TouserUsers = users.filter(
                    (u) => String(u.value) === String(userId),
                  );
                  const TouserName =
                    TouserUsers.length > 0 ? TouserUsers[0].label : null;
                  const CCUserName = users.filter(
                    (u) =>
                      String(u.value) ===
                      String(selectedProjectTask.InitiatedBy),
                  );
                  const CCName =
                    CCUserName.length > 0 ? CCUserName[0].label : null;

                  const emailBody = buildApprovalEmailBody(
                    "APPROVAL_ASSIGNED",
                    TouserName,
                    CCName,
                    projectDocumentInfo.map((d) => ({
                      deliverable: selectedProjectTask.Deliverable,
                      fileName: d.fileName,
                      sharedLink: d.sharedLink,
                      id: d.id,
                    })),
                  );
                  await createEmailTriggerDetails(
                    userId,
                    [selectedProjectTask.InitiatedBy],
                    emailBody,
                    `${actionConfig["APPROVAL_ASSIGNED"].subject} – ${selectedProjectTask.Deliverable}`,
                    projectDocumentInfo.map((d) => d.id),
                  );
                } catch (error) {
                  console.error(
                    `Error creating task for user ${userId}:`,
                    error,
                  );
                }
              }
            }
          }
        } else if (
          currentRole === "Document Controller" &&
          projectNeedsFurtherApproval === "No"
        ) {
          // ===== SIMPLE APPROACH: DELETE ALL BY DELIVERABLE ID =====
          const deliverableId = selectedProjectTask.DeliverableId;
          const projectId = selectedProjectTask.ProjectId;

          // Delete all hierarchy entries
          try {
            const hierarchyItems = await sp.web.lists
              .getByTitle("ApprovalHierarchy")
              .items.select(
                "Id",
                "ProjectCreationListID/ID",
                "DeliverablesDetailsId/ID",
              )
              .filter(
                `ProjectCreationListID/ID eq ${projectId} and DeliverablesDetailsId/ID eq ${deliverableId}`,
              )
              .expand("ProjectCreationListID", "DeliverablesDetailsId")();

            for (const item of hierarchyItems) {
              if (item && item.Id) {
                await sp.web.lists
                  .getByTitle("ApprovalHierarchy")
                  .items.getById(item.Id)
                  .delete();
              }
            }
          } catch (error) {
            console.log("No entries to delete");
          }

          await updateDeliverableStatus(statusUpdate);

          // Send email notification
          const TouserUsers = users.filter(
            (u) => String(u.value) === String(dccId),
          );
          const TouserName =
            TouserUsers.length > 0 ? TouserUsers[0].label : null;
          const CCUserName = users.filter(
            (u) => String(u.value) === String(selectedProjectTask.InitiatedBy),
          );
          const CCName = CCUserName.length > 0 ? CCUserName[0].label : null;

          const emailBody = buildApprovalEmailBody(
            "SENT_FOR_PUBLISH",
            TouserName,
            CCName,
            projectDocumentInfo.map((d) => ({
              deliverable: selectedProjectTask.Deliverable,
              fileName: d.fileName,
              sharedLink: d.sharedLink,
              id: d.id,
            })),
          );
          await createEmailTriggerDetails(
            dccId!,
            [selectedProjectTask.InitiatedBy],
            emailBody,
            `${actionConfig["SENT_FOR_PUBLISH"].subject} – ${selectedProjectTask.Deliverable}`,
            projectDocumentInfo.map((d) => d.id),
          );

          // Update publish status
          const consolidatorDocIds = projectDocumentInfo
            .filter((d) => d.isConsolidatorCopy === "Yes")
            .map((d) => d.id);

          if (consolidatorDocIds.length > 0) {
            await updateDeliverablePublishStatus(
              selectedProjectTask.DeliverableId!,
              "No",
              consolidatorDocIds,
            );
          } else {
            const originalDocIds = projectDocumentInfo
              .filter((d) => d.isConsolidatorCopy !== "Yes")
              .map((d) => d.id);

            if (originalDocIds.length > 0) {
              await updateDeliverablePublishStatus(
                selectedProjectTask.DeliverableId!,
                "No",
                originalDocIds,
              );
            }
          }

          await updateDeliverableStatus(statusUpdate);
        } else {
          // For non-DC roles
          console.log("Processing approval for role:", currentRole);

          const currentApprovals = allApprovals.filter(
            (x: any) =>
              x.ApproverRole === currentRole &&
              x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
              x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId,
          );

          const allApproved =
            currentApprovals.length > 0 &&
            currentApprovals.every(
              (item: any) =>
                item.Status === "Approved" || item.Status === "Rework",
            );

          if (currentCriteria === "Anyone" && allApprovals.length > 0) {
            for (const approvalItem of allApprovals) {
              if (
                approvalItem.Id &&
                approvalItem.ApproverRole === currentRole
              ) {
                await sp.web.lists
                  .getByTitle("ProjectApprovals")
                  .items.getById(approvalItem.Id)
                  .update({ Status: statusUpdate, ApprovalDate: new Date() });

                const TouserUsers = users.filter(
                  (u) => String(u.value) === String(documentControllerId),
                );
                const TouserName =
                  TouserUsers.length > 0 ? TouserUsers[0].label : null;
                const CCUserName = users.filter(
                  (u) =>
                    String(u.value) === String(selectedProjectTask.InitiatedBy),
                );
                const CCName =
                  CCUserName.length > 0 ? CCUserName[0].label : null;

                const emailBody = buildApprovalEmailBody(
                  "APPROVED",
                  TouserName,
                  CCName,
                  projectDocumentInfo.map((d) => ({
                    deliverable: selectedProjectTask.Deliverable,
                    fileName: d.fileName,
                    sharedLink: d.sharedLink,
                    id: d.id,
                  })),
                );
                await createEmailTriggerDetails(
                  documentControllerId!,
                  [selectedProjectTask.InitiatedBy],
                  emailBody,
                  `${actionConfig["APPROVED"].subject} – ${selectedProjectTask.Deliverable}`,
                  projectDocumentInfo.map((d) => d.id),
                );
              }
            }
            await updateApprovalHierarchyStatus(allApprovals);
          }

          const shouldProceed =
            (currentCriteria === "Everyone" && allApproved) ||
            currentCriteria === "Anyone";
          if (shouldProceed) {
            console.log(
              "Should proceed - calling handleNextApprovers for role:",
              currentRole,
            );
            await handleNextApprovers();
          } else {
            console.log("Not proceeding - waiting for other approvers");
          }
        }
      }
    } catch (error) {
      console.error("Error in addMultipleApproval:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "addMultipleApproval",
        "MyApprovalWebPart",
      );
      throw error;
    }
  };

  const updateDeliverableStatus = async (status: string, isRework?: string) => {
    if (!selectedProjectTask) return;
    const data: any = { Status: status };
    if (isRework) data.IsReworked = isRework;
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
      const currentApprovals = allApprovals.filter(
        (x: any) =>
          x.ApproverRole === currentRole &&
          x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
          x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId,
      );
      if (currentApprovals.length === 0) return;

      const allApproved = currentApprovals.every(
        (item: any) => item.Status === "Approved",
      );
      let newStatus = "";
      if (currentCriteria === "Anyone") newStatus = "Approved";
      else if (currentCriteria === "Everyone")
        newStatus = allApproved ? "Approved" : "In Progress";

      const currentHierarchy = projectHierarchy.find(
        (x) => x.approverRole === currentRole && x.id,
      );
      if (currentHierarchy && currentHierarchy.id && newStatus !== "") {
        await sp.web.lists
          .getByTitle("ApprovalHierarchy")
          .items.getById(currentHierarchy.id)
          .update({ Status: newStatus });
      }
    } catch (error) {
      console.error("Error in updateApprovalHierarchyStatus:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "updateApprovalHierarchyStatus",
        "MyApprovalWebPart",
      );
    }
  };

  const handleNextApprovers = async () => {
    if (!selectedProjectTask) return;

    const currentHierarchy = projectHierarchy.find(
      (x) =>
        x.approverRole === selectedProjectTask.ApprovalRole &&
        x.level === selectedProjectTask.CurrentApprovalLevel,
    );

    // Find next approval level in the hierarchy
    const nextApprovers = currentHierarchy
      ? projectHierarchy.filter(
          (x) => x.serialNumber === currentHierarchy.serialNumber! + 1,
        )
      : [];

    if (!nextApprovers || nextApprovers.length === 0) {
      // No more approvers in hierarchy - final step before consolidator or DCC

      if (selectedProjectTask.ApprovalRole === "DCC") {
        // DCC (Document Control Committee) is final step - update publication status

        const publishValue =
          projectWantsToPublishInDossier !== "Select"
            ? projectWantsToPublishInDossier
            : "No";

        // Separate documents into consolidator backups and originals
        // Consolidator backups are created during consolidator review
        const consolidatorDocIds = projectDocumentInfo
          .filter((d) => d.isConsolidatorCopy === "Yes")
          .map((d) => d.id);
        const originalDocIds = projectDocumentInfo
          .filter((d) => d.isConsolidatorCopy !== "Yes")
          .map((d) => d.id);

        // Update deliverable status to Approved
        await updateDeliverableStatus("Approved");

        // Prioritize updating consolidator backup documents
        // If no backups exist, update original documents instead
        if (consolidatorDocIds.length > 0) {
          await updateDeliverablePublishStatus(
            selectedProjectTask.DeliverableId!,
            publishValue,
            consolidatorDocIds,
          );
        } else if (originalDocIds.length > 0) {
          await updateDeliverablePublishStatus(
            selectedProjectTask.DeliverableId!,
            publishValue,
            originalDocIds,
          );
        }
      } else {
        await createDCCApproval();
      }
    } else {
      // Create tasks for next approval level in the hierarchy
      for (const approvalItem of nextApprovers) {
        if (approvalItem.assignedTo && approvalItem.assignedTo.length > 0) {
          const allUsers = ([] as any[]).concat.apply(
            [],
            approvalItem.assignedTo,
          );
          console.log(
            "Creating tasks for users:",
            allUsers.map((u: any) => u.ID),
          );

          for (const user of allUsers) {
            try {
              // Create approval task for each user assigned to this level
              await sp.web.lists.getByTitle("ProjectApprovals").items.add({
                ProjectCreationListIDId: selectedProjectTask.ProjectId,
                DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
                AssignedToId: user.ID,
                RequestedById:
                  props.context.pageContext.legacyPageContext.userId,
                RequestedDate: new Date(),
                DocumentType: selectedProjectTask.DocType,
                ApproverRole: approvalItem.approverRole,
                ApprovalCriteria: approvalItem.approvalCriteria,
                Status: "Pending",
                Level: approvalItem.level,
                SerialNumber: approvalItem.serialNumber,
                RevisionNumber: selectedProjectTask.RevisionNumber,
              });

              // Send approval assignment email to the assigned user
              const TouserUsers = users.filter(
                (u) => String(u.value) === String(user.ID),
              );
              const TouserName =
                TouserUsers.length > 0 ? TouserUsers[0].label : null;
              const CCUserName = users.filter(
                (u) =>
                  String(u.value) === String(selectedProjectTask.InitiatedBy),
              );
              const CCName = CCUserName.length > 0 ? CCUserName[0].label : null;

              const emailBody = buildApprovalEmailBody(
                "APPROVAL_ASSIGNED",
                TouserName,
                CCName,
                projectDocumentInfo.map((d) => ({
                  deliverable: selectedProjectTask.Deliverable,
                  fileName: d.fileName,
                  sharedLink: d.sharedLink,
                  id: d.id,
                })),
              );
              await createEmailTriggerDetails(
                user.ID,
                [selectedProjectTask.InitiatedBy, documentControllerId],
                emailBody,
                `${actionConfig["APPROVAL_ASSIGNED"].subject} – ${selectedProjectTask.Deliverable}`,
                projectDocumentInfo.map((d) => d.id),
              );
            } catch (error) {
              console.error(`Error creating task for user ${user.ID}:`, error);
              await ErrorLogger.logError(
                sp,
                error,
                `handleNextApprovers.taskCreate`,
                "MyApprovalWebPart",
              );
            }
          }
        }
      }
    }
  };

  const getCleanDisplayName = (fileName: string): string => {
    if (!fileName) return "";

    // Remove file extension first (we'll add it back at the end)
    const lastDotIndex = fileName.lastIndexOf(".");
    const nameWithoutExt =
      lastDotIndex > -1 ? fileName.substring(0, lastDotIndex) : fileName;
    const extension = lastDotIndex > -1 ? fileName.substring(lastDotIndex) : "";

    // Apply the cleaning patterns
    let cleanName = nameWithoutExt
      .replace(/^\d+[_\s-]+/, "") // Remove leading numbers with separators
      .replace(/\d{6,}/g, "") // Remove long number sequences (timestamps)
      .replace(/-/g, " ") // Replace hyphens with spaces
      .replace(/_ConsolidatorBackup$/i, "") // Remove _ConsolidatorBackup suffix (case insensitive)
      .replace(/\s{2,}/g, " ") // Replace multiple spaces with single space
      .replace(/_+$/, "") // Remove trailing underscores at the end
      .replace(/^_+/, "") // Remove leading underscores at the start
      .trim();

    // Return cleaned name with original extension
    return cleanName + extension;
  };

  const createDCCApproval = async () => {
    if (!selectedProjectTask || !dccId) return;
    try {
      await sp.web.lists.getByTitle("ProjectApprovals").items.add({
        ProjectCreationListIDId: selectedProjectTask.ProjectId,
        DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
        AssignedToId: dccId,
        RequestedById: props.context.pageContext.legacyPageContext.userId,
        RequestedDate: new Date(),
        DocumentType: selectedProjectTask.DocType,
        ApproverRole: "DCC",
        ApprovalCriteria: "Anyone",
        Status: "Pending",
        Level: `Level ${selectedProjectTask.ApprovalSN + 1}`,
        SerialNumber: selectedProjectTask.ApprovalSN + 1,
        RevisionNumber: selectedProjectTask.RevisionNumber,
      });
      await updateDeliverableStatus("Approved");

      const TouserUsers = users.filter(
        (u) => String(u.value) === String(dccId),
      );
      const TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
      const CCUserName = users.filter(
        (u) => String(u.value) === String(selectedProjectTask.InitiatedBy),
      );
      const CCName = CCUserName.length > 0 ? CCUserName[0].label : null;

      const emailBody = buildApprovalEmailBody(
        "SENT_FOR_PUBLISH",
        TouserName,
        CCName,
        projectDocumentInfo.map((d) => ({
          deliverable: selectedProjectTask.Deliverable,
          fileName: d.fileName,
          sharedLink: d.sharedLink,
          id: d.id,
        })),
      );
      await createEmailTriggerDetails(
        dccId,
        [selectedProjectTask.InitiatedBy, documentControllerId],
        emailBody,
        `${actionConfig["SENT_FOR_PUBLISH"].subject} – ${selectedProjectTask.Deliverable}`,
        projectDocumentInfo.map((d) => d.id),
      );
    } catch (error) {
      console.error("Error creating DCC approval:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "createDCCApproval",
        "MyApprovalWebPart",
      );
      throw error;
    }
  };

  const handleReworkAction = async () => {
    if (!selectedProjectTask) return;
    if (selectedProjectTask.ApprovalRole !== "Document Controller") {
      const approvals = await getAllProjectApproval(
        selectedProjectTask.ProjectId!,
        selectedProjectTask.DeliverableId!,
      );
      const pendingApprovals = approvals.filter(
        (x: any) =>
          x.Level === selectedProjectTask?.CurrentApprovalLevel &&
          x.ApproverRole === selectedProjectTask?.ApprovalRole &&
          x.Id !== selectedProjectTask.Id &&
          x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
          x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId,
      );
      for (const item of pendingApprovals) {
        await sp.web.lists
          .getByTitle("ProjectApprovals")
          .items.getById(item.Id)
          .update({
            Status: "Rework",
            Remarks: "Auto Rework",
            ApprovalDate: new Date(),
          });
      }
      await createDocumentControllerApproval();
      await updateApprovalHierarchyStatusForRework("Rework");
    } else {
      await updateDeliverableStatus("Pending", "Yes");
      await createVendorApproval();
    }
  };

  const createDocumentControllerApproval = async () => {
    if (!selectedProjectTask) return;
    await sp.web.lists.getByTitle("ProjectApprovals").items.add({
      ProjectCreationListIDId: selectedProjectTask.ProjectId,
      DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
      AssignedToId: documentControllerId,
      RequestedById: props.context.pageContext.legacyPageContext.userId,
      RequestedDate: new Date(),
      DocumentType: selectedProjectTask.DocType,
      ApproverRole: "Document Controller",
      ApprovalCriteria: "Everyone",
      Doyouneedapproval: "Yes",
      Status: "Pending",
      Level: "Level 1",
      SerialNumber: 1,
      RevisionNumber: selectedProjectTask.RevisionNumber,
    });

    const TouserUsers = users.filter(
      (u) => String(u.value) === String(documentControllerId),
    );
    const TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
    const CCUserName = users.filter(
      (u) => String(u.value) === String(selectedProjectTask.InitiatedBy),
    );
    const CCName = CCUserName.length > 0 ? CCUserName[0].label : null;

    const emailBody = buildApprovalEmailBody(
      "REWORK_BY_APPROVER",
      TouserName,
      CCName,
      projectDocumentInfo.map((d) => ({
        deliverable: selectedProjectTask.Deliverable,
        fileName: d.fileName,
        sharedLink: d.sharedLink,
        id: d.id,
      })),
    );
    await createEmailTriggerDetails(
      documentControllerId!,
      [selectedProjectTask.InitiatedBy],
      emailBody,
      `${actionConfig["REWORK_BY_APPROVER"].subject} – ${selectedProjectTask.Deliverable}`,
      projectDocumentInfo.map((d) => d.id),
    );
  };

  const updateApprovalHierarchyStatusForRework = async (status: string) => {
    if (!selectedProjectTask) return;
    try {
      const filterQuery = `ProjectCreationListIDId eq ${selectedProjectTask.ProjectId} and DeliverablesDetailsIdId eq ${selectedProjectTask.DeliverableId} and Level eq '${selectedProjectTask.CurrentApprovalLevel}'`;
      const items = await sp.web.lists
        .getByTitle("ApprovalHierarchy")
        .items.filter(filterQuery)();
      if (items && items.length > 0) {
        for (const item of items) {
          await sp.web.lists
            .getByTitle("ApprovalHierarchy")
            .items.getById(item.Id)
            .update({ Status: status });
        }
      }
    } catch (error) {
      console.error("Error in updateApprovalHierarchyStatusForRework:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "updateApprovalHierarchyStatusForRework",
        "MyApprovalWebPart",
      );
    }
  };

  const createVendorApproval = async () => {
    if (!selectedProjectTask) return;
    await sp.web.lists.getByTitle("ProjectApprovals").items.add({
      ProjectCreationListIDId: selectedProjectTask.ProjectId,
      DeliverablesDetailsIdId: selectedProjectTask.DeliverableId,
      AssignedToId: selectedProjectTask.VendorId,
      RequestedById: props.context.pageContext.legacyPageContext.userId,
      RequestedDate: new Date(),
      DocumentType: selectedProjectTask.DocType,
      ApproverRole: "Vendor",
      ApprovalCriteria: "Anyone",
      Status: "pending",
      Level: "Level 0",
      SerialNumber: 0,
      RevisionNumber: selectedProjectTask.RevisionNumber,
    });

    const TouserUsers = users.filter(
      (u) => String(u.value) === String(selectedProjectTask.VendorId),
    );
    const TouserName = TouserUsers.length > 0 ? TouserUsers[0].label : null;
    const CCUserName = users.filter(
      (u) => String(u.value) === String(selectedProjectTask.InitiatedBy),
    );
    const CCName = CCUserName.length > 0 ? CCUserName[0].label : null;

    const emailBody = buildApprovalEmailBody(
      "REWORK_BY_DOCUMENT_CONTROLLER",
      TouserName,
      CCName,
      projectDocumentInfo.map((d) => ({
        deliverable: selectedProjectTask.Deliverable,
        fileName: d.fileName,
        sharedLink: d.sharedLink,
        id: d.id,
      })),
    );
    await createEmailTriggerDetails(
      selectedProjectTask.VendorId,
      [selectedProjectTask.InitiatedBy],
      emailBody,
      `${actionConfig["REWORK_BY_DOCUMENT_CONTROLLER"].subject} – ${selectedProjectTask.Deliverable}`,
      projectDocumentInfo.map((d) => d.id),
    );
  };

  const handleRejectedAction = async () => {
    if (!selectedProjectTask) return;
    const approvals = await getAllProjectApproval(
      selectedProjectTask.ProjectId!,
      selectedProjectTask.DeliverableId!,
    );
    const pendingApprovals = approvals.filter(
      (x: any) =>
        x.Level === selectedProjectTask?.CurrentApprovalLevel &&
        x.ApproverRole === selectedProjectTask?.ApprovalRole &&
        x.Id !== selectedProjectTask.Id &&
        x.ProjectCreationListIDId === selectedProjectTask.ProjectId &&
        x.DeliverablesDetailsIdId === selectedProjectTask.DeliverableId,
    );
    for (const item of pendingApprovals) {
      await sp.web.lists
        .getByTitle("ProjectApprovals")
        .items.getById(item.Id)
        .update({
          Status: "Rejected",
          Remarks: "Auto Reject",
          ApprovalDate: new Date(),
        });
    }
    await updateDeliverableStatus("Rejected");
    await updateApprovalHierarchyStatusForRework("Rejected");
  };

  const updateDeliverablePublishStatus = async (
    deliverableId: number,
    publishStatus: string,
    documentIds: number[],
  ): Promise<void> => {
    try {
      // Update deliverable's publication flag
      try {
        await sp.web.lists
          .getByTitle("DeliverablesDetails")
          .items.getById(deliverableId)
          .update({
            IsPublished: publishStatus,
          });
      } catch (err) {
        console.error("Error updating DeliverablesDetails:", err);
        await ErrorLogger.logError(
          sp,
          err,
          "updateDeliverablePublishStatus.DeliverablesDetails",
          "MyApprovalWebPart",
        );
      }

      // Update each document's publication and code status
      const deliverablesDocList = sp.web.lists.getByTitle(
        "DeliverablesDocument",
      );
      for (const docId of documentIds) {
        try {
          const updateData: any = {
            IsPublish: publishStatus,
          };

          // Include CodeStatus if DCC has selected one
          if (projectCodestatus && projectCodestatus !== "Select") {
            updateData.CodeStatus = projectCodestatus;
          }

          await deliverablesDocList.items.getById(docId).update(updateData);
        } catch (err) {
          console.error(`Error updating document ${docId}:`, err);
          await ErrorLogger.logError(
            sp,
            err,
            `updateDeliverablePublishStatus.document_${docId}`,
            "MyApprovalWebPart",
          );
        }
      }
    } catch (error) {
      console.error("Error in updateDeliverablePublishStatus:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "updateDeliverablePublishStatus",
        "MyApprovalWebPart",
      );
      throw error;
    }
  };

  const buildbulkApprovalEmailBody = (
    actionType: ActionType,
    documentNumber: string,
    approverName: string,
    senderName: string,
    projectName: string,
    docs: {
      deliverable: string;
      fileName: string;
      sharedLink: string;
      docrevisionNo: string;
      id: number;
    }[],
  ) => {
    const { message } = actionConfig[actionType];
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
        const tronUrl = `https://officeindia.sharepoint.com/sites/ESSA/SitePages/PDFViewer.aspx?docId=${d.id}`;
        return `<tr><td style="border:1px solid #ccc;padding:6px;text-align:center;">${no}</td><td style="border:1px solid #ccc;padding:6px;"><a href="${tronUrl}" target="_blank">${documentNumber}</a></td><td style="border:1px solid #ccc;padding:6px;">${d.deliverable}</td><td style="border:1px solid #ccc;padding:6px;text-align:center;">${d.docrevisionNo ?? "-"}</td><td style="border:1px solid #ccc;padding:6px;text-align:center;">${outgoingStatus ?? "-"}</td><td style="border:1px solid #ccc;padding:6px;">${transmittal}</td><td style="border:1px solid #ccc;padding:6px;">${sentDate}</td><td style="border:1px solid #ccc;padding:6px;">Sent for rework</td></tr>`;
      })
      .join("");

    const extraNo = docs.length + 1;
    const extraTransmittal = `TR-${String(extraNo).padStart(3, "0")}`;
    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const serverRel = props.context.pageContext.web.serverRelativeUrl;
    const projectNameEncoded = encodeURIComponent(projectName || "");
    const docNumberEncoded = encodeURIComponent(documentNumber || "");

    const dropboxUrl = `${siteUrl}/DeliverablesDocument/Forms/AllItems.aspx?id=${serverRel}%2FDeliverablesDocument%2F${projectNameEncoded}%2F${docNumberEncoded}`;
    rows += `<tr><td style="border:1px solid #ccc;padding:6px;text-align:center;">${extraNo}</td><td style="border:1px solid #ccc;padding:6px;"></td><td style="border:1px solid #ccc;padding:6px;">${projectName}</td><td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td><td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td><td style="border:1px solid #ccc;padding:6px;">${extraTransmittal}</td><td style="border:1px solid #ccc;padding:6px;">${sentDate}</td><td style="border:1px solid #ccc;padding:6px;"><a href="${dropboxUrl}" target="_blank">Dropbox Link</a></td></tr>`;

    return `Dear ${approverName},<br/><br/>${message}<br/><br/>
    <table style="border-collapse:collapse;width:100%;max-width:900px;font-size:13px;">
      <thead><tr style="background:#8eaada;font-weight:bold;">
        <th style="border:1px solid #ccc;padding:6px;text-align:center;">NO.#</th>
        <th style="border:1px solid #ccc;padding:6px;">Doc. #</th>
        <th style="border:1px solid #ccc;padding:6px;">Deliverables</th>
        <th style="border:1px solid #ccc;padding:6px;text-align:center;">Rev</th>
        <th style="border:1px solid #ccc;padding:6px;text-align:center;">Status</th>
        <th style="border:1px solid #ccc;padding:6px;">Transmittal</th>
        <th style="border:1px solid #ccc;padding:6px;">Sent Dated</th>
        <th style="border:1px solid #ccc;padding:6px;">Remarks</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <br/><br/>
    <a href="https://officeindia.sharepoint.com/sites/ESSA/SitePages/MyTasks.aspx" target="_blank">Click here to view in portal</a>
    <br/><br/>Regards,<br/>${senderName}`;
  };

  const buildApprovalEmailBody = (
    actionType: ActionType,
    approverName: string,
    senderName: string,
    docs: {
      deliverable: string;
      fileName: string;
      sharedLink: string;
      id: number;
    }[],
  ) => {
    const { message } = actionConfig[actionType];
    const redirectUrl =
      actionType === "REWORK_BY_DOCUMENT_CONTROLLER"
        ? "https://officeindia.sharepoint.com/sites/ESSA/SitePages/MyTasks.aspx"
        : "https://officeindia.sharepoint.com/sites/ESSA/SitePages/ESSAAPPROVALS.aspx";
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
        const tronUrl = `https://officeindia.sharepoint.com/sites/ESSA/SitePages/PDFViewer.aspx?docId=${d.id}`;
        return `<tr><td style="border:1px solid #ccc;padding:6px;text-align:center;">${no}</td><td style="border:1px solid #ccc;padding:6px;"><a href="${tronUrl}" target="_blank">${selectedProjectTask?.DocNumber}</a></td><td style="border:1px solid #ccc;padding:6px;">${d.deliverable}</td><td style="border:1px solid #ccc;padding:6px;text-align:center;">${selectedProjectTask?.RevisionNumber ?? "-"}</td><td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td><td style="border:1px solid #ccc;padding:6px;">${transmittal}</td><td style="border:1px solid #ccc;padding:6px;">${sentDate}</td><td style="border:1px solid #ccc;padding:6px;">Attached</td></tr>`;
      })
      .join("");

    const extraNo = docs.length + 1;
    const extraTransmittal = `TR-${String(extraNo).padStart(3, "0")}`;

    const siteUrl = props.context.pageContext.web.absoluteUrl;
    const serverRel = props.context.pageContext.web.serverRelativeUrl;

    const projectName = encodeURIComponent(
      selectedProjectTask?.ProjectName || "",
    );
    const docNumber = encodeURIComponent(selectedProjectTask?.DocNumber || "");

    const dropboxUrl = `${siteUrl}/DeliverablesDocument/Forms/AllItems.aspx?id=${serverRel}%2FDeliverablesDocument%2F${projectName}%2F${docNumber}`;
    rows += `<tr><td style="border:1px solid #ccc;padding:6px;text-align:center;">${extraNo}</td><td style="border:1px solid #ccc;padding:6px;"></td><td style="border:1px solid #ccc;padding:6px;">${selectedProjectTask?.ProjectName}</td><td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td><td style="border:1px solid #ccc;padding:6px;text-align:center;">-</td><td style="border:1px solid #ccc;padding:6px;">${extraTransmittal}</td><td style="border:1px solid #ccc;padding:6px;">${sentDate}</td><td style="border:1px solid #ccc;padding:6px;"><a href="${dropboxUrl}" target="_blank">Dropbox Link</a></td></tr>`;

    return `Dear ${approverName},<br/><br/>${message}<br/><br/>
    <table style="border-collapse:collapse;width:100%;max-width:900px;font-size:13px;">
      <thead><tr style="background:#8eaada;font-weight:bold;">
        <th style="border:1px solid #ccc;padding:6px;text-align:center;">NO.#</th>
        <th style="border:1px solid #ccc;padding:6px;">Doc. #</th>
        <th style="border:1px solid #ccc;padding:6px;">Deliverables</th>
        <th style="border:1px solid #ccc;padding:6px;text-align:center;">Rev</th>
        <th style="border:1px solid #ccc;padding:6px;text-align:center;">Status</th>
        <th style="border:1px solid #ccc;padding:6px;">Transmittal</th>
        <th style="border:1px solid #ccc;padding:6px;">Sent Dated</th>
        <th style="border:1px solid #ccc;padding:6px;">Remarks</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
    <br/><br/>
    <a href="${redirectUrl}" target="_blank">Click here to view in portal</a>
    <br/><br/>Regards,<br/>${senderName}`;
  };

  const createEmailTriggerDetails = async (
    TouserId: number,
    CCuserIds: any[],
    message: string,
    SubjectMsg: string,
    documentIds: number[],
  ) => {
    if (!selectedProjectTask) return;
    await sp.web.lists.getByTitle("EmailTriggerDetails").items.add({
      Title: selectedProjectTask.ProjectName,
      Subject: SubjectMsg,
      Body: message,
      ToUserId: [TouserId],
      CCUserId: CCuserIds,
    });
  };

  const createbulkEmailTriggerDetails = async (
    TouserId: number,
    CCuserIds: number[],
    message: string,
    SubjectMsg: string,
    documentIds: number[],
    Deliverable: string,
  ) => {
    await sp.web.lists.getByTitle("EmailTriggerDetails").items.add({
      Title: Deliverable,
      Subject: SubjectMsg,
      Body: message,
      ToUserId: [TouserId],
      CCUserId: CCuserIds,
    });
  };

  const fetchApproverRoles = async () => {
    try {
      const items = await sp.web.lists
        .getByTitle("ApproverRoleMasterList")
        .items.select("ApproverRole", "IsActive")
        .filter("IsActive eq 'Yes'")
        .orderBy("ApproverRole", true)();
      const roles = items.map((item: any) => ({
        value: item.ApproverRole,
        label: item.ApproverRole,
      }));
      setApproverRoles(roles);
    } catch (error) {
      console.error("Error fetching approver roles:", error);
      await ErrorLogger.logError(
        sp,
        error,
        "fetchApproverRoles",
        "MyApprovalWebPart",
      );
      setApproverRoles([
        { value: "Project Coordinator", label: "Project Coordinator" },
        { value: "Project Team", label: "Project Team" },
        { value: "Project Manager", label: "Project Manager" },
      ]);
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
              alt="Please wait..."
              style={{ width: "90px", height: "90px" }}
            />
            <p className="text-white mt-3 fw-semibold">please wait...</p>
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
                  <div
                    style={{ textAlign: "center" }}
                    className="col-md-3 newtexleft"
                  >
                    <div className="mb-0">
                      <label htmlFor="Status" className="form-label mt-2">
                        {/* <img src={require('../assets/delegation.png')} className='me-1' alt="d" />   */}
                        Acting on behalf of
                      </label>
                    </div>
                  </div>
                  <div style={{ paddingLeft: "0px" }} className="col-md-4">
                    <select
                      id="Type"
                      name="Type"
                      value={selectedBehalfEmail}
                      onChange={(e) => {
                        setSelectedBehalfEmail(e.target.value);
                        handleStatusChange(
                          e.target.name,
                          "Pending",
                          e.target.value,
                        );
                      }}
                      className="form-select"
                      disabled={loading || actingForUser.length === 0}
                    >
                      <option value="">
                        {loading
                          ? "Loading..."
                          : actingForUser.length === 0
                            ? "No Delegation"
                            : "Select an option"}
                      </option>
                      {actingForUser.map((item, index) => (
                        <option key={index} value={item.email}>
                          {item.name}
                        </option>
                      ))}
                    </select>
                  </div>
                  <div
                    style={{ textAlign: "center", padding: "0px" }}
                    className="col-md-1 newtexleft ivon"
                  >
                    <div className="mb-0">
                      <label
                        htmlFor="Status"
                        className="form-label newfil mt-2 mb-0"
                      >
                        Status
                        {/* <svg fill="#3c3c3c" width="23px" height="36px" viewBox="0 0 32 32" xmlns="http://www.w3.org/2000/svg" stroke="#b3b3b3"><g id="SVGRepo_bgCarrier" stroke-width="0"></g><g id="SVGRepo_tracerCarrier" stroke-linecap="round" stroke-linejoin="round"></g><g id="SVGRepo_iconCarrier"> <path d="M12,25l6.67,6.67a1,1,0,0,0,.7.29.91.91,0,0,0,.39-.08,1,1,0,0,0,.61-.92V13.08L31.71,1.71A1,1,0,0,0,31.92.62,1,1,0,0,0,31,0H1A1,1,0,0,0,.08.62,1,1,0,0,0,.29,1.71L11.67,13.08V24.33A1,1,0,0,0,12,25ZM3.41,2H28.59l-10,10a1,1,0,0,0-.3.71V28.59l-4.66-4.67V12.67a1,1,0,0,0-.3-.71Z"></path> </g></svg> */}
                      </label>
                    </div>
                  </div>
                  <div
                    style={{ paddingLeft: "0px" }}
                    className="col-md-4 ivon2"
                  >
                    <select
                      id="Type"
                      name="Type"
                      onChange={(e) =>
                        handleStatusChange(
                          e.target.name,
                          e.target.value,
                          actingforuseremail,
                        )
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
                        <li
                          className="nav-item myapprovalcomingsoon"
                          role="presentation"
                        >
                          {/* <li className="nav-item" role="presentation"> */}
                          <a
                            //onClick={() => handleTabClick("Intranet")}
                            className={`nav-link ${
                              activeTab === "Intranet" ? "active" : ""
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
                            className={`nav-link ${
                              activeTab === "DMS" ? "active" : ""
                            }`}
                            aria-selected={activeTab === "DMS"}
                            role="tab"
                            tabIndex={-1}
                          >
                            <span className="lenbg1">DMS </span>
                            <span className="lenbg">{Mylistdata.length}</span>
                          </a>
                        </li>

                        <li className="nav-item" role="presentation">
                          <a
                            onClick={() => handleTabClick("Automation")}
                            className={`nav-link ${
                              activeTab === "Automation" ? "active" : ""
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
                            className={`nav-link ${
                              activeTab === "ProjectWorkflow" ? "active" : ""
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
                        {selectedProjectItems.length > 0 && (
                          <>
                            <li className="nav-item" role="presentation">
                              <select
                                className="form-select me-2"
                                style={{ display: "inline-block" }}
                                value={outgoingStatus}
                                onChange={(e) =>
                                  setOutgoingStatus(e.target.value)
                                }
                              >
                                <option value="">Select MDR</option>
                                <option value="A">A</option>
                                <option value="B">B</option>
                                <option value="C">C</option>
                              </select>
                            </li>
                            <li className="nav-item" role="presentation">
                              <button
                                type="button"
                                className="btn btn-danger ms-2"
                                onClick={handlebulkReworkAction}
                              >
                                Rework
                              </button>
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
                      <div
                        style={{ width: "100%" }}
                        id="cardCollpase4"
                        className="collapse show"
                      >
                        <div className="table-responsive pt-0">
                          {activeTab === "Intranet" ||
                          activeTab === "Automation" ? (
                            <>
                              <div className="card card-body">
                                <table
                                  className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0"
                                  style={{
                                    position: "relative",
                                    width: "100%",
                                  }}
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
                                          style={{
                                            justifyContent: "space-evenly",
                                          }}
                                        >
                                          <span>S.No.</span>

                                          <span
                                            onClick={() =>
                                              handleSortChange("SNo")
                                            }
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
                                              if (
                                                e.key === "Enter" &&
                                                !e.shiftKey
                                              ) {
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
                                            style={{
                                              justifyContent: "space-evenly",
                                            }}
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
                                                handleFilterChange(
                                                  e,
                                                  "RequestID",
                                                )
                                              }
                                              onKeyDown={(e) => {
                                                if (
                                                  e.key === "Enter" &&
                                                  !e.shiftKey
                                                ) {
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
                                            style={{
                                              justifyContent: "space-evenly",
                                            }}
                                          >
                                            <span>Title</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("Title")
                                              }
                                            >
                                              <FontAwesomeIcon
                                                icon={faSort}
                                              />{" "}
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
                                                if (
                                                  e.key === "Enter" &&
                                                  !e.shiftKey
                                                ) {
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
                                            style={{
                                              justifyContent: "space-evenly",
                                            }}
                                          >
                                            <span>Process Name</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("ProcessName")
                                              }
                                            >
                                              <FontAwesomeIcon
                                                icon={faSort}
                                              />{" "}
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Process Name"
                                              onChange={(e) =>
                                                handleFilterChange(
                                                  e,
                                                  "ProcessName",
                                                )
                                              }
                                              onKeyDown={(e) => {
                                                if (
                                                  e.key === "Enter" &&
                                                  !e.shiftKey
                                                ) {
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
                                            style={{
                                              justifyContent: "space-evenly",
                                            }}
                                          >
                                            <span>Requested By</span>{" "}
                                            <span
                                              onClick={() =>
                                                handleSortChange("RequestedBy")
                                              }
                                            >
                                              <FontAwesomeIcon
                                                icon={faSort}
                                              />{" "}
                                            </span>
                                          </div>

                                          <div className=" bd-highlight">
                                            <input
                                              type="text"
                                              placeholder="Filter by Requested By"
                                              onChange={(e) =>
                                                handleFilterChange(
                                                  e,
                                                  "RequestedBy",
                                                )
                                              }
                                              onKeyDown={(e) => {
                                                if (
                                                  e.key === "Enter" &&
                                                  !e.shiftKey
                                                ) {
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
                                            style={{
                                              justifyContent: "space-evenly",
                                            }}
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
                                                  "RequestedDate",
                                                )
                                              }
                                              onKeyDown={(e) => {
                                                if (
                                                  e.key === "Enter" &&
                                                  !e.shiftKey
                                                ) {
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
                                            style={{
                                              justifyContent: "space-evenly",
                                            }}
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
                                                if (
                                                  e.key === "Enter" &&
                                                  !e.shiftKey
                                                ) {
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
                                            style={{
                                              justifyContent: "space-evenly",
                                            }}
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
                                    isActivedata,
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
                                    {((loading && currentData?.length == 0) ||
                                      StatusChange) && (
                                      <tr>
                                        <td
                                          colSpan={8}
                                          style={{
                                            textAlign: "center",
                                            padding: "100px 0",
                                            border: "none",
                                            height: "450px",
                                          }}
                                        >
                                          <div
                                            className="loadernewadd"
                                            style={{
                                              position: "static",
                                              display: "flex",
                                              flexDirection: "column",
                                              alignItems: "center",
                                            }}
                                          >
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
                                        <td
                                          colSpan={8}
                                          style={{ border: "none" }}
                                        >
                                          <div
                                            className="no-results card card-body align-items-center annusvg text-center"
                                            style={{
                                              display: "flex",
                                              justifyContent: "center",
                                              position: "relative",
                                              marginTop: "10px",
                                              height: "500px",
                                            }}
                                          >
                                            <svg
                                              style={{ top: "0%" }}
                                              xmlns="http://www.w3.org/2000/svg"
                                              width="28"
                                              height="28"
                                              viewBox="0 0 24 24"
                                              fill="none"
                                              stroke="currentColor"
                                              stroke-width="2"
                                              stroke-linecap="round"
                                              stroke-linejoin="round"
                                            >
                                              <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>
                                            </svg>
                                            <p className="font-14 text-muted text-center">
                                              No Approval found{" "}
                                            </p>
                                          </div>
                                        </td>
                                      </tr>
                                    ) : (
                                      // end of no approval found
                                      !StatusChange &&
                                      currentData?.map(
                                        (item: any, index: number) => (
                                          <tr key={index}>
                                            <td
                                              style={{
                                                minWidth: "40px",
                                                maxWidth: "40px",
                                              }}
                                            >
                                              <div className="d-flex align-items-center justify-content-center">
                                                {" "}
                                                <div
                                                  style={{ marginLeft: "0px" }}
                                                  className="indexdesign"
                                                >
                                                  {" "}
                                                  {startIndex + index + 1}
                                                </div>{" "}
                                              </div>{" "}
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "80px",

                                                maxWidth: "80px",
                                                textAlign: "center",

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
                                              title={
                                                activeTab == "Intranet"
                                                  ? item.Title
                                                  : item.ApprovalTitle
                                              }
                                            >
                                              {activeTab == "Intranet"
                                                ? item.Title
                                                : item.ApprovalTitle}
                                            </td>
                                            {/* )} */}
                                            <td
                                              style={{
                                                minWidth: "120px",
                                                maxWidth: "120px",
                                                textAlign: "center",
                                              }}
                                            >
                                              <span className="badge font-12 bg-secondary">
                                                {" "}
                                                {item.ProcessName}
                                              </span>
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "100px",
                                                maxWidth: "100px",
                                              }}
                                              title={
                                                activeTab == "Automation"
                                                  ? item?.Author?.Title
                                                  : item?.Requester?.Title
                                              }
                                            >
                                              {activeTab == "Automation"
                                                ? item?.Author?.Title
                                                : item?.Requester?.Title}
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "130px",
                                                maxWidth: "130px",
                                                textAlign: "center",
                                              }}
                                            >
                                              <div
                                                className="btn btn-light1 mt-1"
                                                style={{ width: "79%" }}
                                              >
                                                {/* {new Date(
                                                  item?.Created
                                                ).toLocaleDateString()} */}
                                                {new Date(
                                                  item?.Created,
                                                ).toLocaleString("en-US", {
                                                  month: "2-digit",
                                                  day: "2-digit",
                                                  year: "numeric",
                                                  hour: "2-digit",
                                                  minute: "2-digit",
                                                  //second: '2-digit',
                                                  hour12: true,
                                                })}
                                              </div>
                                            </td>

                                            <td
                                              style={{
                                                minWidth: "80px",
                                                maxWidth: "80px",
                                                textAlign: "center",
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
                                                textAlign: "center",
                                              }}
                                              className="fe-eye font-18"
                                            >
                                              {item?.Status.toLowerCase() ==
                                                "approved" ||
                                              item?.Status.toLowerCase() ==
                                                "rejected" ||
                                              item?.Status.toLowerCase() ==
                                                "completed" ? (
                                                <Eye
                                                  onClick={(e) =>
                                                    handleRedirect(
                                                      e,
                                                      item,
                                                      "view",
                                                    )
                                                  }
                                                  style={{
                                                    minWidth: "20px",

                                                    maxWidth: "20px",

                                                    cursor: "pointer",
                                                  }}
                                                />
                                              ) : (
                                                <Edit
                                                  onClick={(e) =>
                                                    handleRedirect(
                                                      e,
                                                      item,
                                                      "approval",
                                                    )
                                                  }
                                                  style={{
                                                    marginLeft: "0px",

                                                    cursor: "pointer",
                                                  }}
                                                />
                                              )}
                                            </td>
                                          </tr>
                                        ),
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
                                        className={`prevPage page-item ${
                                          currentGroup === 1 ? "disabled" : ""
                                        }`}
                                        onClick={() =>
                                          handleGroupChange("prev")
                                        }
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
                                              className={`page-item ${
                                                currentPage === pageNum
                                                  ? "active"
                                                  : ""
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
                                          );
                                        },
                                      )}

                                      <li
                                        className={`nextPage page-item ${
                                          currentGroup === totalGroups
                                            ? "disabled"
                                            : ""
                                        }`}
                                        onClick={() =>
                                          handleGroupChange("next")
                                        }
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
                                )}{" "}
                              </div>
                            </>
                          ) : null}

                          {/* // Aman added filter logic 18/3/26 */}

                          {activeTab === "DMS" && (
                            <div>
                              {!showNestedDMSTable ? (
                                <div className="card card-body">
                                  <table
                                    className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0"
                                    style={{
                                      position: "relative",
                                      width: "100%",
                                    }}
                                  >
                                    <thead>
                                      <tr>
                                        {/* aman 17/03/26 - Added search inputs for all DMS fields */}
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
                                            style={{
                                              justifyContent: "space-evenly",
                                            }}
                                          >
                                            <span>S.No.</span>
                                            <span
                                              onClick={() =>
                                                handleSortChange("SNo")
                                              }
                                            >
                                              <FontAwesomeIcon icon={faSort} />
                                            </span>
                                          </div>
                                          <div style={{ height: "34px" }}></div>
                                        </th>

                                        <th
                                          style={{
                                            minWidth: "80px",
                                            maxWidth: "80px",
                                          }}
                                        >
                                          <div className="d-flex flex-column bd-highlight">
                                            <div
                                              className="d-flex pb-2"
                                              style={{
                                                justifyContent: "space-evenly",
                                              }}
                                            >
                                              <span>Request ID</span>
                                              <span
                                                onClick={() =>
                                                  handleSortChange("RequestID")
                                                }
                                              >
                                                <FontAwesomeIcon
                                                  icon={faSort}
                                                />
                                              </span>
                                            </div>
                                            <div className="bd-highlight">
                                              <input
                                                type="text"
                                                className="inputcss"
                                                placeholder="Filter"
                                                value={filters.RequestID}
                                                onChange={(e) =>
                                                  handleFilterChange(
                                                    e,
                                                    "RequestID",
                                                  )
                                                }
                                                style={{ width: "100%" }}
                                              />
                                            </div>
                                          </div>
                                        </th>

                                        <th
                                          style={{
                                            minWidth: "120px",
                                            maxWidth: "120px",
                                          }}
                                        >
                                          <div className="d-flex flex-column bd-highlight">
                                            <div
                                              className="d-flex pb-2"
                                              style={{
                                                justifyContent: "space-evenly",
                                              }}
                                            >
                                              <span>Title</span>
                                              <span
                                                onClick={() =>
                                                  handleSortChange("Title")
                                                }
                                              >
                                                <FontAwesomeIcon
                                                  icon={faSort}
                                                />
                                              </span>
                                            </div>
                                            <div className="bd-highlight">
                                              <input
                                                type="text"
                                                className="inputcss"
                                                placeholder="Filter"
                                                value={filters.Title}
                                                onChange={(e) =>
                                                  handleFilterChange(e, "Title")
                                                }
                                                style={{ width: "100%" }}
                                              />
                                            </div>
                                          </div>
                                        </th>

                                        <th
                                          style={{
                                            minWidth: "120px",
                                            maxWidth: "120px",
                                          }}
                                        >
                                          <div className="d-flex flex-column bd-highlight">
                                            <div
                                              className="d-flex pb-2"
                                              style={{
                                                justifyContent: "space-evenly",
                                              }}
                                            >
                                              <span>Process Name</span>
                                              <span
                                                onClick={() =>
                                                  handleSortChange(
                                                    "ProcessName",
                                                  )
                                                }
                                              >
                                                <FontAwesomeIcon
                                                  icon={faSort}
                                                />
                                              </span>
                                            </div>
                                            <div className="bd-highlight">
                                              <input
                                                type="text"
                                                className="inputcss"
                                                placeholder="Filter"
                                                value={filters.ProcessName}
                                                onChange={(e) =>
                                                  handleFilterChange(
                                                    e,
                                                    "ProcessName",
                                                  )
                                                }
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
                                          <div className="d-flex flex-column bd-highlight">
                                            <div
                                              className="d-flex pb-2"
                                              style={{
                                                justifyContent: "space-evenly",
                                              }}
                                            >
                                              <span>Requested By</span>
                                              <span
                                                onClick={() =>
                                                  handleSortChange(
                                                    "RequestedBy",
                                                  )
                                                }
                                              >
                                                <FontAwesomeIcon
                                                  icon={faSort}
                                                />
                                              </span>
                                            </div>
                                            <div className="bd-highlight">
                                              <input
                                                type="text"
                                                className="inputcss"
                                                placeholder="Filter"
                                                value={filters.RequestedBy}
                                                onChange={(e) =>
                                                  handleFilterChange(
                                                    e,
                                                    "RequestedBy",
                                                  )
                                                }
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
                                          <div className="d-flex flex-column bd-highlight">
                                            <div
                                              className="d-flex pb-2"
                                              style={{
                                                justifyContent: "space-evenly",
                                              }}
                                            >
                                              <span>Requested Date</span>
                                            </div>
                                            <div className="bd-highlight">
                                              <input
                                                type="text"
                                                className="inputcss"
                                                placeholder="Filter"
                                                value={filters.RequestedDate}
                                                onChange={(e) =>
                                                  handleFilterChange(
                                                    e,
                                                    "RequestedDate",
                                                  )
                                                }
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
                                          <div className="d-flex flex-column bd-highlight">
                                            <div
                                              className="d-flex pb-2"
                                              style={{
                                                justifyContent: "space-evenly",
                                              }}
                                            >
                                              <span>Status</span>
                                            </div>
                                            <div className="bd-highlight">
                                              <input
                                                type="text"
                                                className="inputcss"
                                                placeholder="Filter"
                                                value={filters.Status}
                                                onChange={(e) =>
                                                  handleFilterChange(
                                                    e,
                                                    "Status",
                                                  )
                                                }
                                                style={{ width: "100%" }}
                                              />
                                            </div>
                                          </div>
                                        </th>

                                        <th
                                          style={{
                                            minWidth: "50px",
                                            maxWidth: "50px",
                                            textAlign: "center",
                                            verticalAlign: "top",
                                          }}
                                        >
                                          <div className="d-flex flex-column bd-highlight">
                                            <div
                                              className="d-flex pb-2"
                                              style={{
                                                justifyContent: "space-between",
                                              }}
                                            >
                                              <span>Action</span>
                                            </div>
                                          </div>
                                        </th>
                                      </tr>
                                    </thead>
                                    {console.log(
                                      "currentData",
                                      currentData,
                                      isActivedata,
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
                                        StatusChange) && (
                                        <tr>
                                          <td
                                            colSpan={8}
                                            style={{
                                              textAlign: "center",
                                              padding: "100px 0",
                                              border: "none",
                                              height: "450px",
                                            }}
                                          >
                                            <div
                                              className="loadernewadd"
                                              style={{
                                                position: "static",
                                                display: "flex",
                                                flexDirection: "column",
                                                alignItems: "center",
                                              }}
                                            >
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
                                          <td
                                            colSpan={8}
                                            style={{ border: "none" }}
                                          >
                                            <div
                                              className="no-results card card-body align-items-center annusvg text-center"
                                              style={{
                                                display: "flex",
                                                justifyContent: "center",
                                                position: "relative",
                                                marginTop: "10px",
                                                height: "500px",
                                              }}
                                            >
                                              <svg
                                                style={{ top: "0%" }}
                                                xmlns="http://www.w3.org/2000/svg"
                                                width="28"
                                                height="28"
                                                viewBox="0 0 24 24"
                                                fill="none"
                                                stroke="currentColor"
                                                stroke-width="2"
                                                stroke-linecap="round"
                                                stroke-linejoin="round"
                                              >
                                                <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>
                                              </svg>
                                              <p className="font-14 text-muted text-center">
                                                No Approval found{" "}
                                              </p>
                                            </div>
                                          </td>
                                        </tr>
                                      ) : (
                                        // end of no approval found for DMS
                                        !StatusChange &&
                                        currentData?.map(
                                          (item: any, index: number) => (
                                            <tr key={index}>
                                              <td
                                                style={{
                                                  minWidth: "40px",
                                                  maxWidth: "40px",
                                                }}
                                              >
                                                <div className="d-flex align-items-center justify-content-center">
                                                  {" "}
                                                  <div
                                                    style={{
                                                      marginLeft: "0px",
                                                    }}
                                                    className="indexdesign"
                                                  >
                                                    {" "}
                                                    {startIndex + index + 1}
                                                  </div>{" "}
                                                </div>{" "}
                                              </td>

                                              <td
                                                style={{
                                                  minWidth: "80px",

                                                  maxWidth: "80px",
                                                  textAlign: "center",

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
                                                  textAlign: "center",
                                                }}
                                              >
                                                <span className="badge font-12 bg-secondary">
                                                  {" "}
                                                  {
                                                    item?.FileUID?.Processname
                                                  }{" "}
                                                </span>
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
                                                  minWidth: "130px",
                                                  maxWidth: "130px",
                                                  // maxWidth: "100px",
                                                  textAlign: "center",
                                                }}
                                              >
                                                <div>
                                                  {/* {item?.FileUID?.Created} */}
                                                  {new Date(
                                                    item?.FileUID?.Created,
                                                  ).toLocaleString("en-US", {
                                                    month: "2-digit",
                                                    day: "2-digit",
                                                    year: "numeric",
                                                    hour: "2-digit",
                                                    minute: "2-digit",
                                                    //second: '2-digit',
                                                    hour12: true,
                                                  })}
                                                </div>
                                              </td>

                                              <td
                                                style={{
                                                  minWidth: "80px",
                                                  maxWidth: "80px",
                                                  textAlign: "center",
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
                                                  onClick={(e) => {
                                                    getTaskItemsbyID(
                                                      e,
                                                      item?.FileUID?.FileUID,
                                                      item?.FileUID
                                                        ?.Processname,
                                                    );
                                                    handleShowNestedDMSTable();
                                                  }}
                                                  style={{
                                                    marginLeft: "15px",

                                                    cursor: "pointer",
                                                  }}
                                                />
                                              </td>
                                            </tr>
                                          ),
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
                                          className={`prevPage page-item ${
                                            currentGroup === 1 ? "disabled" : ""
                                          }`}
                                          onClick={() =>
                                            handleGroupChange("prev")
                                          }
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
                                                className={`page-item ${
                                                  currentPage === pageNum
                                                    ? "active"
                                                    : ""
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
                                            );
                                          },
                                        )}

                                        <li
                                          className={`nextPage page-item ${
                                            currentGroup === totalGroups
                                              ? "disabled"
                                              : ""
                                          }`}
                                          onClick={() =>
                                            handleGroupChange("next")
                                          }
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
                                  {folderActionOrFileAction ===
                                    "New File Request" && (
                                    <>
                                      <DMSMyApprovalAction
                                        props={{
                                          currentItemID,
                                          actingforuseremail,
                                        }}
                                      />
                                      <div className="col-sm-12 text-center">
                                        {/* <button style={{ float: 'right' }} type="button" className="btn btn-secondary" onClick={() => setShowNestedDMSTable(false)}> Back </button> */}

                                        <button
                                          type="button"
                                          className="btn cancel-btn newp waves-effect waves-light m-3"
                                          style={{ fontSize: "0.875rem" }}
                                          onClick={() =>
                                            setShowNestedDMSTable(false)
                                          }
                                        >
                                          {/* <img
                                            src={require("../../../Assets/ExtraImage/xIcon.svg")}
                                            style={{ width: "1rem" }}
                                            className="me-1"
                                            alt="x"
                                          /> */}
                                          Cancel
                                        </button>
                                      </div>
                                    </>
                                  )}
                                  {folderActionOrFileAction ===
                                    "New Folder Request" && (
                                    <>
                                      <DMSMyFolderApprovalAction
                                        props={{
                                          currentItemID,
                                          actingforuseremail,
                                        }}
                                      />
                                      <div className="col-sm-12 text-center">
                                        {/* <button style={{ float: 'right' }} type="button" className="btn btn-secondary" onClick={() => setShowNestedDMSTable(false)}> Back </button> */}

                                        <button
                                          type="button"
                                          className="btn cancel-btn newp waves-effect waves-light m-3"
                                          style={{ fontSize: "0.875rem" }}
                                          onClick={() =>
                                            setShowNestedDMSTable(false)
                                          }
                                        >
                                          {/* <img
                                            src={require("../../../Assets/ExtraImage/xIcon.svg")}
                                            style={{ width: "1rem" }}
                                            className="me-1"
                                            alt="x"
                                          /> */}
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
                            <div className="mt-2">
                              <div id="cardCollpase4" className="collapse show">
                                <div className="table-responsive pt-0">
                                  {!showProjectForm ? (
                                    <div className="card card-body">
                                      <table
                                        className="mtbalenew mt-0 table-centered table-nowrap table-borderless respot mb-0"
                                        style={{
                                          position: "relative",
                                          width: "100%",
                                        }}
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
                                                style={{
                                                  justifyContent:
                                                    "space-evenly",
                                                }}
                                              >
                                                <span>S.No.</span>
                                                <span
                                                  onClick={() =>
                                                    handleSortChange("SNo")
                                                  }
                                                >
                                                  <FontAwesomeIcon
                                                    icon={faSort}
                                                  />
                                                </span>
                                              </div>
                                              <div
                                                style={{ height: "34px" }}
                                              ></div>
                                            </th>

                                            {currentData?.some(
                                              (item) =>
                                                item.ApprovalRole ===
                                                "Document Controller",
                                            ) && (
                                              <th
                                                style={{
                                                  minWidth: "40px",
                                                  maxWidth: "40px",
                                                }}
                                              >
                                                <div className="d-flex flex-column bd-highlight">
                                                  <div
                                                    className="d-flex pb-2"
                                                    style={{
                                                      justifyContent: "center",
                                                    }}
                                                  >
                                                    <span>Select</span>
                                                  </div>
                                                  <div
                                                    style={{ height: "34px" }}
                                                  ></div>
                                                </div>
                                              </th>
                                            )}

                                            <th
                                              style={{
                                                minWidth: "80px",
                                                maxWidth: "80px",
                                              }}
                                            >
                                              <div className="d-flex flex-column bd-highlight">
                                                <div
                                                  className="d-flex pb-2"
                                                  style={{
                                                    justifyContent:
                                                      "space-evenly",
                                                  }}
                                                >
                                                  <span>Request ID</span>
                                                  <span
                                                    onClick={() =>
                                                      handleSortChange(
                                                        "RequestID",
                                                      )
                                                    }
                                                  >
                                                    <FontAwesomeIcon
                                                      icon={faSort}
                                                    />
                                                  </span>
                                                </div>
                                                <div className="bd-highlight">
                                                  <input
                                                    type="text"
                                                    className="inputcss"
                                                    placeholder="Filter"
                                                    value={filters.RequestID}
                                                    onChange={(e) =>
                                                      handleFilterChange(
                                                        e,
                                                        "RequestID",
                                                      )
                                                    }
                                                    style={{ width: "100%" }}
                                                  />
                                                </div>
                                              </div>
                                            </th>

                                            <th
                                              style={{
                                                minWidth: "120px",
                                                maxWidth: "120px",
                                              }}
                                            >
                                              <div className="d-flex flex-column bd-highlight">
                                                <div
                                                  className="d-flex pb-2"
                                                  style={{
                                                    justifyContent:
                                                      "space-evenly",
                                                  }}
                                                >
                                                  <span>Project Name</span>
                                                  <span
                                                    onClick={() =>
                                                      handleSortChange("Title")
                                                    }
                                                  >
                                                    <FontAwesomeIcon
                                                      icon={faSort}
                                                    />
                                                  </span>
                                                </div>
                                                <div className="bd-highlight">
                                                  <input
                                                    type="text"
                                                    className="inputcss"
                                                    placeholder="Filter"
                                                    value={filters.Title}
                                                    onChange={(e) =>
                                                      handleFilterChange(
                                                        e,
                                                        "Title",
                                                      )
                                                    }
                                                    style={{ width: "100%" }}
                                                  />
                                                </div>
                                              </div>
                                            </th>

                                            <th
                                              style={{
                                                minWidth: "120px",
                                                maxWidth: "120px",
                                              }}
                                            >
                                              <div className="d-flex flex-column bd-highlight">
                                                <div
                                                  className="d-flex pb-2"
                                                  style={{
                                                    justifyContent:
                                                      "space-evenly",
                                                  }}
                                                >
                                                  <span>Deliverable Name</span>
                                                  <span
                                                    onClick={() =>
                                                      handleSortChange(
                                                        "DeliverableName",
                                                      )
                                                    }
                                                  >
                                                    <FontAwesomeIcon
                                                      icon={faSort}
                                                    />
                                                  </span>
                                                </div>
                                                <div className="bd-highlight">
                                                  <input
                                                    type="text"
                                                    className="inputcss"
                                                    placeholder="Filter"
                                                    onChange={(e) =>
                                                      handleFilterChange(
                                                        e,
                                                        "DeliverableName",
                                                      )
                                                    }
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
                                              <div className="d-flex flex-column bd-highlight">
                                                <div
                                                  className="d-flex pb-2"
                                                  style={{
                                                    justifyContent:
                                                      "space-evenly",
                                                  }}
                                                >
                                                  <span>Requested By</span>
                                                  <span
                                                    onClick={() =>
                                                      handleSortChange(
                                                        "RequestedBy",
                                                      )
                                                    }
                                                  >
                                                    <FontAwesomeIcon
                                                      icon={faSort}
                                                    />
                                                  </span>
                                                </div>
                                                <div className="bd-highlight">
                                                  <input
                                                    type="text"
                                                    className="inputcss"
                                                    placeholder="Filter"
                                                    value={filters.RequestedBy}
                                                    onChange={(e) =>
                                                      handleFilterChange(
                                                        e,
                                                        "RequestedBy",
                                                      )
                                                    }
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
                                              <div className="d-flex flex-column bd-highlight">
                                                <div
                                                  className="d-flex pb-2"
                                                  style={{
                                                    justifyContent:
                                                      "space-evenly",
                                                  }}
                                                >
                                                  <span>Requested Date</span>
                                                </div>
                                                <div className="bd-highlight">
                                                  <input
                                                    type="text"
                                                    className="inputcss"
                                                    placeholder="Filter"
                                                    value={
                                                      filters.RequestedDate
                                                    }
                                                    onChange={(e) =>
                                                      handleFilterChange(
                                                        e,
                                                        "RequestedDate",
                                                      )
                                                    }
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
                                              <div className="d-flex flex-column bd-highlight">
                                                <div
                                                  className="d-flex pb-2"
                                                  style={{
                                                    justifyContent:
                                                      "space-evenly",
                                                  }}
                                                >
                                                  <span>Status</span>
                                                </div>
                                                <div className="bd-highlight">
                                                  <input
                                                    type="text"
                                                    className="inputcss"
                                                    placeholder="Filter"
                                                    value={filters.Status}
                                                    onChange={(e) =>
                                                      handleFilterChange(
                                                        e,
                                                        "Status",
                                                      )
                                                    }
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
                                              <div className="d-flex flex-column bd-highlight">
                                                <div
                                                  className="d-flex pb-2"
                                                  style={{
                                                    justifyContent:
                                                      "space-evenly",
                                                  }}
                                                >
                                                  <span>Action</span>
                                                </div>
                                              </div>
                                            </th>
                                          </tr>
                                        </thead>
                                        <tbody>
                                          {((loading &&
                                            currentData?.length === 0) ||
                                            StatusChange) && (
                                            <tr>
                                              <td
                                                colSpan={
                                                  currentData?.some(
                                                    (item) =>
                                                      item.ApprovalRole ===
                                                      "Document Controller",
                                                  )
                                                    ? 9
                                                    : 8
                                                }
                                                style={{
                                                  textAlign: "center",
                                                  padding: "100px 0",
                                                  border: "none",
                                                  height: "450px",
                                                }}
                                              >
                                                <div
                                                  className="loadernewadd"
                                                  style={{
                                                    position: "static",
                                                    display: "flex",
                                                    flexDirection: "column",
                                                    alignItems: "center",
                                                  }}
                                                >
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
                                          {!loading &&
                                            currentData?.length === 0 &&
                                            !StatusChange && (
                                              <tr>
                                                <td
                                                  colSpan={
                                                    currentData?.some(
                                                      (item) =>
                                                        item.ApprovalRole ===
                                                        "Document Controller",
                                                    )
                                                      ? 9
                                                      : 8
                                                  }
                                                  style={{ border: "none" }}
                                                >
                                                  <div
                                                    className="no-results card card-body align-items-center annusvg text-center"
                                                    style={{
                                                      display: "flex",
                                                      justifyContent: "center",
                                                      position: "relative",
                                                      marginTop: "10px",
                                                      height: "500px",
                                                    }}
                                                  >
                                                    <svg
                                                      style={{ top: "0%" }}
                                                      xmlns="http://www.w3.org/2000/svg"
                                                      width="28"
                                                      height="28"
                                                      viewBox="0 0 24 24"
                                                      fill="none"
                                                      stroke="currentColor"
                                                      stroke-width="2"
                                                      stroke-linecap="round"
                                                      stroke-linejoin="round"
                                                    >
                                                      <polyline points="22 12 18 12 15 21 9 3 6 12 2 12"></polyline>
                                                    </svg>
                                                    <p className="font-14 text-muted text-center">
                                                      No Project Approvals found
                                                    </p>
                                                  </div>
                                                </td>
                                              </tr>
                                            )}
                                          {!StatusChange &&
                                            currentData?.map(
                                              (item: any, index: number) => (
                                                <tr key={index}>
                                                  <td
                                                    style={{
                                                      minWidth: "40px",
                                                      maxWidth: "40px",
                                                    }}
                                                  >
                                                    <div className="d-flex align-items-center justify-content-center">
                                                      <div
                                                        style={{
                                                          marginLeft: "0px",
                                                        }}
                                                        className="indexdesign"
                                                      >
                                                        {startIndex + index + 1}
                                                      </div>
                                                    </div>
                                                  </td>
                                                  {currentData?.some(
                                                    (item) =>
                                                      item.ApprovalRole ===
                                                      "Document Controller",
                                                  ) && (
                                                    <td
                                                      style={{
                                                        minWidth: "30px",
                                                        maxWidth: "30px",
                                                        textAlign: "center",
                                                      }}
                                                    >
                                                      {item.ApprovalRole ===
                                                        "Document Controller" &&
                                                        item.Status ===
                                                          "Pending" && (
                                                          <input
                                                            type="checkbox"
                                                            className="form-check-input"
                                                            checked={selectedProjectItems.some(
                                                              (selected) =>
                                                                selected.Id ===
                                                                item.Id,
                                                            )}
                                                            onChange={() =>
                                                              handleProjectItemSelect(
                                                                item.Id,
                                                              )
                                                            }
                                                            style={{
                                                              cursor: "pointer",
                                                            }}
                                                          />
                                                        )}
                                                    </td>
                                                  )}
                                                  <td
                                                    style={{
                                                      minWidth: "80px",
                                                      maxWidth: "80px",
                                                      textAlign: "center",
                                                      textTransform:
                                                        "capitalize",
                                                    }}
                                                    title={item.RequestID}
                                                  >
                                                    {item.RequestID}
                                                  </td>
                                                  <td
                                                    style={{
                                                      minWidth: "120px",
                                                      maxWidth: "120px",
                                                    }}
                                                    title={
                                                      item.ProjectName ||
                                                      item.Title
                                                    }
                                                  >
                                                    {item.ProjectName ||
                                                      item.Title}
                                                  </td>
                                                  <td
                                                    style={{
                                                      minWidth: "120px",
                                                      maxWidth: "120px",
                                                      textAlign: "center",
                                                    }}
                                                  >
                                                    <span className="badge font-12 bg-secondary">
                                                      {item.Deliverable}
                                                    </span>
                                                  </td>
                                                  <td
                                                    style={{
                                                      minWidth: "100px",
                                                      maxWidth: "100px",
                                                    }}
                                                    title={
                                                      item.Requester?.Title
                                                    }
                                                  >
                                                    {item.Requester?.Title}
                                                  </td>
                                                  <td
                                                    style={{
                                                      minWidth: "130px",
                                                      maxWidth: "130px",
                                                      textAlign: "center",
                                                    }}
                                                  >
                                                    <div
                                                      className="btn btn-light1 mt-1"
                                                      style={{ width: "79%" }}
                                                    >
                                                      {new Date(
                                                        item.Created,
                                                      ).toLocaleString(
                                                        "en-US",
                                                        {
                                                          month: "2-digit",
                                                          day: "2-digit",
                                                          year: "numeric",
                                                          hour: "2-digit",
                                                          minute: "2-digit",
                                                          hour12: true,
                                                        },
                                                      )}
                                                    </div>
                                                  </td>
                                                  <td
                                                    style={{
                                                      minWidth: "80px",
                                                      maxWidth: "80px",
                                                      textAlign: "center",
                                                    }}
                                                  >
                                                    <div className="btn btn-status">
                                                      {item.Status}
                                                    </div>
                                                  </td>
                                                  <td
                                                    style={{
                                                      minWidth: "50px",
                                                      maxWidth: "50px",
                                                      textAlign: "center",
                                                    }}
                                                    className="fe-eye font-18"
                                                  >
                                                    <Edit
                                                      onClick={(e) =>
                                                        handleProjectWorkflowAction(
                                                          e,
                                                          item,
                                                        )
                                                      }
                                                      style={{
                                                        cursor: "pointer",
                                                      }}
                                                    />
                                                  </td>
                                                </tr>
                                              ),
                                            )}
                                        </tbody>
                                      </table>
                                      {currentData?.length > 0 && (
                                        <nav className="pagination-container">
                                          <ul className="pagination">
                                            <li
                                              className={`prevPage page-item ${currentGroup === 1 ? "disabled" : ""}`}
                                              onClick={() =>
                                                handleGroupChange("prev")
                                              }
                                            >
                                              <a className="page-link">«</a>
                                            </li>
                                            {Array.from(
                                              {
                                                length: endPage - startPage + 1,
                                              },
                                              (_, num) => {
                                                const pageNum = startPage + num;
                                                return (
                                                  <li
                                                    key={pageNum}
                                                    className={`page-item ${currentPage === pageNum ? "active" : ""}`}
                                                  >
                                                    <a
                                                      className="page-link"
                                                      onClick={() =>
                                                        handlePageChange(
                                                          pageNum,
                                                        )
                                                      }
                                                    >
                                                      {pageNum}
                                                    </a>
                                                  </li>
                                                );
                                              },
                                            )}
                                            <li
                                              className={`nextPage page-item ${currentGroup === totalGroups ? "disabled" : ""}`}
                                              onClick={() =>
                                                handleGroupChange("next")
                                              }
                                            >
                                              <a
                                                className="page-link"
                                                onClick={() =>
                                                  handlePageChange(
                                                    currentPage + 1,
                                                  )
                                                }
                                              >
                                                »
                                              </a>
                                            </li>
                                          </ul>
                                        </nav>
                                      )}
                                    </div>
                                  ) : (
                                    <div className="card mt-2">
                                      <div className="card-body">
                                        <div className="form-header d-flex justify-content-between align-items-center mb-3">
                                          <h4 className="text-dark font-16 fw-bold m-0">
                                            Project Approval Details
                                          </h4>
                                        </div>
                                        <div className="row">
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>Project Name:</strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.ProjectName ||
                                                  ""
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>Project Type:</strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.ProjectType ||
                                                  "N/A"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>Deliverable:</strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.Deliverable ||
                                                  "N/A"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>Document Type:</strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.DocType ||
                                                  "N/A"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>
                                                  Document Number:
                                                </strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.DocNumber ||
                                                  "N/A"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>Area:</strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.Area ||
                                                  "N/A"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>Organization:</strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.Org ||
                                                  "N/A"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>Client Name:</strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.ClientName ||
                                                  "N/A"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>Prepared By:</strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.PreparedBy ||
                                                  "N/A"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          <div className="col-md-4">
                                            <div className="mb-3">
                                              <label className="form-label">
                                                <strong>
                                                  Revision Number:
                                                </strong>
                                              </label>
                                              <input
                                                type="text"
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.RevisionNumber ||
                                                  "0"
                                                }
                                                disabled
                                              />
                                            </div>
                                          </div>
                                          {projectDocumentInfo &&
                                          projectDocumentInfo.length > 0 ? (
                                            <div className="col-md-4">
                                              <label className="form-label">
                                                <strong>Documents:</strong>
                                                {selectedProjectTask?.ApprovalRole ===
                                                  "DCC" &&
                                                  projectDocumentInfo.some(
                                                    (d) =>
                                                      d.isConsolidatorCopy ===
                                                      "Yes",
                                                  ) && (
                                                    <span
                                                      className="text-muted ms-2"
                                                      style={{
                                                        fontSize: "12px",
                                                      }}
                                                    ></span>
                                                  )}
                                                {selectedProjectTask?.ApprovalRole ===
                                                  "DCC" &&
                                                  !projectDocumentInfo.some(
                                                    (d) =>
                                                      d.isConsolidatorCopy ===
                                                      "Yes",
                                                  ) && (
                                                    <span
                                                      className="text-muted ms-2"
                                                      style={{
                                                        fontSize: "12px",
                                                      }}
                                                    ></span>
                                                  )}
                                              </label>
                                              <div className="document-container">
                                                {(() => {
                                                  let docsToShow =
                                                    projectDocumentInfo;

                                                  // If role is Consolidator or DCC, check for backup copies
                                                  if (
                                                    selectedProjectTask?.ApprovalRole ===
                                                    "DCC"
                                                  ) {
                                                    const backupDocs =
                                                      projectDocumentInfo.filter(
                                                        (d) =>
                                                          d.isConsolidatorCopy ===
                                                          "Yes",
                                                      );
                                                    // Only show backup docs if they exist, otherwise show originals
                                                    if (backupDocs.length > 0) {
                                                      docsToShow = backupDocs;
                                                    }
                                                    // For DCC, if no backup docs, show all non-backup documents
                                                    else if (
                                                      selectedProjectTask?.ApprovalRole ===
                                                      "DCC"
                                                    ) {
                                                      docsToShow =
                                                        projectDocumentInfo.filter(
                                                          (d) =>
                                                            d.isConsolidatorCopy !==
                                                            "Yes",
                                                        );
                                                    }
                                                  } else {
                                                    // For other roles, exclude backup copies
                                                    docsToShow =
                                                      projectDocumentInfo.filter(
                                                        (d) =>
                                                          d.isConsolidatorCopy !==
                                                          "Yes",
                                                      );
                                                  }

                                                  return docsToShow.map(
                                                    (doc, index) => {
                                                      const displayName =
                                                        getCleanDisplayName(
                                                          doc.fileName ||
                                                            doc.fileLeafRef ||
                                                            "Document Available",
                                                        );
                                                      return (
                                                        <div
                                                          key={index}
                                                          className="document-link-container p-3 mb-2 border rounded bg-light"
                                                          style={{
                                                            cursor: "pointer",
                                                          }}
                                                          onClick={() =>
                                                            window.open(
                                                              `https://officeindia.sharepoint.com/sites/ESSA/SitePages/PDFViewer.aspx?docId=${doc.id}`,
                                                              "_blank",
                                                              "noopener,noreferrer",
                                                            )
                                                          }
                                                          title="Click to open document"
                                                        >
                                                          <div className="d-flex align-items-center">
                                                            <span
                                                              className="document-icon me-2"
                                                              style={{
                                                                fontSize:
                                                                  "1.5rem",
                                                              }}
                                                            >
                                                              📄
                                                            </span>
                                                            <div>
                                                              <div className="document-name fw-bold">
                                                                {displayName}
                                                              </div>
                                                              <div className="document-hint text-muted small">
                                                                Click to open
                                                                document in new
                                                                tab
                                                              </div>
                                                            </div>
                                                          </div>
                                                        </div>
                                                      );
                                                    },
                                                  );
                                                })()}
                                              </div>
                                            </div>
                                          ) : (
                                            <div className="alert alert-info mb-0">
                                              <small>
                                                No documents available for this
                                                deliverable
                                              </small>
                                            </div>
                                          )}

                                          {selectedProjectTask?.Remarks && (
                                            <div className="col-md-4">
                                              <label className="form-label">
                                                <strong>
                                                  Previous Remarks:
                                                </strong>
                                              </label>
                                              <textarea
                                                className="form-control"
                                                value={
                                                  selectedProjectTask?.Remarks ||
                                                  ""
                                                }
                                                disabled
                                                rows={3}
                                              />
                                            </div>
                                          )}

                                          {selectedProjectTask?.ApprovalRole ===
                                            "Document Controller" && (
                                            <>
                                              <div className="col-md-4">
                                                <label className="form-label">
                                                  <strong>
                                                    Do you need further
                                                    approval?
                                                  </strong>
                                                </label>
                                                <select
                                                  value={
                                                    projectNeedsFurtherApproval
                                                  }
                                                  onChange={(e) => {
                                                    setProjectNeedsFurtherApproval(
                                                      e.target.value,
                                                    );
                                                    setFieldErrors((prev) => ({
                                                      ...prev,
                                                      needsFurtherApproval: false,
                                                    }));
                                                  }}
                                                  disabled={isReadOnly}
                                                  className={`form-select ${fieldErrors.needsFurtherApproval ? styles.inputError : ""}`}
                                                  style={
                                                    fieldErrors.needsFurtherApproval
                                                      ? {
                                                          border:
                                                            "2px solid #fe0100",
                                                          backgroundColor:
                                                            "#fee6e6",
                                                        }
                                                      : {}
                                                  }
                                                >
                                                  <option
                                                    value="Select"
                                                    disabled
                                                  >
                                                    Select
                                                  </option>
                                                  <option value="Yes">
                                                    Yes
                                                  </option>
                                                  <option value="No">No</option>
                                                </select>
                                              </div>

                                              <div className="col-md-4">
                                                <label className="form-label">
                                                  <strong>
                                                    Outgoing status
                                                  </strong>
                                                </label>
                                                <select
                                                  value={projectOutgoingstatus}
                                                  onChange={(e) => {
                                                    setprojectOutgoingstatus(
                                                      e.target.value,
                                                    );
                                                    setFieldErrors((prev) => ({
                                                      ...prev,
                                                      outgoingStatus: false,
                                                    }));
                                                  }}
                                                  disabled={isReadOnly}
                                                  className={`form-select ${fieldErrors.outgoingStatus ? styles.inputError : ""}`}
                                                >
                                                  <option
                                                    value="Select"
                                                    disabled
                                                  >
                                                    Select
                                                  </option>
                                                  <option value="A">A</option>
                                                  <option value="B">B</option>
                                                  <option value="C">C</option>
                                                </select>
                                              </div>
                                            </>
                                          )}

                                          {selectedProjectTask?.ApprovalRole ===
                                            "DCC" && (
                                            <>
                                              <div className="col-md-4">
                                                <label className="form-label">
                                                  <strong>
                                                    Want to publish in Dossier?
                                                  </strong>
                                                </label>
                                                <select
                                                  value={
                                                    projectWantsToPublishInDossier
                                                  }
                                                  onChange={(e) => {
                                                    setProjectWantsToPublishInDossier(
                                                      e.target.value,
                                                    );
                                                    setFieldErrors((prev) => ({
                                                      ...prev,
                                                      wantsToPublishInDossier: false,
                                                    }));
                                                  }}
                                                  disabled={isReadOnly}
                                                  className={`form-select ${fieldErrors.wantsToPublishInDossier ? styles.inputError : ""}`}
                                                >
                                                  <option
                                                    value="Select"
                                                    disabled
                                                  >
                                                    Select
                                                  </option>
                                                  <option value="Yes">
                                                    Yes
                                                  </option>
                                                  <option value="No">No</option>
                                                </select>
                                              </div>

                                              <div className="col-md-4">
                                                <label className="form-label">
                                                  <strong>Code status</strong>
                                                </label>
                                                <select
                                                  value={projectCodestatus}
                                                  onChange={(e) => {
                                                    setprojectCodestatus(
                                                      e.target.value,
                                                    );
                                                    setFieldErrors((prev) => ({
                                                      ...prev,
                                                      codeStatus: false,
                                                    }));
                                                  }}
                                                  disabled={isReadOnly}
                                                  className={`form-select ${fieldErrors.codeStatus ? styles.inputError : ""}`}
                                                >
                                                  <option
                                                    value="Select"
                                                    disabled
                                                  >
                                                    Select
                                                  </option>
                                                  <option value="A">A</option>
                                                  <option value="B">B</option>
                                                  <option value="C">C</option>
                                                </select>
                                              </div>
                                            </>
                                          )}
                                          <div className="col-md-4">
                                            <div className="col-12">
                                              <label className="form-label">
                                                <strong>Remarks</strong>
                                              </label>
                                              <textarea
                                                value={projectRemarks}
                                                onChange={(e) => {
                                                  setProjectRemarks(
                                                    e.target.value,
                                                  );
                                                  setFieldErrors((prev) => ({
                                                    ...prev,
                                                    remarks: false,
                                                  }));
                                                }}
                                                placeholder="Enter your remarks here..."
                                                rows={2}
                                                disabled={isReadOnly}
                                                className={`form-control ${fieldErrors.remarks ? styles.textareaError : ""}`}
                                                style={{
                                                  resize: "both",
                                                  overflow: "auto",
                                                }}
                                              />
                                            </div>
                                          </div>
                                          {showDocumentComments && (
                                            <div
                                              className={styles.accordionItem}
                                            >
                                              <h2
                                                className={
                                                  styles.accordionHeader
                                                }
                                              >
                                                <div
                                                  className={
                                                    styles.accordionButton
                                                  }
                                                >
                                                  Document Comments
                                                </div>
                                              </h2>
                                              <div
                                                className={styles.accordionBody}
                                              >
                                                <div
                                                  className={styles.customCard}
                                                >
                                                  <div
                                                    className={
                                                      styles.documentCommentsHeader
                                                    }
                                                  >
                                                    <select
                                                      className={
                                                        styles.formSelect
                                                      }
                                                      value={selectedVersion}
                                                      onChange={(e) =>
                                                        onVersionChange(
                                                          e.target.value,
                                                        )
                                                      }
                                                    >
                                                      <option value="">
                                                        -- Select Revision --
                                                      </option>
                                                      {versionList.map(
                                                        (version) => (
                                                          <option
                                                            key={version}
                                                            value={version}
                                                          >
                                                            {version}
                                                          </option>
                                                        ),
                                                      )}
                                                    </select>
                                                    <button
                                                      className={
                                                        styles.btnOutlineSuccess
                                                      }
                                                      type="button"
                                                      onClick={
                                                        exportCommentsToExcel
                                                      }
                                                    >
                                                      Export to Excel
                                                    </button>
                                                    <button
                                                      className={
                                                        styles.btnOutlineSuccess
                                                      }
                                                      type="button"
                                                      onClick={
                                                        refreshDocComment
                                                      }
                                                    >
                                                      ↻
                                                    </button>
                                                  </div>
                                                  <div
                                                    className={
                                                      styles.ribbonContent
                                                    }
                                                  >
                                                    <table
                                                      className={
                                                        styles.commentsTable
                                                      }
                                                    >
                                                      <thead>
                                                        <tr>
                                                          <th
                                                            style={{
                                                              minWidth: "80px",
                                                            }}
                                                          >
                                                            Users
                                                          </th>
                                                          <th
                                                            style={{
                                                              minWidth: "100px",
                                                            }}
                                                          >
                                                            Comment Date
                                                          </th>
                                                          <th
                                                            style={{
                                                              minWidth: "80px",
                                                            }}
                                                          >
                                                            Page No.
                                                          </th>
                                                          <th
                                                            style={{
                                                              minWidth: "80px",
                                                            }}
                                                          >
                                                            Revision
                                                          </th>
                                                          <th
                                                            style={{
                                                              minWidth: "200px",
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
                                                        {documentComments.map(
                                                          (commentItem) => (
                                                            <tr
                                                              key={
                                                                commentItem.id
                                                              }
                                                            >
                                                              <td
                                                                style={{
                                                                  padding:
                                                                    "10px",
                                                                }}
                                                              >
                                                                {
                                                                  commentItem.userName
                                                                }
                                                              </td>
                                                              <td
                                                                style={{
                                                                  padding:
                                                                    "10px",
                                                                }}
                                                              >
                                                                {
                                                                  commentItem.commentDate
                                                                }
                                                              </td>
                                                              <td
                                                                style={{
                                                                  padding:
                                                                    "10px",
                                                                }}
                                                              >
                                                                {
                                                                  commentItem.pageNumber
                                                                }
                                                              </td>
                                                              <td
                                                                style={{
                                                                  padding:
                                                                    "10px",
                                                                }}
                                                              >
                                                                {
                                                                  commentItem.revision
                                                                }
                                                              </td>
                                                              <td
                                                                style={{
                                                                  padding:
                                                                    "15px",
                                                                }}
                                                              >
                                                                {
                                                                  commentItem.comment
                                                                }
                                                              </td>
                                                            </tr>
                                                          ),
                                                        )}
                                                      </tbody>
                                                    </table>
                                                  </div>
                                                </div>
                                              </div>
                                            </div>
                                          )}
                                        </div>

                                        {selectedProjectTask?.ApprovalRole ===
                                          "Document Controller" &&
                                          projectNeedsFurtherApproval ===
                                            "Yes" && (
                                            <div className="card">
                                              <div className="card-body">
                                                <div className="approval-projectHierarchy mt-4">
                                                  <div className="d-flex justify-content-between align-items-center mb-3">
                                                    <h5
                                                      style={{
                                                        margin: "inherit",
                                                      }}
                                                    >
                                                      Approval Hierarchy
                                                    </h5>
                                                    {!isReadOnly && (
                                                      <button
                                                        type="button"
                                                        className="btn btn-primary btn-sm"
                                                        onClick={() =>
                                                          addNewProjectApprovalRow()
                                                        }
                                                      >
                                                        + Add New Row
                                                      </button>
                                                    )}
                                                  </div>
                                                  <div className="table-responsive">
                                                    <table className="table table-bordered">
                                                      <thead className="table-light">
                                                        <tr>
                                                          <th
                                                            style={{
                                                              width: "15%",
                                                            }}
                                                          >
                                                            Level
                                                          </th>
                                                          <th
                                                            style={{
                                                              width: "25%",
                                                            }}
                                                          >
                                                            Approver Role
                                                          </th>
                                                          <th
                                                            style={{
                                                              width: "30%",
                                                            }}
                                                          >
                                                            Approver
                                                          </th>
                                                          <th
                                                            style={{
                                                              width: "20%",
                                                            }}
                                                          >
                                                            Approval Criteria
                                                          </th>
                                                          {!isReadOnly && (
                                                            <th
                                                              style={{
                                                                width: "10%",
                                                              }}
                                                            >
                                                              Action
                                                            </th>
                                                          )}
                                                        </tr>
                                                      </thead>
                                                      <tbody>
                                                        {projectHierarchy.map(
                                                          (row, index) => (
                                                            <tr key={index}>
                                                              <td>
                                                                <input
                                                                  type="text"
                                                                  className="form-control"
                                                                  value={
                                                                    row.level
                                                                  }
                                                                  disabled
                                                                />
                                                              </td>
                                                              <td>
                                                                <select
                                                                  value={
                                                                    row.approverRole
                                                                  }
                                                                  onChange={(
                                                                    e,
                                                                  ) => {
                                                                    updateProjectApprovalRow(
                                                                      index,
                                                                      "approverRole",
                                                                      e.target
                                                                        .value,
                                                                    );
                                                                    setFieldErrors(
                                                                      (
                                                                        prev,
                                                                      ) => ({
                                                                        ...prev,
                                                                        [`hierarchy_role_${index}`]: false,
                                                                      }),
                                                                    );
                                                                  }}
                                                                  disabled={
                                                                    isReadOnly
                                                                  }
                                                                  className={`form-select ${fieldErrors[`hierarchy_role_${index}`] ? styles.inputError : ""}`}
                                                                  style={
                                                                    fieldErrors[
                                                                      `hierarchy_role_${index}`
                                                                    ]
                                                                      ? {
                                                                          border:
                                                                            "2px solid #fe0100",
                                                                          backgroundColor:
                                                                            "#fee6e6",
                                                                        }
                                                                      : {}
                                                                  }
                                                                >
                                                                  <option
                                                                    value=""
                                                                    disabled
                                                                  >
                                                                    Select Role
                                                                  </option>
                                                                  {approverRoles.length >
                                                                  0 ? (
                                                                    approverRoles.map(
                                                                      (
                                                                        role,
                                                                        idx,
                                                                      ) => (
                                                                        <option
                                                                          key={
                                                                            idx
                                                                          }
                                                                          value={
                                                                            role.value
                                                                          }
                                                                        >
                                                                          {
                                                                            role.label
                                                                          }
                                                                        </option>
                                                                      ),
                                                                    )
                                                                  ) : (
                                                                    <>
                                                                      <option value="Project Coordinator">
                                                                        Project
                                                                        Coordinator
                                                                      </option>
                                                                      <option value="Project Team">
                                                                        Project
                                                                        Team
                                                                      </option>
                                                                      <option value="Project Manager">
                                                                        Project
                                                                        Manager
                                                                      </option>
                                                                    </>
                                                                  )}
                                                                </select>
                                                              </td>
                                                              <td>
                                                                <Select
                                                                  isMulti
                                                                  options={
                                                                    users
                                                                  }
                                                                  value={
                                                                    row.assignedTo
                                                                      ? users.filter(
                                                                          (
                                                                            user,
                                                                          ) =>
                                                                            row.assignedTo?.some(
                                                                              (
                                                                                assigned: any,
                                                                              ) =>
                                                                                (assigned.ID
                                                                                  ? assigned.ID.toString()
                                                                                  : assigned.toString()) ===
                                                                                user.value,
                                                                            ),
                                                                        )
                                                                      : []
                                                                  }
                                                                  onChange={(
                                                                    selectedOptions: any,
                                                                  ) => {
                                                                    handleProjectApproverChange(
                                                                      index,
                                                                      selectedOptions,
                                                                    );
                                                                    setFieldErrors(
                                                                      (
                                                                        prev,
                                                                      ) => ({
                                                                        ...prev,
                                                                        [`hierarchy_approver_${index}`]: false,
                                                                      }),
                                                                    );
                                                                  }}
                                                                  isDisabled={
                                                                    isReadOnly
                                                                  }
                                                                  placeholder="Select Approver(s)"
                                                                  className={`people-picker ${fieldErrors[`hierarchy_approver_${index}`] ? styles.selectError : ""}`}
                                                                  classNamePrefix="react-select"
                                                                  closeMenuOnSelect={
                                                                    false
                                                                  }
                                                                  isClearable={
                                                                    false
                                                                  }
                                                                  styles={
                                                                    fieldErrors[
                                                                      `hierarchy_approver_${index}`
                                                                    ]
                                                                      ? {
                                                                          control:
                                                                            (
                                                                              base: any,
                                                                            ) => ({
                                                                              ...base,
                                                                              border:
                                                                                "2px solid #fe0100",
                                                                              backgroundColor:
                                                                                "#fee6e6",
                                                                            }),
                                                                        }
                                                                      : {}
                                                                  }
                                                                />
                                                              </td>
                                                              <td>
                                                                <select
                                                                  value={
                                                                    row.approvalCriteria
                                                                  }
                                                                  onChange={(
                                                                    e,
                                                                  ) => {
                                                                    updateProjectApprovalRow(
                                                                      index,
                                                                      "approvalCriteria",
                                                                      e.target
                                                                        .value,
                                                                    );
                                                                    setFieldErrors(
                                                                      (
                                                                        prev,
                                                                      ) => ({
                                                                        ...prev,
                                                                        [`hierarchy_criteria_${index}`]: false,
                                                                      }),
                                                                    );
                                                                  }}
                                                                  disabled={
                                                                    isReadOnly
                                                                  }
                                                                  className={`form-select ${fieldErrors[`hierarchy_criteria_${index}`] ? styles.inputError : ""}`}
                                                                  style={
                                                                    fieldErrors[
                                                                      `hierarchy_criteria_${index}`
                                                                    ]
                                                                      ? {
                                                                          border:
                                                                            "2px solid #fe0100",
                                                                          backgroundColor:
                                                                            "#fee6e6",
                                                                        }
                                                                      : {}
                                                                  }
                                                                >
                                                                  <option
                                                                    value=""
                                                                    disabled
                                                                  >
                                                                    Select
                                                                    Criteria
                                                                  </option>
                                                                  <option value="Everyone">
                                                                    Everyone
                                                                  </option>
                                                                  <option value="Anyone">
                                                                    Anyone
                                                                  </option>
                                                                </select>
                                                              </td>
                                                              {!isReadOnly && (
                                                                <td className="text-center">
                                                                  <button
                                                                    type="button"
                                                                    className="btn btn-outline-danger btn-sm"
                                                                    onClick={() =>
                                                                      deleteProjectApprovalRow(
                                                                        index,
                                                                      )
                                                                    }
                                                                    title={
                                                                      projectHierarchy.length <=
                                                                      1
                                                                        ? "Cannot delete last row"
                                                                        : "Delete Row"
                                                                    }
                                                                    disabled={
                                                                      projectHierarchy.length <=
                                                                      1
                                                                    }
                                                                  >
                                                                    🗑️
                                                                  </button>
                                                                </td>
                                                              )}
                                                            </tr>
                                                          ),
                                                        )}
                                                      </tbody>
                                                    </table>
                                                  </div>
                                                </div>
                                              </div>
                                            </div>
                                          )}

                                        <div className="row mt-4">
                                          <div className="col-12">
                                            <div className="d-flex justify-content-center gap-3">
                                              {selectedProjectTask?.Status ===
                                                "Pending" && (
                                                <>
                                                  {selectedProjectTask?.ApprovalRole ===
                                                  "DCC" ? (
                                                    <>
                                                      <button
                                                        type="button"
                                                        className="btn btn-success"
                                                        disabled={
                                                          showSubmitLoader
                                                        }
                                                        onClick={() => {
                                                          if (
                                                            !validateApprovalAction(
                                                              "Approved",
                                                            )
                                                          )
                                                            return;
                                                          setPendingAction(
                                                            "Approved",
                                                          );
                                                          setPopupType(
                                                            "confirmation",
                                                          );
                                                          setPopupTitle(
                                                            "Confirm Submission",
                                                          );
                                                          setPopupMessage(
                                                            `Are you sure you want to submit?${projectWantsToPublishInDossier === "Yes" ? "\n\nPublish in Dossier: Yes\nCode Status: " + projectCodestatus : "\n\nPublish in Dossier: No"}`,
                                                          );
                                                          setPopupOpen(true);
                                                        }}
                                                      >
                                                        Submit
                                                      </button>
                                                      <button
                                                        type="button"
                                                        className="btn btn-secondary"
                                                        onClick={
                                                          handleProjectBackClick
                                                        }
                                                      >
                                                        Cancel
                                                      </button>
                                                    </>
                                                  ) : (
                                                    <>
                                                      <button
                                                        type="button"
                                                        className="btn btn-success"
                                                        disabled={
                                                          showSubmitLoader
                                                        }
                                                        onClick={() => {
                                                          if (
                                                            !validateApprovalAction(
                                                              "Approved",
                                                            )
                                                          )
                                                            return;
                                                          setPendingAction(
                                                            "Approved",
                                                          );
                                                          setPopupType(
                                                            "confirmation",
                                                          );
                                                          setPopupTitle(
                                                            "Confirm Approval",
                                                          );
                                                          setPopupMessage(
                                                            "Are you sure you want to approve this request?",
                                                          );
                                                          setPopupOpen(true);
                                                        }}
                                                      >
                                                        Approve
                                                      </button>
                                                      <button
                                                        type="button"
                                                        className="btn btn-danger"
                                                        disabled={
                                                          showSubmitLoader
                                                        }
                                                        onClick={() => {
                                                          if (
                                                            !validateApprovalAction(
                                                              "Rejected",
                                                            )
                                                          )
                                                            return;
                                                          setPendingAction(
                                                            "Rejected",
                                                          );
                                                          setPopupType(
                                                            "confirmation",
                                                          );
                                                          setPopupTitle(
                                                            "Confirm Rejection",
                                                          );
                                                          setPopupMessage(
                                                            "Are you sure you want to reject this request?",
                                                          );
                                                          setPopupOpen(true);
                                                        }}
                                                      >
                                                        Reject
                                                      </button>
                                                      <button
                                                        type="button"
                                                        className="btn btn-warning"
                                                        disabled={
                                                          showSubmitLoader
                                                        }
                                                        onClick={() => {
                                                          if (
                                                            !validateApprovalAction(
                                                              "Rework",
                                                            )
                                                          )
                                                            return;
                                                          setPendingAction(
                                                            "Rework",
                                                          );
                                                          setPopupType(
                                                            "confirmation",
                                                          );
                                                          setPopupTitle(
                                                            "Send for Rework",
                                                          );
                                                          setPopupMessage(
                                                            "Do you want to send this item for rework?",
                                                          );
                                                          setPopupOpen(true);
                                                        }}
                                                      >
                                                        Rework
                                                      </button>
                                                      <button
                                                        type="button"
                                                        className="btn btn-secondary"
                                                        onClick={
                                                          handleProjectBackClick
                                                        }
                                                      >
                                                        Cancel
                                                      </button>
                                                    </>
                                                  )}
                                                </>
                                              )}
                                              {(selectedProjectTask?.Status ===
                                                "Approved" ||
                                                selectedProjectTask?.Status ===
                                                  "Rejected" ||
                                                selectedProjectTask?.Status ===
                                                  "Rework") && (
                                                <button
                                                  type="button"
                                                  className="btn btn-secondary"
                                                  onClick={
                                                    handleProjectBackClick
                                                  }
                                                >
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
                              </div>
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
                                            e.target.value,
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
