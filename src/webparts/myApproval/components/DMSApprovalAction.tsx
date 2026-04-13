declare global {
  interface Window {
    // managePermission:(message:string) => void;
    // manageWorkflow:(documentLibrayName: string, SiteTitle: string) => void;
    view: (message: string) => void;
    // PreviewFile: (path: string, siteID: string, docLibName:any) => void;
    deleteFile: (fileId: string, siteID: string, listToUpdate: any) => void;
  }
}
interface UploadFileProps {
  currentfolderpath: {
    CurrentEntity: string;
    currentEntityURL: string;
    currentsiteID: string;
    // ... other properties
  };
}

// @ts-ignore
import * as React from "react";
import { getSP } from "../loc/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "bootstrap/dist/css/bootstrap.min.css";
// import "bootstrap//dist/"
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss";
import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";
import UserContext from "../../../GlobalContext/context";
// import { useState , useEffect } from "react";
import Provider from "../../../GlobalContext/provider";
import { useMediaQuery } from "react-responsive";
import "@pnp/sp/webs";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import "@pnp/sp/sites"
import "@pnp/sp/presets/all";
import moment from "moment";
import { PermissionKind } from "@pnp/sp/security";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
// import "../../verticalSideBar/components/VerticalSidebar.scss";
// import "./dmscss";
// import "./DMSAdmincss"
import { useState, useRef, useEffect } from "react";
import { IDmsapprovalProps } from './IDmsapprovalProps'
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
//  import EntityMapping from "./EntityMapping";

// import './ApprovalActioncss.css'

import { faEye } from "@fortawesome/free-regular-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import DMSMyApproval from "./MyApprovals";
import Swal from "sweetalert2";
import { Web } from "@pnp/sp/webs";
import { AssignFrom } from "@pnp/core";
let approvedLevel: any = ''
let filepreviewurl = ''
let remark: any = ''
let readablefilepreviewurl:any
let editablefilepreviewurl:any
let Level: any = ''
let setFinalStatus: any = ''
let FileUID: any = ''
const DMSMyApprovalAction = ({ props }: any) => {
  console.log(props, "here is my props")
  const sp: SPFI = getSP();

  const { useHide }: any = React.useContext(UserContext);
  const elementRef = React.useRef<HTMLDivElement>(null);



  React.useEffect(() => {
    // console.log("This function is called only once", useHide);

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
  // Media query to check if the screen width is less than 768px
  const isMobile = useMediaQuery({ query: "(max-width: 768px)" });

  React.useEffect(() => {
    // console.log("This function is called only once", useHide);

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
  React.useEffect(() => {
    const handleEscape = (e: KeyboardEvent) => {
      if (e.key === "Escape") {
        if (document.fullscreenElement) {
          document.exitFullscreen();
        }
      }
    };

    window.addEventListener("keydown", handleEscape);
    return () => window.removeEventListener("keydown", handleEscape);
  }, []);
  ////////////////////////////////////  DMS Code Start From Here //////////////////////////////////////////////////////////////////
  // const [Mylistdata, setMylistdata] = useState([]);
  // Add <any[]> here
const [Mylistdata, setMylistdata] = useState<any[]>([]);
  ////
  const [isFullScreen, setIsFullScreen] = useState(false);
  //save button for edit file
  const [isEditMode, setIsEditMode] = useState(false);


  const [storedUserInfo, setStoredUserInfo] = useState(null);
  const [ApprovedStatus, setApprovedStatus] = useState('');  // State for ApprovalType 0
  // const [approvedLevel, setApprovedLevel] = useState<number>();
  const [targetSiteUrl, setTargetSiteUrl] = useState<string>("");

  const [activeComponent, setActiveComponent] = useState<string>('');
  const [showAISummary, setShowAISummary] = useState(false);
  const [aiSummary, setAiSummary] = useState<string>("");
  const [openfileon, setopenfileon] = React.useState(false);
   useEffect(() => {
      getApprovalmasterTasklist();
    getCurrrentuser()

  }, []);
// Add state for editable URL
  const [editableUrl, setEditableUrl] = useState<string>("");
  ////
  console.log(activeComponent, "activeComponent")
  const getUserTitleByEmail = async (userEmail: any) => {
    try {
      const user = await sp.web.siteUsers.getByEmail(userEmail)();
      return user.Title;
    } catch (error) {
      console.error("Error fetching user title:", error);
      return null;
    }
  };
  const handleReturnToMain = (Name: any) => {
    setActiveComponent(Name); // Reset to show the main component
    console.log(activeComponent, "activeComponent updated")
  };
//   const getApprovalmasterTasklist = async () => {
//     const spinner = document.getElementById("spinner");
//     if (spinner) {
//       spinner.style.display = "block"; // Show the spinner
//     }
//     try {
//       const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
//         "Log", "CurrentUser", "Remark"
//         , "LogHistory", "ID"
//         , "FileUID/FileUID"
//         , "FileUID/SiteName"
//         , "FileUID/DocumentLibraryName"
//         , "FileUID/FileName"
//         , "FileUID/RequestNo"
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
//         .filter(`FileUID/FileUID eq '${props.currentItemID}'`)
//         .orderBy("Modified", false)();
//         console.log(items, "DMSFileApprovalTaskList");
//          // Fetch user titles
//       const updatedItems = await Promise.all(items.map(async (item) => {
//         const userTitle = await getUserTitleByEmail(item.FileUID.RequestedBy);
//         const assignedtouserTitle = await getUserTitleByEmail(item.CurrentUser);

//         // alert(userTitle)
//         return { ...item, RequestedByTitle: userTitle, AssignedToTitle: assignedtouserTitle };
//       }));



//       // start
//       items.forEach((item) => {
//         console.log(item, "item data ")
//         console.log(currentUserEmailRef.current, "currentUserEmailRef.current")
//         console.log(item.CurrentUser, "item.CurrentUser")
//         console.log(item.Log, "item.Log")
//         if(props.actingforuseremail !== null && props.actingforuseremail !== undefined && props.actingforuseremail !== ""){
//           if(props.actingforuseremail === item.CurrentUser){
//             if (item.Log === null) {
//               setToggleLog(true);
//           }
//           }
//         }else{
//           if(currentUserEmailRef.current === item.CurrentUser ){
//             if (item.Log === null) {
//               setToggleLog(true);
//             }
//           }
//         }
//         //     if(currentUserEmailRef.current === item.CurrentUser || props.actingforuseremail === item.CurrentUser){
//         //   if (item.LogHistory === null) {
//         //     setToggleLog(true);
//         //   }

//         // }
//         console.log("FileUID", item.FileUID.FileUID)
//         FileUID = item.FileUID.FileUID

//       })
//       // end

//       setMylistdata(updatedItems);


//     } catch (error) {
//       console.error("Error fetching list items:", error);
//     }
//     // try {
//     //   console.log("here")
//     //   const updatedData:any = await sp.web.lists.getByTitle("DMSFileApprovalList").items
//     //   .select("FileUID", "ID", "ApproveAction", "ApprovedLevel", "SiteName", "DocumentLibraryName", "ApprovedLevel" , "FilePreviewUrl")
//     //   .filter(`FileUID eq '${FileUID}'`)()
//     //   .catch((error) => console.error("Error fetching data from DMSFileApprovalList:", error));
//     //   console.log(updatedData , "updatedData")
//     //     filepreviewurl = updatedData[0]?.FilePreviewUrl;
//     //     console.log(filepreviewurl , "file url")
//     // } catch (error) {
//     //   console.error("Error fetching list items:", error);
//     // }
//     try {
//       console.log("here")
//       const updatedData: any = await sp.web.lists.getByTitle("DMSFileApprovalList").items
//         .select("FileUID", "ID", "ApproveAction", "ApprovedLevel", "SiteName", "DocumentLibraryName", "ApprovedLevel", "FilePreviewUrl", "FolderPath", "FileName","AISummary" , "RequestedBy")
//         .filter(`FileUID eq '${FileUID}'`)()
//         .catch((error) => console.error("Error fetching data from DMSFileApprovalList:", error));
//       console.log(updatedData, "updatedData")
//       filepreviewurl = updatedData[0]?.FilePreviewUrl;
//       setAiSummary(updatedData[0]?.AISummary || "");
//       console.log(filepreviewurl, "file url")

//       const siteData = await sp.web.lists.getByTitle('MasterSiteURL').items.select("Id", "SiteID").filter(`Title eq '${updatedData[0]?.SiteName}'`)();
//       console.log("siteData", siteData);
//       const { web } = await sp.site.openWebById(siteData[0].SiteID);

//       // Get the list item  corresponding to the file
//          const fileItem:any = await web.getFileById(props.currentItemID).expand("ListItemAllFields")();
//       console.log("fileItem", fileItem.ListItemAllFields.Status);

//       // fetched the columns details corresponding to the file 
//       const fileColumns = await sp.web.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName", "SiteName", "DocumentLibraryName", "IsRename").filter(`SiteName eq '${updatedData[0]?.SiteName}' and DocumentLibraryName eq '${updatedData[0]?.DocumentLibraryName}' and IsDocumentLibrary ne 1`)();
//       console.log("fileColumns", fileColumns);

//       // Create an array of objects to store the columnName with there corresponding value
//       const resultArrayThatContainstheColumnDetails = fileColumns.map((column) => {

//         let columnName = column.ColumnName;
//         const columnValue = fileItem.ListItemAllFields[columnName];
//         if (column.IsRename !== null) {
//           columnName = column.IsRename
//         }

//         return {
//           label: columnName,
//           value: columnValue !== undefined ? columnValue : null // Handle missing fields
//         };
//       });

//       const objectForStatus = {
//         label: "Status",
//         value: fileItem.ListItemAllFields.Status || ""
//       }
//       const objectforfilename = {
//         label: "File Name",
//         value: updatedData[0]?.FileName
//       }
//       const objectforRequestedby = {
//         label: "Request By",
//         value: updatedData[0]?.RequestedBy
//       }
//       const objectForDocumentLibrary = {
//         label: "Folder Name",
//         value: updatedData[0]?.FolderPath.substring(updatedData[0]?.FolderPath.lastIndexOf('/') + 1) || ""
//       }

//       const objectForSiteName = {
//         label: "Site Name",
//         value: updatedData[0]?.SiteName || ""
//       }

//       resultArrayThatContainstheColumnDetails.push(objectForStatus);
//       resultArrayThatContainstheColumnDetails.push(objectforfilename);
//       resultArrayThatContainstheColumnDetails.push(objectForSiteName);
//       resultArrayThatContainstheColumnDetails.push(objectForDocumentLibrary);
//       resultArrayThatContainstheColumnDetails.push(objectforRequestedby);
//       console.log("result", resultArrayThatContainstheColumnDetails);
//       // Variable to hold generated HTML
//       let detailRowsHTML = "";

//       // Dynamically generate HTML with inline CSS
//       resultArrayThatContainstheColumnDetails.forEach((item, index) => {
//         // Start a new row every 3rd item
//         if (index % 3 === 0) {
//           detailRowsHTML += `<div style="margin-bottom: 10px;" class="row">`;
//         }

//         // Add detail column for each item with inline CSS
//       detailRowsHTML += `
//   <div style="padding: 12px; min-width: 0;" class="col-sm-6">
//     <div style="
//       font-weight: bold; 
//       margin-bottom: 5px;
//       overflow: hidden;
//       text-overflow: ellipsis; 
//       white-space: nowrap;
//     " title="${item.label || ''}">${item.label || ''}</div>
//     <div style="
//       overflow: hidden;
//       text-overflow: ellipsis;
//       white-space: nowrap;
//       color: #333;
//     " title="${item.value || 'N/A'}">${item.value || 'N/A'}</div>
//   </div>
// `;

//         // Close the row after every 3rd item
//         if ((index + 1) % 3 === 0) {
//           detailRowsHTML += `</div>`;
//         }
//       });

//       // Close the last row if not already closed
//       if (resultArrayThatContainstheColumnDetails.length % 3 !== 0) {
//         detailRowsHTML += `</div>`;
//       }

//       // Add the generated HTML to a container
//       // document.getElementById("dynamicDetailsContainer").innerHTML = detailRowsHTML;
//       const container = document.getElementById("dynamicDetailsContainer");
// if (container) {
//   container.innerHTML = detailRowsHTML;
// }

//       if (filepreviewurl) {
//         // previewFile(filepreviewurl)
//         readablefilepreviewurl = filepreviewurl
//       }
//     } catch (error) {
//       console.error("Error fetching list items:", error);
//     }

//   };
//  Function to convert URL to editable format
// Function to convert URL to editable format

const getApprovalmasterTasklist = async () => {
    const spinner = document.getElementById("spinner");
    if (spinner) {
      spinner.style.display = "block"; // Show the spinner
    }

    let foundSiteUrl = ""; // To track which remote site contains our data

    try {
      // START CROSS-SITE MODIFICATION
      const masterSites = await sp.web.lists
        .getByTitle("MasterSiteCollection")
        .items.select("SiteURL")();

      let allFoundItems: any[] = [];

      for (const site of masterSites) {
        try {
          const remoteWeb = Web(site.SiteURL).using(AssignFrom(sp.web));
          const items = await remoteWeb.lists.getByTitle('DMSFileApprovalTaskList').items.select(
            "Log", "CurrentUser", "Remark"
            , "LogHistory", "ID"
            , "FileUID/FileUID"
            , "FileUID/SiteName"
            , "FileUID/DocumentLibraryName"
            , "FileUID/FileName"
            , "FileUID/RequestNo"
            , "FileUID/Status"
            , "FileUID/FolderPath"
            , "FileUID/RequestedBy"
            , "FileUID/Created"
            , "FileUID/ApproveAction"
            , "MasterApproval/ApprovalType"
            , "MasterApproval/Level"
            , "MasterApproval/DocumentLibraryName"
          )
            .expand("FileUID", "MasterApproval")
            .filter(`FileUID/FileUID eq '${props.currentItemID}'`)
            .orderBy("Modified", false)();
          
          if(items.length > 0) {
            allFoundItems = [...allFoundItems, ...items];
            foundSiteUrl = site.SiteURL; // Capture the correct site URL
            setTargetSiteUrl(site.SiteURL); // Save for button actions
          }
        } catch (e) {
          console.error("Error fetching from site: " + site.SiteURL, e);
        }
      }
      
      const items = allFoundItems; 
      // END CROSS-SITE MODIFICATION

      console.log(items, "DMSFileApprovalTaskList");
      // Fetch user titles
      const updatedItems = await Promise.all(items.map(async (item) => {
        const userTitle = await getUserTitleByEmail(item.FileUID.RequestedBy);
        const assignedtouserTitle = await getUserTitleByEmail(item.CurrentUser);

        // alert(userTitle)
        return { ...item, RequestedByTitle: userTitle, AssignedToTitle: assignedtouserTitle };
      }));

      // start
      items.forEach((item) => {
        console.log(item, "item data ")
        console.log(currentUserEmailRef.current, "currentUserEmailRef.current")
        console.log(item.CurrentUser, "item.CurrentUser")
        console.log(item.Log, "item.Log")
        if(props.actingforuseremail !== null && props.actingforuseremail !== undefined && props.actingforuseremail !== ""){
          if(props.actingforuseremail === item.CurrentUser){
            if (item.Log === null) {
              setToggleLog(true);
            }
          }
        }else{
          if(currentUserEmailRef.current === item.CurrentUser ){
            if (item.Log === null) {
              setToggleLog(true);
            }
          }
        }
        console.log("FileUID", item.FileUID.FileUID)
        FileUID = item.FileUID.FileUID
      })
      // end

      setMylistdata(updatedItems);

    } catch (error) {
      console.error("Error fetching list items:", error);
    }

    try {
      console.log("here")
      
      // Pivot the web context to the site where the data was found to avoid 404
      const targetWeb = foundSiteUrl ? Web(foundSiteUrl).using(AssignFrom(sp.web)) : sp.web;

      const updatedData: any = await targetWeb.lists.getByTitle("DMSFileApprovalList").items
        .select("FileUID", "ID", "ApproveAction", "ApprovedLevel", "SiteName", "DocumentLibraryName", "ApprovedLevel", "FilePreviewUrl", "FolderPath", "FileName","AISummary" , "RequestedBy")
        .filter(`FileUID eq '${FileUID}'`)()
        .catch((error) => console.error("Error fetching data from DMSFileApprovalList:", error));
      
      console.log(updatedData, "updatedData")
      filepreviewurl = updatedData[0]?.FilePreviewUrl;
      setAiSummary(updatedData[0]?.AISummary || "");
      console.log(filepreviewurl, "file url")

      const siteData = await sp.web.lists.getByTitle('MasterSiteURL').items.select("Id", "SiteURL", "SiteID").filter(`Title eq '${updatedData[0]?.SiteName}'`)();
      console.log("siteData", siteData);
      
     if (siteData.length > 0) {
            // FIX: Create the Web context using the specific Sub-site URL (e.g., .../N27JAN)
            const fileWeb = Web(siteData[0].SiteURL).using(AssignFrom(sp.web));

            // Get the file and its library columns
            const fileItem: any = await fileWeb.getFileById(props.currentItemID).expand("ListItemAllFields")();
            console.log("fileItem", fileItem.ListItemAllFields.Status);

      // fetched the columns details corresponding to the file 
      // Note: These master config lists are usually on the ESSA site, so we use sp.web here
      const fileColumns = await targetWeb.lists.getByTitle("DMSPreviewFormMaster").items.select("ColumnName", "SiteName", "DocumentLibraryName", "IsRename").filter(`SiteName eq '${updatedData[0]?.SiteName}' and DocumentLibraryName eq '${updatedData[0]?.DocumentLibraryName}' and IsDocumentLibrary ne 1`)();
      console.log("fileColumns", fileColumns);

      // Create an array of objects to store the columnName with there corresponding value
      const resultArrayThatContainstheColumnDetails = fileColumns.map((column) => {

        let columnName = column.ColumnName;
        const columnValue = fileItem.ListItemAllFields[columnName];
        if (column.IsRename !== null) {
          columnName = column.IsRename
        }

        return {
          label: columnName,
          value: columnValue !== undefined ? columnValue : null // Handle missing fields
        };
      });
      

      const objectForStatus = {
        label: "Status",
        value: fileItem.ListItemAllFields.Status || ""
      }
      const objectforfilename = {
        label: "File Name",
        value: updatedData[0]?.FileName
      }
      const objectforRequestedby = {
        label: "Request By",
        value: updatedData[0]?.RequestedBy
      }
      const objectForDocumentLibrary = {
        label: "Folder Name",
        value: updatedData[0]?.FolderPath.substring(updatedData[0]?.FolderPath.lastIndexOf('/') + 1) || ""
      }

      const objectForSiteName = {
        label: "Site Name",
        value: updatedData[0]?.SiteName || ""
      }

      resultArrayThatContainstheColumnDetails.push(objectForStatus);
      resultArrayThatContainstheColumnDetails.push(objectforfilename);
      resultArrayThatContainstheColumnDetails.push(objectForSiteName);
      resultArrayThatContainstheColumnDetails.push(objectForDocumentLibrary);
      resultArrayThatContainstheColumnDetails.push(objectforRequestedby);
      console.log("result", resultArrayThatContainstheColumnDetails);
      
      let detailRowsHTML = "";

      resultArrayThatContainstheColumnDetails.forEach((item, index) => {
        if (index % 3 === 0) {
          detailRowsHTML += `<div style="margin-bottom: 10px;" class="row">`;
        }

      detailRowsHTML += `
  <div style="padding: 12px; min-width: 0;" class="col-sm-6">
    <div style="
      font-weight: bold; 
      margin-bottom: 5px;
      overflow: hidden;
      text-overflow: ellipsis; 
      white-space: nowrap;
    " title="${item.label || ''}">${item.label || ''}</div>
    <div style="
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      color: #333;
    " title="${item.value || 'N/A'}">${item.value || 'N/A'}</div>
  </div>
`;

        if ((index + 1) % 3 === 0) {
          detailRowsHTML += `</div>`;
        }
      });

      if (resultArrayThatContainstheColumnDetails.length % 3 !== 0) {
        detailRowsHTML += `</div>`;
      }

      const container = document.getElementById("dynamicDetailsContainer");
      if (container) {
        container.innerHTML = detailRowsHTML;
      }

      if (filepreviewurl) {
        readablefilepreviewurl = filepreviewurl
      }
    }
    } catch (error) {
      console.error("Error fetching list items:", error);
    } finally {
      if (spinner) {
        spinner.style.display = "none";
      }
    }
};

const getEditableUrl = (readUrl: string): string => {
  try {
    const url = new URL(readUrl);
    const idParam = url.searchParams.get('id');
    
    if (idParam) {
      const serverRelativePath = decodeURIComponent(idParam);
      const path = serverRelativePath.split("/sites/Intranetdemos")[1];
      return `https://officeindia.sharepoint.com/:w:/r/sites/Intranetdemos${encodeURIComponent(path)}?web=1`;
    }
  } catch (error) {
    console.error("Error converting URL:", error);
  }
  return readUrl; // Return original URL if conversion fails
};
// / Handle toggle click
// Handle toggle click
const handleToggleClick = (event: any) => {
  event.preventDefault();
  event.stopPropagation();
  
  const newState = !openfileon;
  setopenfileon(newState);
  
  if (newState && readablefilepreviewurl) {
    // Opening - set editable URL
    const editUrl = readablefilepreviewurl;
    setEditableUrl(editUrl);
    setIsEditMode(true);
  } else {
    // Closing - clear the URL
    setEditableUrl("");
    setIsEditMode(false);
  }
};
  console.log(Mylistdata, "Mylistdata")
  // start
  const [toggleLog, setToggleLog] = useState(false);
  // end

  const currentUserEmailRef = useRef('');
  const getCurrrentuser = async () => {
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
  
  }


  const getTaskItemsbyID = async (e: any, itemid: any) => {
    console.log("itemid", itemid)
    const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select("CurrentUser", "FileUID/FileUID", "Log").expand("FileUID").filter(`FileUID/FileUID eq '${itemid}'`)();
    //  console.log(items , "items")
  }

  const truncateText = (text: string, maxLength: number) => {
    return text?.length > maxLength ? text.substring(0, maxLength) + "..." : text;
  };

  // start
  // const [remark,setRemark]=useState('');
  // console.log("remark",remark);
  const handleRemark = (event: any) => {
    event.preventDefault();
    event.stopPropagation();
    // console.log("value of remark",event.target.value)
    // setRemark(event.target.value);
    remark = event.target.value
    console.log(remark, "remaksss")
  }

  const handleLogAndLogHistory=async(event:any)=>{
    event.preventDefault(); 
    const buttonText = event.target.innerText || event.target.textContent;
    // const date =new Date();
    // const isoDate = date.toISOString();
    // console.log(isoDate);
    // const filterData=Mylistdata.find((item)=> item.CurrentUser === currentUserEmailRef.current || item.CurrentUser === props.actingforuseremail);
    // let filterData:any;
    // if(props.actingforuseremail !== '' && props.actingforuseremail !== null && props.actingforuseremail !== undefined){
    //   filterData=Mylistdata.find((item)=> item.CurrentUser === props.actingforuseremail);
    // }else{
    //   filterData=Mylistdata.find((item)=> item.CurrentUser === currentUserEmailRef.current);
    // }
    let filterData:any;
    if(props.actingforuseremail !== '' && props.actingforuseremail !== null && props.actingforuseremail !== undefined){
      filterData=Mylistdata.find((item)=> item.CurrentUser === props.actingforuseremail && item.Log === null);
    }else{
      filterData=Mylistdata.find((item)=> item.CurrentUser === currentUserEmailRef.current && item.Log === null);
    }
    console.log("filtered data Level",filterData.MasterApproval.Level);
    console.log("filterData id",filterData.FileUID.FileUID);
    console.log("filterData Id",filterData.Id);

    const isoDate = new Date().toISOString().slice(0, 19) + 'Z';
    // console.log(isoDate);
    // console.log("remark value",remark);

      // check and Set FinalApproved
    // let payload;
    let payload: any = {};
    // TASK CHANGE: Use the targetSiteUrl set during fetch to bypass 404
    const remoteWeb = targetSiteUrl ? Web(targetSiteUrl).using(AssignFrom(sp.web)) : sp.web;

    if(buttonText === "Approve"){
      try {
        const userConfirmed = await Swal.fire({
          title: 'Are you sure?',
          text: "Do you want to approve this File Request?",
          icon: 'warning',
          showCancelButton: true,
          confirmButtonColor: '#3085d6',
          cancelButtonColor: '#d33',
          confirmButtonText: 'Yes, approve it!',
          cancelButtonText: 'No, cancel',
        });
    
        if (!userConfirmed.isConfirmed) {
          console.log("User canceled the action.");
          return;
        }

        if(props.actingforuseremail !== '' && props.actingforuseremail !== null && props.actingforuseremail !== undefined){
          // alert(`deligate user`)
          // FIX: Query remoteWeb for all updates
          const updatedData = await remoteWeb.lists.getByTitle("DMSFileApprovalList").items
          .select("FileUID", "ID", "ApproveAction", "ApprovedLevel", "SiteName", "DocumentLibraryName", "ApprovedLevel" , "FilePreviewUrl")
          .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
          .catch((error) => console.error("Error fetching data from DMSFileApprovalList:", error));
          console.log(updatedData , "updatedData")
          if (updatedData && updatedData.length > 0) {
            const mydat = updatedData[0];
            filepreviewurl =  mydat?.FilePreviewUrl;
            approvedLevel = mydat.ApprovedLevel;
            const getTaskdata = await remoteWeb.lists.getByTitle("DMSFileApprovalTaskList").items.filter(`FileUID/FileUID eq '${filterData.FileUID.FileUID}'`)
            .select("FileUID/FileUID", "MasterApproval/ApprovalType", "CurrentUser", "Log")
            .expand("FileUID", "MasterApproval")()
            console.log(getTaskdata ,"getTaskdata");

            const getdatafromfoldermaster = await remoteWeb.lists.getByTitle("DMSFolderPermissionMaster").items
            .filter(`SiteName eq '${updatedData[0]?.SiteName}' and DocumentLibraryName eq '${updatedData[0]?.DocumentLibraryName}'`)()
            console.log(getdatafromfoldermaster , "getdatafromfoldermaster")
            let maxLevel = 0;
            getdatafromfoldermaster.forEach((item) => {
              if (item.Level > maxLevel) {
                maxLevel = item.Level;
              }
            })
            console.log("MaxLevel ", maxLevel);
            const Level = updatedData[0].ApprovedLevel;
            getTaskdata.forEach(async(item) => {
              if(item.Log === null){  
              if (item.CurrentUser === props.actingforuseremail){
                if (item.MasterApproval.ApprovalType === 0) {
                     approvedLevel =mydat.ApprovedLevel+1
                     setFinalStatus = "Approved";
                      // alert(`approval type 0 and its not a final level`);
                     if(Level >= maxLevel){
                      setFinalStatus="FinalApproved";
                      // alert(`approval type 0 and its  a final level`);
                      try {
                        await remoteWeb.lists.getByTitle("DMSFileApprovalList").items
                        .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
                        .then(async (items) => {
                            if (items.length > 0) {
                                const itemId = items[0].Id; // Assuming one item per FileUID
                                await remoteWeb.lists.getByTitle("DMSFileApprovalList").items.getById(itemId).update({
                                    Status: "Approved",
                                });
                                console.log("Updated DMSFileApprovalList with Approved status");
                            }
                        });
                        
                      } catch (error) {
                         console.log(error , "Error updating DMSFileApprovalList status");
                      }
                      try {
                       const siteName=updatedData[0].SiteName
                       const siteLookup = await remoteWeb.lists.getByTitle('MasterSiteURL').items.select("SiteURL").filter(`Title eq '${siteName}'`)();
                       if (siteLookup.length > 0) {
                        const fileWebContext = Web(siteLookup[0].SiteURL).using(AssignFrom(sp.web));
                        const file=  fileWebContext.getFileById(filterData.FileUID.FileUID);
                        const listItem = await file.getItem();        
                        const updatedDataLibrary =await listItem.update({
                          Status:"Approved"  
                        });
                        console.log("updatedData",updatedDataLibrary);
                       }
                      } catch (error) {
                        console.log(error , "Error updating status column on Libray or Folder");
                      }

                     try {
                      const filemasterlist = updatedData[0].SiteName;
                      const updateApprovedStatusstatus = await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.filter(`FileUID eq '${filterData.FileUID.FileUID}'`)();
                      
                    for (const item of updateApprovedStatusstatus) {
                      await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.getById(item.Id).update({
                        Status: "Approved",
                      });
                      console.log(`Item with ID ${item.Id} updated successfully.`);
                    }
                     } catch (error) {
                      console.log(`error in updateting file master list`,error);
                     }
                     }
                }else if (item.MasterApproval.ApprovalType === 1) {
                      let nonNullLogCount = 0;
                      let totalItems = getTaskdata.length;
                  //   alert(`approval type 1`);
                      getTaskdata.forEach(logItem => {
                        if (logItem.Log !== null) {
                          nonNullLogCount++;
                        }
                      });
                      if (nonNullLogCount >= totalItems - 1) {
                      //   alert(`final user to approve but not a final level`);
                        approvedLevel =mydat.ApprovedLevel+1
                        setFinalStatus = "Approved"
                        if(Level >= maxLevel){
                          // alert(`final user to approve and final level`);
                          setFinalStatus="FinalApproved";
                          approvedLevel=maxLevel+1;
                          try {
                            await remoteWeb.lists.getByTitle("DMSFileApprovalList").items
                            .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
                            .then(async (items) => {
                                if (items.length > 0) {
                                    const itemId = items[0].Id; // Assuming one item per FileUID
                                    await remoteWeb.lists.getByTitle("DMSFileApprovalList").items.getById(itemId).update({
                                        Status: "Approved",
                                    });
                                    console.log("Updated DMSFileApprovalList with Approved status");
                                }
                            });
                            
                          } catch (error) {
                             console.log(error , "Error updating DMSFileApprovalList status");
                          }
                          try {
                           const siteName=updatedData[0].SiteName
                           const siteLookup = await remoteWeb.lists.getByTitle('MasterSiteURL').items.select("SiteURL").filter(`Title eq '${siteName}'`)();
                           if (siteLookup.length > 0) {
                            const fileWebContext = Web(siteLookup[0].SiteURL).using(AssignFrom(sp.web));
                            const file=  fileWebContext.getFileById(filterData.FileUID.FileUID);
                            const listItem = await file.getItem();        
                            const updatedDataLibrary =await listItem.update({
                              Status:"Approved"  
                            });
                            console.log("updatedData",updatedDataLibrary);
                           }
                          } catch (error) {
                            console.log(error , "Error updating status column on Libray or Folder");
                          }
    
                         try {
                          const filemasterlist = updatedData[0].SiteName;
                          const updateApprovedStatusstatus = await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.filter(`FileUID eq '${filterData.FileUID.FileUID}'`)();
                          
                        for (const item of updateApprovedStatusstatus) {
                          await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.getById(item.Id).update({
                            Status: "Approved",
                          });
                          console.log(`Item with ID ${item.Id} updated successfully.`);
                        }
                         } catch (error) {
                          console.log(`error in updateting file master list`,error);
                         }
                         }
                      }else{
                        // thhis i have added
                          setFinalStatus = "Approved"
                          // alert(`not a final  user to approve and its not a final level`);
                          if(Level >= maxLevel){
                          //   alert(`not a final  user to approve but its a final level`);
                            approvedLevel=maxLevel+1;
                          }
                      }
                }
              }
              }
            })
          }
        }else{
          // alert('not a deligate user');
          const updatedData = await remoteWeb.lists.getByTitle("DMSFileApprovalList").items
          .select("FileUID", "ID", "ApproveAction", "ApprovedLevel", "SiteName", "DocumentLibraryName", "ApprovedLevel" , "FilePreviewUrl")
          .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
          .catch((error) => console.error("Error fetching data from DMSFileApprovalList:", error));
          console.log(updatedData , "updatedData")
          if (updatedData && updatedData.length > 0) {
            const mydat = updatedData[0];
            filepreviewurl =  mydat?.FilePreviewUrl;
            approvedLevel = mydat.ApprovedLevel;
            const getTaskdata = await remoteWeb.lists.getByTitle("DMSFileApprovalTaskList").items.filter(`FileUID/FileUID eq '${filterData.FileUID.FileUID}'`)
            .select("FileUID/FileUID", "MasterApproval/ApprovalType", "CurrentUser", "Log")
            .expand("FileUID", "MasterApproval")()
            console.log(getTaskdata ,"getTaskdata");

            const getdatafromfoldermaster = await remoteWeb.lists.getByTitle("DMSFolderPermissionMaster").items
            .filter(`SiteName eq '${updatedData[0]?.SiteName}' and DocumentLibraryName eq '${updatedData[0]?.DocumentLibraryName}'`)()
            console.log(getdatafromfoldermaster , "getdatafromfoldermaster")
            let maxLevel = 0;
            getdatafromfoldermaster.forEach((item) => {
              if (item.Level > maxLevel) {
                maxLevel = item.Level;
              }
            })
            console.log("MaxLevel ", maxLevel);
            const Level = updatedData[0].ApprovedLevel;
            getTaskdata.forEach(async(item) => {
              if(item.Log === null){  
              if (item.CurrentUser === currentUserEmailRef.current){
                if (item.MasterApproval.ApprovalType === 0) {
                     approvedLevel =mydat.ApprovedLevel+1
                     setFinalStatus = "Approved";
                      // alert(`approval type 0 and its not a final level`);
                     if(Level >= maxLevel){
                      setFinalStatus="FinalApproved";
                      // alert(`approval type 0 and its  a final level`);
                      try {
                        await remoteWeb.lists.getByTitle("DMSFileApprovalList").items
                        .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
                        .then(async (items) => {
                            if (items.length > 0) {
                                const itemId = items[0].Id; // Assuming one item per FileUID
                                await remoteWeb.lists.getByTitle("DMSFileApprovalList").items.getById(itemId).update({
                                    Status: "Approved",
                                });
                                console.log("Updated DMSFileApprovalList with Approved status");
                            }
                        });
                        
                      } catch (error) {
                         console.log(error , "Error updating DMSFileApprovalList status");
                      }
                      try {
                       const siteName=updatedData[0].SiteName
                       const siteLookup = await remoteWeb.lists.getByTitle('MasterSiteURL').items.select("SiteURL").filter(`Title eq '${siteName}'`)();
                       if (siteLookup.length > 0) {
                        const fileWebContext = Web(siteLookup[0].SiteURL).using(AssignFrom(sp.web));
                        const file=  fileWebContext.getFileById(filterData.FileUID.FileUID);
                        const listItem = await file.getItem();        
                        const updatedDataLibrary =await listItem.update({
                          Status:"Approved"  
                        });
                        console.log("updatedData",updatedDataLibrary);
                       }
                      } catch (error) {
                        console.log(error , "Error updating status column on Libray or Folder");
                      }

                     try {
                      const filemasterlist = updatedData[0].SiteName;
                      const updateApprovedStatusstatus = await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.filter(`FileUID eq '${filterData.FileUID.FileUID}'`)();
                      
                    for (const item of updateApprovedStatusstatus) {
                      await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.getById(item.Id).update({
                        Status: "Approved",
                      });
                      console.log(`Item with ID ${item.Id} updated successfully.`);
                    }
                     } catch (error) {
                      console.log(`error in updateting file master list`,error);
                     }
                     }
                }else if (item.MasterApproval.ApprovalType === 1) {
                      let nonNullLogCount = 0;
                      let totalItems = getTaskdata.length;
                  //   alert(`approval type 1`);
                      getTaskdata.forEach(logItem => {
                        if (logItem.Log !== null) {
                          nonNullLogCount++;
                        }
                      });
                      if (nonNullLogCount >= totalItems - 1) {
                      //   alert(`final user to approve but not a final level`);
                        approvedLevel =mydat.ApprovedLevel+1
                        setFinalStatus = "Approved"
                        if(Level >= maxLevel){
                          // alert(`final user to approve and final level`);
                          setFinalStatus="FinalApproved";
                          approvedLevel=maxLevel+1;
                          try {
                            await remoteWeb.lists.getByTitle("DMSFileApprovalList").items
                            .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
                            .then(async (items) => {
                                if (items.length > 0) {
                                    const itemId = items[0].Id; // Assuming one item per FileUID
                                    await remoteWeb.lists.getByTitle("DMSFileApprovalList").items.getById(itemId).update({
                                        Status: "Approved",
                                    });
                                    console.log("Updated DMSFileApprovalList with Approved status");
                                }
                            });
                            
                          } catch (error) {
                             console.log(error , "Error updating DMSFileApprovalList status");
                          }
                          try {
                           const siteName=updatedData[0].SiteName
                           const siteLookup = await remoteWeb.lists.getByTitle('MasterSiteURL').items.select("SiteURL").filter(`Title eq '${siteName}'`)();
                           if (siteLookup.length > 0) {
                            const fileWebContext = Web(siteLookup[0].SiteURL).using(AssignFrom(sp.web));
                            const file=  fileWebContext.getFileById(filterData.FileUID.FileUID);
                            const listItem = await file.getItem();        
                            const updatedDataLibrary =await listItem.update({
                              Status:"Approved"  
                            });
                            console.log("updatedData",updatedDataLibrary);
                           }
                          } catch (error) {
                            console.log(error , "Error updating status column on Libray or Folder");
                          }
    
                         try {
                          const filemasterlist = updatedData[0].SiteName;
                          const updateApprovedStatusstatus = await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.filter(`FileUID eq '${filterData.FileUID.FileUID}'`)();
                          
                        for (const item of updateApprovedStatusstatus) {
                          await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.getById(item.Id).update({
                            Status: "Approved",
                          });
                          console.log(`Item with ID ${item.Id} updated successfully.`);
                        }
                         } catch (error) {
                          console.log(`error in updateting file master list`,error);
                         }
                         }
                      }else{
                        // thhis i have added
                          setFinalStatus = "Approved"
                          // alert(`not a final  user to approve and its not a final level`);
                          if(Level >= maxLevel){
                          //   alert(`not a final  user to approve but its a final level`);
                            approvedLevel=maxLevel+1;
                          }
                      }
                }
              }
              }
            })
          }
        }

        payload={
                Log:setFinalStatus,
                LogHistory:isoDate,
                Remark:remark,
        }

        Swal.fire('Success', 'File Approved Successfully', 'success').then((result)=>{
          if(result.isConfirmed){
            window.location.reload();
          }
        });
      } catch (error) {
          console.log("Error Approving file");
          Swal.fire('Error', 'Error Approving file', 'error').then((result)=>{
            if(result.isConfirmed){
              window.location.reload();
            }
          });
      }

    }
 

    // here is reject case
    else if(buttonText === "Reject"){
      setFinalStatus = 'Rejected'
      try{
        const userConfirmed = await Swal.fire({
          title: 'Are you sure?',
          text: "Do you want to Reject this File Request?",
          icon: 'warning',
          showCancelButton: true,
          confirmButtonColor: '#3085d6',
          cancelButtonColor: '#d33',
          confirmButtonText: 'Yes, reject it!',
          cancelButtonText: 'No, cancel',
        });
    
        if (!userConfirmed.isConfirmed) {
          console.log("User canceled the action.");
          return;
        }
        setFinalStatus = 'Rejected'
        payload={
            Log:setFinalStatus,
            LogHistory:isoDate,
            Remark:remark,
            // ApprovedLevel:approvedLevel
        }
        const data=await remoteWeb.lists.getByTitle('DMSFileApprovalList').items.select("ID","ApproveAction","ApprovedLevel" , "SiteName").filter(` FileUID eq '${filterData.FileUID.FileUID}'`)();
      //   debugger
        console.log("data ",data);
      
        const id=data[0].Id; 
        // alert(` here is item id ${id} `)
        const filemasterlist = data[0].SiteName;
        // alert(` here is master list  ${filemasterlist} `)
        approvedLevel= data[0].ApprovedLevel
        // alert(` here is approved level ${approvedLevel} : type ${typeof(approvedLevel)}`)
        debugger
        try {
          // const updateRejectstatus1 = await sp.web.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.getById(id).update({ Status: 'Rejected' }); 
          // console.log("updateRejectstatus",updateRejectstatus1);
          const updateRejectstatus = await remoteWeb.lists
          .getByTitle(`DMS${filemasterlist}FileMaster`)
          .items.filter(`FileUID eq '${filterData.FileUID.FileUID}'`)();
      
        if (updateRejectstatus.length === 0) {
          console.log("No items found for the given FileUID:", filterData.FileUID.FileUID);
          return;
        }
  
        // Update the 'Status' column for each filtered item
        for (const item of updateRejectstatus) {
          await remoteWeb.lists.getByTitle(`DMS${filemasterlist}FileMaster`).items.getById(item.Id).update({
              Status: "Rejected",
          });

        const siteLookup = await remoteWeb.lists.getByTitle('MasterSiteURL').items.select("SiteURL").filter(`Title eq '${filemasterlist}'`)();
        if (siteLookup.length > 0) {
          const fileWebContext = Web(siteLookup[0].SiteURL).using(AssignFrom(sp.web));
          const file=  fileWebContext.getFileById(filterData.FileUID.FileUID);
          const listItem = await file.getItem();        
          const updatedData =await listItem.update({
            Status:"Rejected"  
          });
          console.log(`Item with ID ${item.Id} updated successfully.`);
        }
      }
  

        } catch (error) {
          Swal.fire(`'Error', 'Error Rejecting file 1', ${error}`).then((result)=>{
            if(result.isConfirmed){
              window.location.reload();
            }
          });
        }
   

      Swal.fire('Success', 'File Rejected Successfully', 'success').then((result)=>{
        if(result.isConfirmed){
          window.location.reload();
        }
      });;
    

      }catch{
        console.log("Error Rejecting file");
        Swal.fire('Error', 'Error Rejecting file', 'error').then((result)=>{
          if(result.isConfirmed){
            window.location.reload();
          }
        });
        
      }
      // setToggleLog((prevData)=>!prevData);
      // getApprovalmasterTasklist();
      // getCurrrentuser()
      Swal.fire('Success', 'File Rejected Successfully', 'success').then((result)=>{
        if(result.isConfirmed){
          window.location.reload();
        }
      });
      

    }
    else if(buttonText === "Rework"){
      try {
        const userConfirmed = await Swal.fire({
          title: 'Are you sure?',
          text: "Do you want to Rework this File Request?",
          icon: 'warning',
          showCancelButton: true,
          confirmButtonColor: '#3085d6',
          cancelButtonColor: '#d33',
          confirmButtonText: 'Yes, rework it!',
          cancelButtonText: 'No, cancel',
        });
    
        if (!userConfirmed.isConfirmed) {
          console.log("User canceled the action.");
          return;
        }
        setFinalStatus = 'Rework'
        payload={
            Log:setFinalStatus,
            LogHistory:isoDate,
            Remark:remark,
            // ApprovedLevel:approvedLevel
        }
        // setToggleLog((prevData)=>!prevData);
        // getApprovalmasterTasklist();
        Swal.fire('Success', 'File Rework Successfully', 'success').then((result)=>{
          if(result.isConfirmed){
            window.location.reload();
          }
        });
      } catch (error) {
        console.log("Error Reworking file");
        Swal.fire('Error', 'Error Reworking file', 'error').then((result)=>{
          if(result.isConfirmed){
            window.location.reload();
          }
        });

      }
  
      
    }

    console.log("payload for DMSFileApprovalTaskList",payload);
    
    const updateddata=await remoteWeb.lists.getByTitle("DMSFileApprovalTaskList").items.getById(filterData.Id).update(payload);
    
    console.log("Updated data",updateddata)
    if(buttonText === "Rework"){
      // alert(`Rework this file 1`)
      setFinalStatus = 'Rework'
      // alert(`this is SiteName ${filterData.FileUID.SiteName}`)
      // alert(`this is Filereqno ${filterData.FileUID.RequestNo}`)
      const updateStatusinMaster = await remoteWeb.lists.getByTitle(`DMS${filterData.FileUID.SiteName}FileMaster`).items.filter(`RequestNo eq '${filterData.FileUID.RequestNo}'`)()
      console.log(updateStatusinMaster , "updateStatusinMaster Rework")
      for (let item of updateStatusinMaster) { 
      //   alert(`Rework this file 2`)
        item.Status = 'Rework'; 
        await remoteWeb.lists.getByTitle(`DMS${filterData.FileUID.SiteName}FileMaster`).items.getById(item.ID).update({ Status: 'Rework' }); 
      }
      debugger
      const getTaskdata = await remoteWeb.lists.getByTitle("DMSFileApprovalTaskList").items.filter(`FileUID/FileUID eq '${filterData.FileUID.FileUID}' and LogHistory eq null`) 
      .select("*", "FileUID/FileUID", "MasterApproval/ApprovalType", "CurrentUser", "Log") 
      .expand("FileUID", "MasterApproval")();
     
  if (getTaskdata && getTaskdata.length > 0) {
      console.log(getTaskdata, "getTaskdatagetTaskdata");
      for (const item of getTaskdata) {
        // alert(item.ID)
          await remoteWeb.lists.getByTitle("DMSFileApprovalTaskList").items.getById(item.ID).delete();
      }
  } else {
      console.log("No items found to delete.");
  }
    }

    const data=await remoteWeb.lists.getByTitle('DMSFileApprovalList').items.select("ID","ApproveAction","ApprovedLevel").filter(` FileUID eq '${filterData.FileUID.FileUID}'`)();

    console.log("data ",data);
    const id=data[0].Id;
      if(buttonText === "Rework"){
      const paylaodForDMSFileApprovalList = {
        ApprovedLevel: 1,
        ApproveAction: "Rework",
        CurrentLevel:Number(data[0].ApprovedLevel),
        // Status:setFinalStatus,
        FilePreviewUrl: filepreviewurl
      }
      console.log("paylaodForDMSFileApprovalList", paylaodForDMSFileApprovalList);
  
      const updateddata1 = await remoteWeb.lists.getByTitle("DMSFileApprovalList").items.getById(id).update(paylaodForDMSFileApprovalList);
  
      console.log("updateddata1", updateddata1);
    }else{
      const paylaodForDMSFileApprovalList = {
        ApprovedLevel: Number(approvedLevel || data[0].ApprovedLevel),
        ApproveAction: payload.Log,
        CurrentLevel:Number(data[0].ApprovedLevel),
        // Status:setFinalStatus,
        FilePreviewUrl: filepreviewurl
      }
      console.log("paylaodForDMSFileApprovalList", paylaodForDMSFileApprovalList);
      if(buttonText === "Reject"){
        (paylaodForDMSFileApprovalList as any).Status = "Rejected";
      }
      const updateddata1 = await remoteWeb.lists.getByTitle("DMSFileApprovalList").items.getById(id).update(paylaodForDMSFileApprovalList);
  
      console.log("updateddata1", updateddata1);
    }

    setToggleLog((prevData)=>!prevData);
    // getApprovalmasterTasklist();
    getCurrrentuser()
   
}
  
  const openfileinedit = (e:any) =>{

      e.preventDefault();
      e.stopPropagation();
      setIsEditMode(true);
      const availreadurl:any = readablefilepreviewurl;
      console.log(availreadurl , "availreadurl");

      // Function to convert read-only URL to editable URL
      function convertToEditableUrl(availreadurl:any) {
          try {
              // Extract the 'id' parameter which contains the server-relative path
              const url = new URL(availreadurl);
              console.log(url , "url");
              const idParam = url.searchParams.get('id');
              
              if (idParam) {
                  // URL decode the parameter to get the actual path
                  const serverRelativePath = decodeURIComponent(idParam);
                  
                  // Construct the editable Office Online URL
                  const editableUrl = `https://officeindia.sharepoint.com/:w:/r/sites/Intranetdemos${encodeURIComponent(serverRelativePath.split("/sites/Intranetdemos")[1])}?web=1`;
                   console.log(editableUrl , "editableUrl");
                   const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
                   iframe.src = editableUrl;
                  return editableUrl;
              }
          } catch (error) {
              console.error("Error converting URL:", error);
          }
          
          return null;
      }
      
      const editableUrl = convertToEditableUrl(availreadurl);
      console.log(editableUrl);
      // Result: https://officeindia.sharepoint.com/:w:/r/sites/Intranetdemos/Location/Section/Transmittal_1758641391232.docx?web=1
  }
 
    const handleSave = () => {
      console.log("Saving file");
      const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
                   iframe.src = readablefilepreviewurl;
      
      setIsEditMode(false);
    };
  
  return (

    <div>
      {activeComponent === "" ? (
        <div style={{ float: 'left', width: '100%', clear: 'both' }}>
          <div>
            <div className="content">


              <div className="row">
                <div className="col-12">
<div style={{ display: "flex", gap: "20px", marginTop: "20px" }}>
  {/* Basic Information - Left Side */}
  <div 
    className="" 
    style={{ 
      backgroundColor: 'white',
      borderRadius: '5px', 
      padding: '15px',
      flex: "0 0 400px"  // Fixed width for basic info
    }}
  >
    <div
      style={{
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center"
      }}
    >
      <h3 className="text-dark fw-bold m-0 font-16 mb-1 bg-basic">Basic Information</h3>
    </div>
    <div id="dynamicDetailsContainer"></div>
  </div>

  {/* AI Summary - Right Side */}

    <div
      style={{
        backgroundColor: "#ffffff",
        borderRadius: "5px",
        padding: "20px",
        position: "relative",
        boxShadow: "0 2px 10px rgba(0,0,0,0.15)",
        flex: "1"  // Takes remaining space
      }}
    >
     
<div className="d-flex align-items-center justify-content-end gap-1">
      <h5 style={{ marginBottom: "10px", margin:'0px' }}>AI Suggestion</h5>
<div
  style={{
    width: "40px",
    height: "22px",
    background: openfileon ? "#0d6efd" : "#ddd",
    borderRadius: "11px",
    position: "relative",
    cursor: "pointer"
  }}
  onClick={handleToggleClick}
>
  <div
    style={{
      width: "18px",
      height: "18px",
      background: "#fff",
      borderRadius: "50%",
      position: "absolute",
      top: "2px",
      left: openfileon ? "20px" : "2px",
      transition: "left 0.3s ease"
    }}
  />
</div>
</div>
      <div>
        {aiSummary ? aiSummary : "No AI Summary available."}
      </div>
    </div>

</div>
                  {/* <div className="" style={{ backgroundColor: 'white',  marginTop: '20px', borderRadius: '5px', padding: '15px' }}>
                    
                    <div
style={{
  display: "flex",
  justifyContent: "space-between",
  alignItems: "center"
}}
>
<h3 className="text-dark font-16 mb-1">Basic Information</h3>

<button
  type="button"
  style={{
    padding: "6px 12px",
    borderRadius: "4px",
    border: "none",
    backgroundColor: "#0d6efd",
    color: "#ffffff",
    fontSize: "13px",
    fontWeight: 600,
    cursor: "pointer"
  }}
  onClick={() => setShowAISummary(true)}
>
  AI Summary
</button>
</div>
                    <div id="dynamicDetailsContainer"></div>
                  </div> */}
                  {openfileon && (
  <div className="" style={{ backgroundColor: 'white', marginTop: '20px', borderRadius: '5px', padding: '15px' }}>
    <div style={{display:'flex', gap:'5px', justifyContent:'end', marginBottom:'5px'}}>
      <button
        type="button"
        style={{
          padding: "10px 15px",
          border: "none",
          borderRadius: "4px",
          cursor: "pointer",
          fontWeight: "600",
          backgroundColor: "#6c757d",
          color: "white",
          position: isFullScreen ? "fixed" : "static",
          top: isFullScreen ? "10px" : "auto",
          right: isFullScreen ? "10px" : "auto",
          zIndex: 10000,
          transition: "background-color 0.3s ease",
        }}
        onMouseOver={(e) => {
          const target = e.target as HTMLButtonElement;
          target.style.backgroundColor = "#5a6268";
        }}
        onMouseOut={(e) => {
          const target = e.target as HTMLButtonElement;
          target.style.backgroundColor = "#6c757d";
        }}
        onClick={() => setIsFullScreen(!isFullScreen)}
      >
        {isFullScreen ? "Exit Full Screen" : "Full Screen"}
      </button>
    </div>

    {/* <div style={{textAlign:'center'}} className='spinner' id="spinner">
      <img style={{width:'116px', margin: '0px auto'}} src={require("../assets/ESSAROLLER.gif")} alt="Loading..." />
      <div style={{color:"black", marginBottom:'10px'}}>Loading Preview File</div>
    </div> */}

    <iframe
      id="filePreview"
      src={editableUrl || ""}
      style={{
        width: isFullScreen ? "100vw" : "100%",
        height: isFullScreen ? "100vh" : "1200px",
        border: "none",
        position: isFullScreen ? "fixed" : "relative",
        top: isFullScreen ? 0 : "auto",
        left: isFullScreen ? 0 : "auto",
        zIndex: isFullScreen ? 9999 : "auto",
        background: "#fff",
      }}
      title="File Preview"
    />
  </div>
)}
                  {/* {showAISummary && (
<div
  style={{
    backgroundColor: "#ffffff",
    marginTop: "20px",
    borderRadius: "5px",
    padding: "20px",
    position: "relative",
    boxShadow: "0 2px 10px rgba(0,0,0,0.15)"
  }}
>
  <button
    onClick={() => setShowAISummary(false)}
    style={{
      position: "absolute",
      top: "10px",
      right: "10px",
      border: "none",
      background: "transparent",
      fontSize: "18px",
      cursor: "pointer"
    }}
  >
    ✕
  </button>

  <h5 style={{ marginBottom: "10px" }}>AI Summary</h5>

  <div>
    {aiSummary ? aiSummary : "No AI Summary available."}
  </div>
</div>
)} */}
                  {toggleLog && (
                    <div className="" style={{ backgroundColor: 'white',  marginTop: '20px', borderRadius: '5px', padding: '15px' }}>
                      
                      <div className="">
                        <div className="">

                          <div className="row mt-3">

                            <div className="col-lg-12">
                              <div className="mb-0">
                                <label className="form-label text-dark font-14">
                                  Remarks:
                                </label>
                                <input
                                  type="text" style={{ height: '70px' }} className="form-control"
                                  onChange={handleRemark}
                                //  value={remark}
                                />
                              </div>
                            </div>
                          </div>

                          <div className="row mt-3">
                            <div className="col-12 text-center">
                              <a >
                                {" "}
                                <button

                                  onClick={handleLogAndLogHistory}
                                  type="button"
                                  className="btn btn-success waves-effect waves-light m-1"
                                >
                                  <i className="fe-check-circle me-1"></i>{" "}
                                  Approve
                                </button>
                              </a>
                              <a >
                                {" "}
                                <button

                                  onClick={handleLogAndLogHistory}
                                  type="button"
                                  className="btn btn-warning waves-effect waves-light m-1"
                                >
                                  <i className="fe-check-circle me-1"></i>{" "}
                                  Rework
                                </button>
                              </a>
                              {/* <a >  <button type="button" className="btn btn-warning waves-effect waves-light m-1"><i className="fe-corner-up-left me-1"></i> Rework</button></a>   */}
                              <a >
                                {" "}
                                <button
                                  onClick={handleLogAndLogHistory}
                                  type="button"
                                  className="btn btn-danger waves-effect waves-light m-1"
                                >
                                  <i className="fe-x-circle me-1"></i>{" "}
                                  Reject
                                </button>
                              </a>
                              {/* <button
                                    type="button"
                                    className="btn btn-light waves-effect waves-light m-1"
                                  >
                                    <i className="fe-x me-1"></i> Cancel
                                  </button> */}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                  )}
                  <div>
                    <div className="DMSMasterContainer">
                      {/* <h4 className="page-title fw-bold mb-1 font-20">Settings</h4> */}
                      <div className="" style={{ backgroundColor: 'white',  marginTop: '20px', borderRadius: '5px', padding: '15px' }}>
                        <table style={{width:'100%'}} className="mtbalenew">
                          <thead >
                            <tr style={{width:'100%',display:'table'}} >
                              <th
                                style={{
                                  minWidth: '55px',
                                  maxWidth: '55px',

                                }}
                              >
                                S.No
                              </th>
                              <th style={{ minWidth: '60px', maxWidth: '60px' }}>Level</th>
                              {/* <th style={{ minWidth: '120px', maxWidth: '120px' }}>Process Name</th> */}
                              <th style={{ minWidth: '110px', maxWidth: '110px' }}>Assigned To</th>
                              <th style={{ minWidth: '110px', maxWidth: '110px' }}>Requester Name</th>
                              <th style={{ minWidth: '110px', maxWidth: '110px' }}>Requested Date</th>
                              <th style={{ minWidth: '110px', maxWidth: '110px' }}>Action Taken By</th>
                              <th style={{ minWidth: '110px', maxWidth: '110px' }}>Action Taken On</th>
                              <th
                                style={{
                                  minWidth: '70px',
                                  maxWidth: '70px',

                                }}
                              >
                                Remark
                              </th>
                              <th
                                style={{
                                  minWidth: '70px',
                                  maxWidth: '70px',

                                }}
                              >
                                Status
                              </th>

                            </tr>
                          </thead>
                          <tbody style={{ maxHeight: '8007px' }}>

                            {Mylistdata.length > 0 ? Mylistdata.map((item, index) => {
                              return (
                              <tr style={{width:'100%',display:'table'}} >
                                  <td style={{ minWidth: '55px', maxWidth: '55px' }}>
                                    <span style={{ marginLeft: '0px' }} className="indexdesign">
                                      {index + 1}</span></td>
                                  <td style={{ minWidth: '60px', maxWidth: '60px' }}>
                                    {/* {(truncateText(item?.FileUID?.Requ110pxestNo, 20))} */}
                                    {item?.MasterApproval?.Level}
                                  </td>
                                  <td style={{ minWidth: '110px', maxWidth: '110px' }}>{
                                    item?.AssignedToTitle
                                  }</td>
                                  <td style={{ minWidth: '110px', maxWidth: '110px' }}>


                                    {item.RequestedByTitle}

                                  </td>
                                  <td style={{ minWidth: '110px', maxWidth: '110px', textAlign: 'center' }}>
                                    <div className="btn btn-light1">{new Date(item?.FileUID?.Created).toLocaleString('en-US', {
                                      month: '2-digit',
                                      day: '2-digit',
                                      year: 'numeric',
                                      // hour: '2-digit',
                                      // minute: '2-digit',
                                      // second: '2-digit',
                                      // hour12: true 
                                    })}
                                    </div>
                                  </td>
                                  <td style={{ minWidth: '110px', maxWidth: '110px', textAlign: 'center' }}>

                                    {item?.AssignedToTitle}


                                  </td>
                                  <td style={{ minWidth: '110px', maxWidth: '110px' }}>

                                    {/* {moment(item?.LogHistory).format("DD-MMM-YYYY")} */}
                                    {item?.LogHistory ? item.LogHistory : ""}


                                  </td>


                                  <td style={{ minWidth: '70px', maxWidth: '70px' }}>
                                    {item.Remark}
                                  </td>
                                  <td style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>
                                    {/* <div className="finish mb-0"></div>{item.FileUID.Status} */}
                                    <div className="finish mb-0">  {item.Log !=""?item.Log :"Pending" } </div>
                                  </td>
                                </tr>
                              )
                            })
                              : ""
                            }
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div> 
        </div>
      ) : (
        <div>
          {activeComponent === "Create Entity" && (
            <div>
              <button onClick={() => handleReturnToMain("")}>
                {" "}
                My Approvals{" "}
              </button>
              {/* <DMSMyApproval /> */}
            </div>
          )}
        </div>
      )}
    </div>

  );
};




export default DMSMyApprovalAction;