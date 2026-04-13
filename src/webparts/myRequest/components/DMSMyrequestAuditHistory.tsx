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
import "@pnp/sp/presets/all"
import { PermissionKind } from "@pnp/sp/security";
//import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
// import "../../verticalSideBar/components/VerticalSidebar.scss";
// import "./dmscss";
// import "./DMSAdmincss"
import { useState, useRef, useEffect } from "react";
// import {IDmsapprovalProps} from './IDmsapprovalProps'
import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";
//  import EntityMapping from "./EntityMapping";

import './ApprovalActioncss'

import { faEye } from "@fortawesome/free-regular-svg-icons";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
// import DMSMyApproval from "./MyApprovals";
let approvedLevel: any = ''
let filepreviewurl = ''
let remark: any = ''
let Level: any = ''
let setFinalStatus: any = ''
let FileUID: any = ''
const DMSMyrequestaudithistory = ({ props }: any) => {
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
  const [Mylistdata, setMylistdata] = useState([]);
  ////
  const [storedUserInfo, setStoredUserInfo] = useState(null);
  const [ApprovedStatus, setApprovedStatus] = useState('');  // State for ApprovalType 0
  // const [approvedLevel, setApprovedLevel] = useState<number>();

  const [activeComponent, setActiveComponent] = useState<string>('');
  ////
  console.log(activeComponent, "activeComponent")
  const handleReturnToMain = (Name: any) => {
    setActiveComponent(Name); // Reset to show the main component
    console.log(activeComponent, "activeComponent updated")
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
  const getApprovalmasterTasklist = async () => {
    //   alert(`DMSFileApprovalTaskList ${props}`)
    try {
      const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
        "*",
        "Log", "CurrentUser", "Remark"
        , "LogHistory", "ID"
        , "FileUID/FileUID"
        , "FileUID/SiteName"
        , "FileUID/DocumentLibraryName"
        , "FileUID/FileName"
        , "FileUID/RequestNo"
        //  ,"FileUID/FilePreviewUrl" 
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
        .filter(`FileUID/FileUID eq '${props}'`)
        .orderBy("Modified", false)();
      console.log(items, "DMSFileApprovalTaskList");


      // Fetch user titles
      const updatedItems = await Promise.all(items.map(async (item) => {
        const userTitle = await getUserTitleByEmail(item.FileUID.RequestedBy);
        const assignedtouserTitle = await getUserTitleByEmail(item.CurrentUser);
        // alert(userTitle)
        return { ...item, RequestedByTitle: userTitle, assignedtouserTitle: assignedtouserTitle };
      }));




      setMylistdata(updatedItems);
      console.log(Mylistdata, "Mylistdata")

    } catch (error) {
      console.error("Error fetching list items:", error);
    }
    try {
      console.log("here")
      const updatedData: any = await sp.web.lists.getByTitle("DMSFileApprovalList").items
        .select("FileUID", "ID", "ApproveAction", "ApprovedLevel", "SiteName", "DocumentLibraryName", "ApprovedLevel", "FilePreviewUrl")
        .filter(`FileUID eq '${FileUID}'`)()
        .catch((error) => console.error("Error fetching data from DMSFileApprovalList:", error));
      console.log(updatedData, "updatedData")
      filepreviewurl = updatedData[0]?.FilePreviewUrl;
      console.log(filepreviewurl, "file url")
    } catch (error) {
      console.error("Error fetching list items:", error);
    }
  };
  // const  getUserTitleByEmail = async (userEmail:any) => {
  //   try {
  //     // Get user by email
  //     const user = await sp.web.siteUsers.getByEmail(userEmail)();
  //     // Extract user title
  //     const userName = user.Title;
  //     // Return user title
  //     return userName;
  //   } catch (error) {
  //     console.error("Error fetching user title:", error);
  //     return null;
  //   }
  // }
  const Mynewdata = async () => {
    const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items
      .filter("FileUID/RequestNo eq 'e274710c-f990-46bb-86e7-934a91251cfa'")
      ();

    console.log(items, "Filter Query");

  }
  console.log(Mylistdata, "Mylistdata")
  // start
  const [toggleLog, setToggleLog] = useState(false);
  // end

  const currentUserEmailRef = useRef('');
  const getCurrrentuser = async () => {
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    getApprovalmasterTasklist();
    Mynewdata()
  }
  useEffect(() => {
    getCurrrentuser()

  }, []);




  const getTaskItemsbyID = async (e: any, itemid: any) => {
    console.log("itemid", itemid)
    const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select("CurrentUser", "FileUID/FileUID", "Log").expand("FileUID").filter(`FileUID/FileUID eq '${itemid}'`)();
    //  console.log(items , "items")
  }

  const truncateText = (text: string, maxLength: number) => {
    return text?.length > maxLength ? text?.substring(0, maxLength) + "..." : text;
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

  const handleLogAndLogHistory = async (event: any) => {
    event.preventDefault();
    const buttonText = event.target.innerText || event.target.textContent;
    // const date =new Date();
    // const isoDate = date.toISOString();
    // console.log(isoDate);
    const filterData = Mylistdata.find((item) => item.CurrentUser === currentUserEmailRef.current)
    console.log("filtered data Level", filterData.MasterApproval.Level);
    console.log("filterData id", filterData.FileUID.FileUID);
    console.log("filterData Id", filterData.Id);

    const isoDate = new Date().toISOString().slice(0, 19) + 'Z';
    // console.log(isoDate);
    // console.log("remark value",remark);

    // check and Set FinalApproved



    let payload;
    if (buttonText === "Approve") {


      const updatedData = await sp.web.lists.getByTitle("DMSFileApprovalList").items
        .select("FileUID", "ID", "ApproveAction", "ApprovedLevel", "SiteName", "DocumentLibraryName", "ApprovedLevel", "FilePreviewUrl")
        .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
        .catch((error) => console.error("Error fetching data from DMSFileApprovalList:", error));
      console.log(updatedData, "updatedData")

      if (updatedData && updatedData.length > 0) {
        const mydat = updatedData[0]; // Assuming you want to compare using the first item's SiteName
        filepreviewurl = mydat?.FilePreviewUrl
        console.log(mydat?.FilePreviewUrl, "items,,,,")
        approvedLevel = mydat.ApprovedLevel
        // Step 3: Fetch data from the second list where SiteName matches
        const updatedata2 = await sp.web.lists.getByTitle("DMSFolderPermissionMaster").items
          .filter(`SiteName eq '${mydat.SiteName}' and DocumentLibraryName eq '${mydat.DocumentLibraryName}' and Level eq ${mydat.ApprovedLevel}`)()
          .catch((error: any) => console.error("Error fetching data from DMSFolderPermissionMaster:", error));

        const getTaskdata = await sp.web.lists.getByTitle("DMSFileApprovalTaskList").items.filter(`FileUID/FileUID eq '${filterData.FileUID.FileUID}'`)
          .select("FileUID/FileUID", "MasterApproval/ApprovalType", "CurrentUser", "Log")
          .expand("FileUID", "MasterApproval")()
        console.log(getTaskdata, "getTaskdata")
        console.log(updatedata2, "here is my data");

        getTaskdata.forEach(item => {
          console.log(item.CurrentUser, "CurrentUser")
          // Step 1: Check if CurrentUser matches the stored user
          if (item.CurrentUser === currentUserEmailRef.current) {
            // Step 2: Check ApprovalType
            if (item.MasterApproval.ApprovalType === 0) {
              console.log(approvedLevel, "approvedLevel first in 0")
              // If ApprovalType is 0, set state to 'done'
              setApprovedStatus('Approved');
              console.log("entere here in 0 for approval level", filterData.FileUID.FileUID, mydat.ApprovedLevel)
              console.log("entere here in 0", ApprovedStatus)
              approvedLevel = mydat.ApprovedLevel + 1
              setFinalStatus = "Approved"
              console.log(approvedLevel, "approvedLevel second in 1")
            } else if (item.MasterApproval.ApprovalType === 1) {
              // Step 3: If ApprovalType is 1, check the Log field
              let nonNullLogCount = 0;
              let totalItems = getTaskdata.length;

              getTaskdata.forEach(logItem => {
                if (logItem.Log !== null) {
                  nonNullLogCount++;
                }
              });

              // If more than 5 out of 6 Logs are not null, set 'approvalInProgress'
              if (nonNullLogCount >= totalItems - 1) {
                console.log(approvedLevel, "approvedLevel first in 1")
                approvedLevel = mydat.ApprovedLevel + 1
                setFinalStatus = "Approved"
                setApprovedStatus('Approved');
                console.log("entere here in 1 for approval level", filterData.FileUID.FileUID, mydat.ApprovedLevel + 1)
                console.log("entere here in 1", ApprovedStatus)
                console.log(approvedLevel, "approvedLevel second in 1")
              }
            }
          }
        });

      } else {
        console.log("No matching data found in DMSFileApprovalList.");
      }
      //start
      try {
        const updatedData1: any = await sp.web.lists.getByTitle("DMSFileApprovalList").items
          .select("FileUID", "ID", "ApproveAction", "ApprovedLevel", "SiteName", "DocumentLibraryName", "ApprovedLevel", "FilePreviewUrl")
          .filter(`FileUID eq '${FileUID}'`)()
          .catch((error) => console.error("Error fetching data from DMSFileApprovalList:", error));
        console.log(updatedData1, "updatedData")


        filepreviewurl = updatedData1[0]?.FilePreviewUrl;
        Level = updatedData1[0].ApprovedLevel;
        console.log(updatedData1[0], "DocumentLibraryName")

        const getdatafromfoldermaster = await sp.web.lists.getByTitle("DMSFolderPermissionMaster").items
          .filter(`SiteName eq '${updatedData1[0].SiteName}' and DocumentLibraryName eq '${updatedData1[0].DocumentLibraryName}'`)()
        console.log(getdatafromfoldermaster, "getdatafromfoldermaster")

        let maxLevel = 0;
        getdatafromfoldermaster.forEach((item) => {
          if (item.Level >= maxLevel) {
            maxLevel = item.Level;
          }
        })

        console.log("MaxLevel ", maxLevel);

        if (Level === maxLevel) {

          const taskdata = await sp.web.lists.getByTitle("DMSFileApprovalTaskList").items.filter(`FileUID/FileUID eq '${filterData.FileUID.FileUID}'`)
            .select("FileUID/FileUID", "MasterApproval/ApprovalType", "CurrentUser", "Log")
            .expand("FileUID", "MasterApproval")()

          console.log("getData from DMSFileApprovalTaskList", taskdata);

          taskdata.forEach(async (item) => {
            if (item.CurrentUser === currentUserEmailRef.current) {

              if (item.MasterApproval.ApprovalType === 0) {
                setFinalStatus = "FinalApproved";
                try {
                  await sp.web.lists.getByTitle("DMSFileApprovalList").items
                    .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
                    .then(async (items) => {
                      if (items.length > 0) {
                        const itemId = items[0].Id; // Assuming one item per FileUID
                        // alert(`${itemId} item id is 1`)
                        await sp.web.lists.getByTitle("DMSFileApprovalList").items.getById(itemId).update({
                          Status: "Approved",
                        });
                        console.log("Updated DMSFileApprovalList with Approved status");
                        // alert(`${itemId} Updated DMSFileApprovalList with Approved status`)
                      }
                    });
                } catch (error) {
                  console.log(error, "Error updating DMSFileApprovalList status");
                }
                try {
                  // Update Column Status on Document library or folder
                  // updatedData1[0].DocumentLibraryName
                  // FileUID
                  // New code start
                  const siteName = updatedData1[0].SiteName
                  // console.log("siteName",siteName);
                  const subsite = await sp.web.webs.filter(`Title eq '${siteName}'`)();
                  // console.log(subsite , "subsite");
                  // console.log("subsite id",subsite[0].Id)

                  const { web } = await sp.site.openWebById(subsite[0].Id)
                  // end
                  // get the details of the file present inside the document library
                  const file = web.getFileById(filterData.FileUID.FileUID);
                  const listItem = await file.getItem();
                  const updatedData = await listItem.update({
                    Status: "Approved"
                  });
                  console.log("updatedData", updatedData);
                } catch (error) {
                  console.log(error, "Error updating status column on Libray or Folder");
                }


              } else if (item.MasterApproval.ApprovalType === 1) {

                let approvedUser = 0;
                let numberOfUser = taskdata.length;

                taskdata.forEach(logItem => {
                  if (logItem.Log !== null) {
                    approvedUser++;
                  }
                });

                if (approvedUser >= numberOfUser - 1) {
                  setFinalStatus = "FinalApproved";
                  try {
                    await sp.web.lists.getByTitle("DMSFileApprovalList").items
                      .filter(`FileUID eq '${filterData.FileUID.FileUID}'`)()
                      .then(async (items) => {
                        if (items.length > 0) {
                          const itemId = items[0].Id; // Assuming one item per FileUID
                          // alert(`${itemId} item id is 2`)
                          await sp.web.lists.getByTitle("DMSFileApprovalList").items.getById(itemId).update({
                            Status: "Approved",
                          });
                          console.log("Updated DMSFileApprovalList with Approved status");
                          // alert(`${itemId} Updated DMSFileApprovalList with Approved status`)
                        }
                      });
                  } catch (error) {
                    console.log(error, "Error updating DMSFileApprovalList status");
                  }
                  try {
                    // Update Column Status on Document library or folder
                    // updatedData1[0].DocumentLibraryName
                    // FileUID
                    // New code start
                    const siteName = updatedData1[0].SiteName
                    // console.log("siteName",siteName);
                    const subsite = await sp.web.webs.filter(`Title eq '${siteName}'`)();
                    // console.log(subsite , "subsite");
                    // console.log("subsite id",subsite[0].Id)

                    const { web } = await sp.site.openWebById(subsite[0].Id)
                    // end
                    // get the details of the file present inside the document library
                    const file = web.getFileById(filterData.FileUID.FileUID);
                    const listItem = await file.getItem();
                    const updatedData = await listItem.update({
                      Status: "Approved"
                    });
                    console.log("updatedData", updatedData);
                  } catch (error) {
                    console.log(error, "Error updating status column on Libray or Folder");
                  }

                }

              }
            }

          })

        } else {
          setFinalStatus = "Approved";
          console.log("Level is not equal to max level", Level);
          console.log("FinalStatus", setFinalStatus);


        }

      } catch (error) {
        console.error("Error fetching list items:", error);
      }
      // end

      payload = {
        Log: setFinalStatus,
        LogHistory: isoDate,
        Remark: remark,
        // ApprovedLevel:approvedLevel
      }
    }
    else if (buttonText === "Reject") {
      setFinalStatus = 'Rejected'
      payload = {
        Log: setFinalStatus,
        LogHistory: isoDate,
        Remark: remark,
        // ApprovedLevel:approvedLevel
      }
    } else if (buttonText === "Rework") {
      setFinalStatus = 'Rework'
      payload = {
        Log: setFinalStatus,
        LogHistory: isoDate,
        Remark: remark,
        // ApprovedLevel:approvedLevel
      }

    }

    console.log("payload for DMSFileApprovalTaskList", payload);

    const updateddata = await sp.web.lists.getByTitle("DMSFileApprovalTaskList").items.getById(filterData.Id).update(payload);

    console.log("Updated data", updateddata)
    if (buttonText === "Rework") {
      setFinalStatus = 'Rework'
      //   alert(`this is SiteName ${filterData.FileUID.SiteName}`)
      //   alert(`this is Filereqno ${filterData.FileUID.RequestNo}`)
      const updateStatusinMaster = await sp.web.lists.getByTitle(`DMS${filterData.FileUID.SiteName}FileMaster`).items.filter(`RequestNo eq '${filterData.FileUID.RequestNo}'`)()
      console.log(updateStatusinMaster, "updateStatusinMaster")
      for (let item of updateStatusinMaster) {
        item.Status = 'Rework';
        await sp.web.lists.getByTitle(`DMS${filterData.FileUID.SiteName}FileMaster`).items.getById(item.ID).update({ Status: 'Rework' });
      }
      debugger
      const getTaskdata = await sp.web.lists.getByTitle("DMSFileApprovalTaskList").items.filter(`FileUID/FileUID eq '${filterData.FileUID.FileUID}' and Log eq null`)
        .select("*", "FileUID/FileUID", "MasterApproval/ApprovalType", "CurrentUser", "Log")
        .expand("FileUID", "MasterApproval")();

      if (getTaskdata && getTaskdata.length > 0) {
        console.log(getTaskdata, "getTaskdatagetTaskdata");
        for (const item of getTaskdata) {
          // alert(item.ID)
          await sp.web.lists.getByTitle("DMSFileApprovalTaskList").items.getById(item.ID).delete();
        }
      } else {
        console.log("No items found to delete.");
      }
    }

    const data = await sp.web.lists.getByTitle('DMSFileApprovalList').items.select("ID", "ApproveAction", "ApprovedLevel").filter(` FileUID eq '${filterData.FileUID.FileUID}'`)();

    console.log("data ", data);
    const id = data[0].Id;

    const paylaodForDMSFileApprovalList = {
      ApprovedLevel: approvedLevel,
      ApproveAction: payload.Log,
      // Status:setFinalStatus,
      FilePreviewUrl: filepreviewurl
    }

    console.log("paylaodForDMSFileApprovalList", paylaodForDMSFileApprovalList);

    const updateddata1 = await sp.web.lists.getByTitle("DMSFileApprovalList").items.getById(id).update(paylaodForDMSFileApprovalList);

    console.log("updateddata1", updateddata1);

  }
  const iframe = document.getElementById("filePreview") as HTMLIFrameElement;
  // const spinner = document.getElementById("spinner") as HTMLElement;

  // Show the spinner and hide the iframe initially
  // spinner.style.display = "block";
  if (iframe) {
    iframe.style.display = "none";
    iframe.src = filepreviewurl;
  }


  // Add an onload event listener to the iframe
  if (iframe) {

    iframe.onload = () => {
      console.log("Iframe has loaded");

      const checkAndHideButton = () => {
        try {
          const iframeDocument = iframe.contentDocument || iframe.contentWindow?.document;
          if (iframeDocument) {
            const button = iframeDocument.getElementById("OneUpCommandBar") as HTMLElement;
            const excelToolbar = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
            if (excelToolbar) {
              excelToolbar.style.display = "none"
            }
            if (button) {
              console.log("Hiding the OneUpCommandBar element");
              button.style.display = "none";

              // Hide the spinner and show the iframe after the button is hidden
              // spinner.style.display = "none";
              iframe.style.display = "block";

              // Exit the loop once the button is found and hidden
            } else {
              console.log("OneUpCommandBar not found, rechecking...");
            }

            const helpbutton = iframeDocument.getElementById("m_excelEmbedRenderer_m_ewaEmbedViewerBar") as HTMLElement;
            if (helpbutton) {
              helpbutton.style.display = "none"
            }
          }
        } catch (error) {
          console.error("Error accessing iframe content:", error);
        }

        // Re-check after a short delay if the button wasn't found
        setTimeout(checkAndHideButton, 100);
      };

      // Start checking for the button
      checkAndHideButton();
    };
  }

  return (

    <div>
      {activeComponent === "" ? (
        <div style={{ float: 'left', width: '100%', clear: 'both' }}>
          <div>
            <div className="content">


              <div className="row">
                <div className="col-12">

                  <div>
                    <div className="DMSMasterContainer">
                      {/* <h4 className="page-title fw-bold mb-1 font-20">Settings</h4> */}
                      <div className="popsn" style={{ backgroundColor: 'white', border: '1px solid #54ade0', marginTop: '20px', borderRadius: '20px', padding: '15px' }}>
                        <table className="mtbalenew" style={{width:'100%'}}>
                          <thead >
                              <tr style={{width:'100%',display:'table'}}>
                              <th
                                style={{minWidth: '40px', maxWidth: '40px', }}
                              >
                                S.No
                              </th>
                              <th   style={{minWidth: '60px', maxWidth: '60px', }}>Level</th>
                              {/* <th style={{ minWidth: '120px', maxWidth: '120px' }}>Process Name</th> */}
                              <th   style={{minWidth: '110px', maxWidth: '110px', }}>Assigned To</th>

                              <th   style={{minWidth: '110px', maxWidth: '110px', }}>Requester Name</th>
                              <th style={{ minWidth: '110px', maxWidth: '110px' }}>Requested Date</th>
                              <th style={{ minWidth: '110px', maxWidth: '110px' }}>Action Taken By</th>
                              <th style={{ minWidth: '110px', maxWidth: '110px' }}>Action Taken On</th>
                              <th
                                style={{
                                  minWidth: '110px',
                                  maxWidth: '110px',

                                }}
                              >
                                Remark
                              </th>
                              <th style={{ minWidth: '70px', maxWidth: '70px' }}>Status</th>

                            </tr>
                          </thead>
                          <tbody style={{ maxHeight: '8007px' }}>

                            {Mylistdata.length > 0 ? Mylistdata.map((item, index) => {
                              return (
                                  <tr style={{width:'100%',display:'table'}}>
                                  <td style={{ minWidth: '40px', maxWidth: '40px' }}>
                                    <span style={{ marginLeft: '0px' }} className="indexdesign">
                                      {index + 1}</span></td>
                                  {/* <td>{(truncateText(item.FileUID.FileUID, 20))}
                                                      </td> */}
                                  <td   style={{minWidth: '60px', maxWidth: '60px', }}>Level {
                                    item.MasterApproval.Level
                                  }</td>
                                  <td   style={{minWidth: '110px', maxWidth: '110px', }}>


                                    {item.assignedtouserTitle}

                                    {/* {getUserTitleByEmail(item?.CurrentUser)} */}

                                  </td>
                                  <td   style={{minWidth: '110px', maxWidth: '110px', }} >
                                    {/* {(truncateText(item.FileUID.RequestedBy, 25))} */}
                                    {item.RequestedByTitle}
                                  </td>
                                  <td style={{ minWidth: '110px', maxWidth: '110px' }}>
                                    <div
                                      style={{
                                        padding: '5px',
                                        border: '1px solid #efefef',
                                        background: '#fff',
                                        borderRadius: '30px', fontSize: '14px',

                                      }}
                                      className="btn btn-light"
                                    >
                                      {/* {new Date(item?.FileUID?.Created).toLocaleDateString()} */}
                                      {new Date(item?.FileUID?.Created).toLocaleString('en-US', {
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
                                  <td style={{ minWidth: '110px', maxWidth: '110px' }}>
                                    {item?.FileUID?.Status === "Approved" && (
                                      <div
                                        style={{
                                          padding: '5px',
                                          border: '1px solid #efefef',
                                          background: '#fff',
                                          borderRadius: '30px', fontSize: '14px',

                                        }}
                                        className="btn btn-light"
                                      >
                                        {item.assignedtouserTitle}

                                      </div>
                                    )}

                                  </td>
                                  <td style={{ minWidth: '110px', maxWidth: '110px' }}>
                                    {item?.FileUID?.Status === "Approved" && (
                                      <div
                                        style={{
                                          padding: '5px',
                                          border: '1px solid #efefef',
                                          background: '#fff',
                                          borderRadius: '30px', fontSize: '14px',

                                        }}
                                        className="btn btn-light"
                                      >
                                        {/* {item.Modified} */}
                                        {new Date(item?.Modified).toLocaleString('en-US', {
                                          month: '2-digit',
                                          day: '2-digit',
                                          year: 'numeric',
                                          hour: '2-digit',
                                          minute: '2-digit',
                                          second: '2-digit',
                                          hour12: true
                                        })}
                                      </div>
                                    )}

                                  </td>
                                  <td style={{ minWidth: '110px', maxWidth: '110px' }}>
                                    {item.Remark}
                                  </td>
                                  <td style={{ minWidth: '70px', maxWidth: '70px', textAlign: 'center' }}>
                                    {/* <div className="finish mb-0"></div> */}
                                    <div className="finish mb-0">  {item.FileUID.Status} </div>
                                  </td>
                                  {/* <td style={{ minWidth: '70px', maxWidth: '70px' }}>
                                                            <a onClick={(e )=>getTaskItemsbyID(e , item.FileUID.FileUID)}>
                                                                <FontAwesomeIcon icon={faEye} />
                                                            </a>
                                                      </td> */}

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


                  {toggleLog && (
                    <div className="" style={{ backgroundColor: 'white', border: '1px solid #54ade0', marginTop: '20px', borderRadius: '20px', padding: '15px' }}>
                      <iframe id="filePreview" width="100%" height="400"></iframe>
                      <div className="">
                        <div className="">

                          <div className="row">

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
                                  className="btn btn-orange waves-effect waves-light m-1"
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
                              <button
                                type="button"
                                className="btn btn-light waves-effect waves-light m-1"
                              >
                                <i className="fe-x me-1"></i> Cancel
                              </button>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                  )}

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




export default DMSMyrequestaudithistory;
