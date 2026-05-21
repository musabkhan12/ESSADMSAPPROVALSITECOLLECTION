
import { escape } from "@microsoft/sp-lodash-subset";

import React, { useEffect, useRef, useState } from "react";

import VerticalSideBar from "../../verticalSideBar/components/VerticalSideBar";

import HorizontalNavbar from "../../horizontalNavBar/components/HorizontalNavBar";

import "bootstrap/dist/css/bootstrap.min.css";

import "../../../CustomCss/mainCustom.scss";

import "../components/MyApproval.scss";

import "bootstrap/dist/js/bootstrap.bundle.min.js";

import "./CustomTable.scss";

import "../../verticalSideBar/components/VerticalSidebar.scss";

import { IMyApprovalProps } from "./IMyApprovalProps";

import Provider from "../../../GlobalContext/provider";

import UserContext from "../../../GlobalContext/context";

// import CustomBreadcrumb from "../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb";

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
let currentItemID = ''
import moment from "moment";


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

import { getDataByID, getMyApproval, getMyRequest, updateItemApproval } from "../../../APISearvice/ApprovalService";
import DMSMyApprovalAction from "./ApprovalAction";

const MyApprovalContext = ({ props }: any) => {

  const sp: SPFI = getSP();
  const [activeComponent, setActiveComponent] = useState<string>('');
  const { useHide }: any = React.useContext(UserContext);

  const [announcementData, setAnnouncementData] = React.useState([]);

  const [myApprovalsData, setMyApprovalsData] = React.useState([]);

  const elementRef = React.useRef<HTMLDivElement>(null);

  const SiteUrl = props.siteUrl;

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

  const [approved, setApproved] = useState('yes');

  const [filters, setFilters] = React.useState({

    SNo: "",

    RequestID: "",

    ProcessName: "",

    RequestedBy: "",

    RequestedDate: "",

    Status: "",

  });


  const [isOpen, setIsOpen] = React.useState(false);

  const [IsinvideHide, setIsinvideHide] = React.useState(false);

  // const [Mylistdata, setMylistdata] = useState([]);
  const [Mylistdata, setMylistdata] = useState<any[]>([]);
  const handleReturnToMain = (Name: any) => {
    setActiveComponent(Name); // Reset to show the main component
    console.log(activeComponent, "activeComponent updated")
  };
  const getApprovalmasterTasklist = async () => {
    try {
      const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select(
        "Log", "CurrentUser", "Remark"
        , "LogHistory"
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
        .filter(`CurrentUser eq '${currentUserEmailRef.current}'`)();;
      console.log(items, "DMSFileApprovalTaskList");
      setMylistdata(items);

    } catch (error) {
      console.error("Error fetching list items:", error);
    }
  };
  console.log(Mylistdata, "Mylistdata")
  const currentUserEmailRef = useRef('');
  const getCurrrentuser = async () => {
    const userdata = await sp.web.currentUser();
    currentUserEmailRef.current = userdata.Email;
    getApprovalmasterTasklist();
  }
  useEffect(() => {
    getCurrrentuser()

  }, []);

  const truncateText = (text: string, maxLength?: any) => {
    if (text) {
      return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;
    }

  };


  const getTaskItemsbyID = async (e: any, itemid: any) => {
    // currentItemID = itemid
    currentItemID = itemid
    setActiveComponent('Approval Action')
    console.log("itemid", itemid)
    const items = await sp.web.lists.getByTitle('DMSFileApprovalTaskList').items.select("CurrentUser", "FileUID/FileUID", "Log").expand("FileUID").filter(`FileUID/RequestNo eq '${itemid}'`)();
    console.log(items, "items")
  }
  const toggleDropdown = () => {

    setIsOpen(!isOpen);

  };


  const [sortConfig, setSortConfig] = React.useState({

    key: "",

    direction: "ascending",

  });


  const [formData, setFormData] = React.useState({

    Remark: '',

  })
  const handleCancel = () => {
    window.location.href = `${siteUrl}/SitePages/MyApprovals.aspx`;
  }
  //#region OnchangeData

  const onChange = (name: string, value: string) => {

    debugger

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

  const [isActivedata, setisActivedata] = useState(false)

  const [DiscussionData, setDiscussion] = useState([]);

  const [CategoryData, setCategoryData] = React.useState([]);

  const [showModal, setShowModal] = React.useState(false);

  const [showDocTable, setShowDocTable] = React.useState(false);

  const [showImgModal, setShowImgTable] = React.useState(false);

  const [showBannerModal, setShowBannerTable] = React.useState(false);

  const [currentUser, setCurrentUser] = React.useState(null);

  const [editForm, setEditForm] = React.useState(false);

  const [richTextValues, setRichTextValues] = React.useState<{

    [key: string]: string;

  }>({});

  const [activeTab, setActiveTab] = useState("Intranet");

  //   const handleTabClick = async (tab: React.SetStateAction<string>) => {

  //     setActiveTab(tab);

  //     if (tab == "profile11") {

  //       FilterDiscussionData("Today");

  //     } else if (tab == "profile1") {

  //       setAnnouncementData(await getDiscussionMeAll(sp));

  //     } else {

  //       const announcementArr = await getDiscussionForum(sp);

  //       let lengArr: any;

  //       for (var i = 0; i < announcementArr.length; i++) {

  //         lengArr = await getDiscussionComments(sp, announcementArr[i].ID);

  //         console.log(lengArr, "rrr");

  //         (announcementArr[i].commentsLength = lengArr.arrLength),

  //           (announcementArr[i].Users = lengArr.arrUser);

  //       }

  //       setAnnouncementData(announcementArr);

  //     }

  //   };


  React.useEffect(() => {

    sessionStorage.removeItem("announcementId");

    ApiCall();

  }, [useHide]);


  const ApiCall = async () => {


    setMyApprovalsData(await getMyApproval(sp))



  };

  //   const FilterDiscussionData = async (optionFilter: string) => {

  //     setAnnouncementData(await getDiscussionFilterAll(sp, optionFilter));

  //   };

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


  const applyFiltersAndSorting = (data: any[]) => {

    // Filter data

    console.log(

      "filters",

      data,

      filters,

      filters.ProcessName,

      filters.RequestID

    );

    const filteredData = data.filter((item, index) => {


      return (

        (filters.SNo === "" || String(index + 1).includes(filters.SNo)) &&

        (filters.ProcessName === "" ||

          (item.ProcessName != undefined &&

            item.ProcessName.toLowerCase().includes(

              filters.ProcessName.toLowerCase()

            ))) &&

        (filters.RequestID === "" ||

          (item.RequestID != undefined &&

            item.RequestID.toLowerCase().includes(

              filters.RequestID.toLowerCase()

            ))) &&

        // (filters.ProcessName === "" ||

        //   item.ProcessName.toLowerCase().includes(filters.ProcessName.toLowerCase())) &&

        (filters.Status === "" ||

          item.Status.toLowerCase().includes(filters.Status.toLowerCase())) &&

        (filters.RequestedBy === "" ||

          item?.Requester?.Title?.toLowerCase().includes(

            filters.RequestedBy.toLowerCase()

          ))

      );

    });


    const sortedData = filteredData.sort((a, b) => {

      if (sortConfig.key === "SNo") {

        // Sort by index

        const aIndex = data.indexOf(a);

        const bIndex = data.indexOf(b);

        return sortConfig.direction === "ascending"

          ? aIndex - bIndex

          : bIndex - aIndex;

      } else if (sortConfig.key) {

        // Sort by other keys

        const aValue = a[sortConfig.key] ? a[sortConfig.key].toLowerCase() : "";

        const bValue = b[sortConfig.key] ? b[sortConfig.key].toLowerCase() : "";

        if (aValue < bValue) {

          return sortConfig.direction === "ascending" ? -1 : 1;

        }

        if (aValue > bValue) {

          return sortConfig.direction === "ascending" ? 1 : -1;

        }

      }

      return 0;

    });

    return sortedData;

  };


  const filteredMyApprovalData = applyFiltersAndSorting(myApprovalsData);

  const filteredNewsData = applyFiltersAndSorting(newsData);

  const [currentPage, setCurrentPage] = React.useState(1);

  const itemsPerPage = 10;

  const totalPages = Math.ceil(filteredMyApprovalData.length / itemsPerPage);


  const [ContentData, setContentData] = React.useState<any>([]);

  const [currentItem, setCurrentItem] = React.useState<any>([]);



  const handlePageChange = (pageNumber: any) => {

    if (pageNumber > 0 && pageNumber <= totalPages) {

      setCurrentPage(pageNumber);

    }

  };

  const startIndex = (currentPage - 1) * itemsPerPage;

  const endIndex = startIndex + itemsPerPage;

  const currentData = filteredMyApprovalData.slice(startIndex, endIndex);

  const newsCurrentData = filteredNewsData.slice(startIndex, endIndex);

  const [editID, setEditID] = React.useState(null);

  const [ImagepostIdsArr, setImagepostIdsArr] = React.useState([]);

  const siteUrl = props.siteUrl;


  const Breadcrumb = [

    {

      MainComponent: "Home",

      MainComponentURl: `${siteUrl}/SitePages/Dashboard.aspx`,

    },

    {

      ChildComponent: "My Approval",

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

  //   const fetchOptions = async () => {

  //     try {

  //       const items = await fetchUserInformationList(sp);

  //       console.log(items, "itemsitemsitems");


  //       const formattedOptions = items.map((item: { Title: any; Id: any }) => ({

  //         name: item.Title, // Adjust according to your list schema

  //         id: item.Id,

  //       }));

  //       setOpions(formattedOptions);

  //     } catch (error) {

  //       console.error("Error fetching options:", error);

  //     }

  //   };


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


  const handleRedirect = async (e: React.MouseEvent<SVGElement, MouseEvent>, Item: any) => {

    e.preventDefault()

    let arr = []


    setCurrentItem(Item)


    setContentData(await getDataByID(sp, Item?.ContentId, Item?.ContentName))

    setisActivedata(true)

    let sessionkey = "";
    let redirecturl = "";
    if (Item?.ProcessName) {
      switch (Item?.ProcessName) {
        case "Announcement":
          sessionkey = "announcementId";
          redirecturl = `${siteUrl}/SitePages/AddAnnouncement.aspx` + "?requestid=" + Item?.Id + "&mode=approval";
          break;
        case "News":
          sessionkey = "announcementId";
          redirecturl = `${siteUrl}/SitePages/AddAnnouncement.aspx` + "?requestid=" + Item?.Id + "&mode=approval";
          break;
        case "Event":
          sessionkey = "EventId";
          redirecturl = `${siteUrl}/SitePages/EventMasterForm.aspx` + "?requestid=" + Item?.Id + "&mode=approval";
          break;
        case "Media":
          sessionkey = "mediaId";
          redirecturl = `${siteUrl}/SitePages/MediaGalleryForm.aspx` + "?requestid=" + Item?.Id + "&mode=approval";
          break;
        default: ;
      }

      const encryptedId = encryptId(String(Item?.ContentId));
      sessionStorage.setItem(sessionkey, encryptedId);
      location.href = redirecturl;

    }
    // const encryptedId = encryptId(String(Item?.ContentId));

    // sessionStorage.setItem("announcementId", encryptedId);


  };


  const handleFromSubmit = async (e: React.MouseEvent<HTMLButtonElement, MouseEvent>, Status: string) => {

    e.preventDefault();


    const postPayload = {

      Remark: formData.Remark,

      Status: Status,
      TriggerUpdateFlow: true

    };

    console.log(postPayload);

    const postResult = await updateItemApproval(postPayload, sp, currentItem.Id);

    if (postResult) {

      setTimeout(() => {

        window.location.reload()

      }, 1000);


    }

  }


  return (

    <div id="wrapper" ref={elementRef}>

      <div className="app-menu" id="myHeader">

        <VerticalSideBar _context={sp} />

      </div>


      <div className="content-page">

        <HorizontalNavbar />

        <div

          className="content"

          style={{

            marginLeft: `${!useHide ? "240px" : "80px"}`,

            marginTop: "0rem",

          }}

        >

          <div className="container-fluid paddb">

            <div className="row" style={{ paddingLeft: "0.5rem" }}>

              {/* <div className="col-lg-6">

                 <CustomBreadcrumb Breadcrumb={Breadcrumb} _context={sp}/>

              </div> */}

            </div>
            <div className="row mt-4">

              <div className="col-12">

                <div className="card mb-0 cardcsss">

                  <div className="card-body">

                    <div className="d-flex flex-wrap align-items-center justify-content-center">

                      <ul

                        className="navs nav-pillss navtab-bgs"

                        role="tablist"

                        style={{

                          gap: "5px",

                          display: "flex",

                          listStyle: "none",

                          marginBottom: "unset",

                        }}

                      >

                        <li className="nav-itemcss">

                          <a

                            className={`nav-linkss ${activeTab === "cardView" ? "active" : ""

                              }`}

                            aria-selected={activeTab === "cardView"}

                            role="tab"

                            onClick={() => setActiveTab("Intranet")}

                          >
                            Intranet


                          </a>

                        </li>

                        <li className="nav-itemcss">

                          <a

                            className={`nav-linkss ${activeTab === "listView" ? "active" : ""

                              }`}

                            aria-selected={activeTab === "listView"}

                            role="tab"

                            onClick={() => setActiveTab("DMS")}


                          // onClick={() => handleTabChange("listView")}

                          // className={`nav-link ${

                          //   activeTab === "listView" ? "active" : ""

                          // }`}

                          >
                            DMS

                          </a>

                        </li>
                        <li className="nav-itemcss">

                          <a

                            className={`nav-linkss ${activeTab === "listView" ? "active" : ""

                              }`}

                            aria-selected={activeTab === "listView"}

                            role="tab"




                            onClick={() => setActiveTab("Automation")}

                          // className={`nav-link ${

                          //   activeTab === "listView" ? "active" : ""

                          // }`}

                          >

                            Automation
                          </a>

                        </li>

                      </ul>

                    </div>

                  </div>

                </div>

              </div>

            </div>
            {activeTab === "Intranet" && (
              <div>
                {

                  !isActivedata && (

                    <div className="card cardCss mt-4">

                      <div className="card-body">

                        <div id="cardCollpase4" className="collapse show">

                          <div className="table-responsive pt-0">

                            <table

                              className="mtable mt-0 table-centered table-nowrap table-borderless mb-0"

                              style={{ position: "relative" }}

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

                                      style={{ justifyContent: "space-between" }}

                                    >

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

                                        className="inputcss"

                                        style={{ width: "100%" }}

                                      />

                                    </div>

                                  </th>

                                  <th style={{ minWidth: "80px", maxWidth: "80px" }}>

                                    <div className="d-flex flex-column bd-highlight ">

                                      <div

                                        className="d-flex pb-2"

                                        style={{ justifyContent: "space-between" }}

                                      >

                                        <span>Request ID</span>

                                        <span

                                          onClick={() => handleSortChange("RequestID")}

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

                                  <th style={{ minWidth: "120px", maxWidth: "120px" }}>

                                    <div className="d-flex flex-column bd-highlight ">

                                      <div

                                        className="d-flex  pb-2"

                                        style={{ justifyContent: "space-between" }}

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

                                  <th style={{ minWidth: "100px", maxWidth: "100px" }}>

                                    <div className="d-flex flex-column bd-highlight ">

                                      <div

                                        className="d-flex  pb-2"

                                        style={{ justifyContent: "space-between" }}

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

                                  <th style={{ minWidth: "100px", maxWidth: "100px" }}>

                                    <div className="d-flex flex-column bd-highlight ">

                                      <div

                                        className="d-flex  pb-2"

                                        style={{ justifyContent: "space-between" }}

                                      >

                                        <span>Requested Date</span>{" "}

                                        <span

                                          onClick={() =>

                                            handleSortChange("RequestedDate")

                                          }

                                        >

                                          <FontAwesomeIcon icon={faSort} />{" "}

                                        </span>

                                      </div>

                                      <div className=" bd-highlight">

                                        <input

                                          type="text"

                                          placeholder="Filter by Requested Date"

                                          onChange={(e) =>

                                            handleFilterChange(e, "RequestedDate")

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

                                  <th style={{ minWidth: "80px", maxWidth: "80px" }}>

                                    <div className="d-flex flex-column bd-highlight ">

                                      <div

                                        className="d-flex  pb-2"

                                        style={{ justifyContent: "space-between" }}

                                      >

                                        <span>Status</span>{" "}

                                        <span

                                          onClick={() => handleSortChange("Status")}

                                        >

                                          <FontAwesomeIcon icon={faSort} />{" "}

                                        </span>

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

                                        style={{ justifyContent: "space-between" }}

                                      >

                                        <span>Action</span>{" "}

                                        {/* <span

          onClick={() => handleSortChange("Category")}

        >

          <FontAwesomeIcon icon={faSort} />{" "}

        </span> */}

                                      </div>

                                      {/* <div className=" bd-highlight">

        <input

          type="text"

          placeholder="Filter by Requested By"

          onChange={(e) =>

            handleFilterChange(e, "Category")

          }

          className="inputcss"

          style={{ width: "100%" }}

        />

      </div> */}

                                    </div>

                                  </th>

                                </tr>

                              </thead>

                              <tbody>

                                {currentData.length === 0 ? (

                                  <div

                                    className="no-results"

                                    style={{

                                      display: "flex",

                                      justifyContent: "center",

                                    }}

                                  >

                                    No results found

                                  </div>

                                ) : (

                                  currentData.map((item: any, index: number) => (

                                    <tr


                                      key={index}

                                      style={{ cursor: "pointer" }}

                                    >
                                      <td style={{ minWidth: '40px', maxWidth: '40px' }}><div className='indexdesign'>   {startIndex + index + 1}</div>  </td>



                                      <td

                                        style={{

                                          minWidth: "80px",

                                          maxWidth: "80px",

                                          textTransform: "capitalize",

                                        }}

                                      >

                                        {item.RequestID}

                                      </td>

                                      <td

                                        style={{ minWidth: "120px", maxWidth: "120px" }}

                                      >

                                        {item.ProcessName}

                                      </td>

                                      <td

                                        style={{ minWidth: "100px", maxWidth: "100px" }}

                                      >

                                        {item?.Requester?.Title}

                                      </td>

                                      <td

                                        style={{ minWidth: "100px", maxWidth: "100px" }}

                                      >
                                        <div className="btn btn-light">
                                          {new Date(item?.Created).toLocaleDateString()}
                                        </div>

                                      </td>

                                      <td

                                        style={{ minWidth: "80px", maxWidth: "80px" }}

                                      >
                                        <div className="btn btn-status">
                                          {item?.Status}
                                        </div>
                                      </td>

                                      <td

                                        style={{ minWidth: "50px", maxWidth: "50px" }}

                                        className="fe-eye font-18"

                                      >

                                        {/* <a href="my-approval-form.html"><i className="fe-eye font-18"></i> </a> */}


                                        {/* <img

          onClick={() =>

            handleRedirect(item.RedirectionLink)

          }

          style={{

            minWidth: "20px",

            maxWidth: "20px",

            marginLeft: "15px",

            cursor: "pointer",

          }}

          src={require("../assets/eye.png")}

          className="fe-eye font-18"

          alt={item.Title || "Untitled"}

        /> */}

                                        <Edit onClick={(e) =>

                                          handleRedirect(e, item)

                                        }

                                          style={{

                                            minWidth: "20px",

                                            maxWidth: "20px",

                                            marginLeft: "15px",

                                            cursor: "pointer",

                                          }} />

                                      </td>

                                    </tr>

                                  ))

                                )}

                              </tbody>

                            </table>

                          </div>


                          {currentData.length > 0 ? (

                            <nav className="pagination-container">

                              <ul className="pagination">

                                <li

                                  className={`page-item ${currentPage === 1 ? "disabled" : ""

                                    }`}

                                >

                                  <a

                                    className="page-link"

                                    onClick={() => handlePageChange(currentPage - 1)}

                                    aria-label="Previous"

                                  >

                                    «

                                  </a>

                                </li>

                                {Array.from({ length: totalPages }, (_, num) => (

                                  <li

                                    key={num}

                                    className={`page-item ${currentPage === num + 1 ? "active" : ""

                                      }`}

                                  >

                                    <a

                                      className="page-link"

                                      onClick={() => handlePageChange(num + 1)}

                                    >

                                      {num + 1}

                                    </a>

                                  </li>

                                ))}


                                <li

                                  className={`page-item ${currentPage === totalPages ? "disabled" : ""

                                    }`}

                                >

                                  <a

                                    className="page-link"

                                    onClick={() => handlePageChange(currentPage + 1)}

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

                      </div>

                    </div>

                  )

                }

                {

                  isActivedata == true && ContentData.length > 0 && currentItem != null &&

                  (

                    <div className="row mt-4">

                      <div className="col-12">

                        <div className="card">

                          <div className="card-body">

                            <h4 className="header-title mb-0">{ContentData[0].Title}</h4>

                            <p className="sub-header">{currentItem.EntityName}</p>


                            <div className="row">

                              <div className="col-lg-4">

                                <div className="mb-3">

                                  <label className="form-label text-dark font-14">Company / Department:</label>

                                  <div><span className="text-muted font-14">{currentItem.EntityName}</span></div>

                                </div>

                              </div>

                              <div className="col-lg-4">

                                <div className="mb-3">

                                  <label className="form-label text-dark font-14">Date of Request:</label>

                                  <div><span className="text-muted font-14">{currentItem.Created}</span></div>

                                </div>

                              </div>

                              <div className="col-lg-4">

                                <div className="mb-3">

                                  <label className="form-label text-dark font-14">Status:</label>

                                  <div><span className="text-muted font-14">{currentItem.Status}</span></div>

                                </div>

                              </div>

                              <div className="col-lg-6">

                                <div className="mb-0">

                                  <label className="form-label text-dark font-14">Content:</label>

                                  <div>

                                    <span className="text-muted font-14">

                                      {ContentData[0].Title || ContentData[0].EventName}

                                    </span>

                                  </div>

                                </div>

                              </div>

                              <div className="col-lg-6">

                                <div className="mb-0">

                                  <label className="form-label text-dark font-14">Overview:</label>

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

                        {

                          ContentData[0]?.Description != null && (

                            <div className="card">

                              <div className="card-body">

                                <h4 className="header-title mb-0">Description</h4>

                                <p className="sub-header">

                                  <div

                                    dangerouslySetInnerHTML={{ __html: ContentData[0].Description }}

                                  ></div>


                                </p>

                              </div>

                            </div>

                          )

                        }


                        {/* card three */}


                        {/* <div className="card">

        <div className="card-body">

          <h4 className="header-title mb-0">Finance Manager's Comment</h4>

          <p className="sub-header">

            Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.

          </p>


          <div className="row">

            <div className="col-lg-4">

              <div className="mb-3">

                <label className="form-label text-dark font-14">Name:</label>

                <div><span className="text-muted font-14">XYX</span></div>

              </div>

            </div>

            <div className="col-lg-4">

              <div className="mb-3">

                <label className="form-label text-dark font-14">Date:</label>

                <div><span className="text-muted font-14">01-10-2023</span></div>

              </div>

            </div>

            <div className="col-lg-4">

              <div className="mb-3">

                <label className="form-label text-dark font-14">Signature:</label>

                <div><span className="text-muted font-14">XYX</span></div>

              </div>

            </div>

            <div className="col-lg-12">

              <div className="mb-0">

                <label className="form-label text-dark font-14">FM's Remark and Recommendation:</label>

                <div>

                  <span className="text-muted font-14">

                    Lorem Ipsum is simply dummy text of the printing and typesetting industry.

                  </span>

                </div>

              </div>

            </div>

          </div>

        </div>

      </div> */}


                        {/* card four */}


                        {/* <div className="card">

        <div className="card-body">

          <h4 className="header-title mb-0">GM's Approval</h4>

          <p className="sub-header">

            Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.

          </p>


          <div className="row">

            <div className="col-lg-4">

              <div className="mb-3">

                <label className="form-label text-dark font-14">Name:</label>

                <div><span className="text-muted font-14">XYX</span></div>

              </div>

            </div>

            <div className="col-lg-4">

              <div className="mb-3">

                <label className="form-label text-dark font-14">Date:</label>

                <div><span className="text-muted font-14">01-10-2023</span></div>

              </div>

            </div>

            <div className="col-lg-4">

              <div className="mb-3">

                <label className="form-label text-dark font-14">Signature:</label>

                <div><span className="text-muted font-14">XYX</span></div>

              </div>

            </div>

            <div className="col-lg-12">

              <div className="mb-0">

                <label className="form-label text-dark font-14">GM's Remark and Recommendation:</label>

                <div>

                  <span className="text-muted font-14">

                    Lorem Ipsum is simply dummy text of the printing and typesetting industry.

                  </span>

                </div>

              </div>

            </div>

          </div>

        </div>

      </div> */}


                        {/* card five */}


                        {

                          currentItem.Status == "Submitted" && (

                            <div className="card">

                              <div className="card-body">

                                {/* <h4 className="header-title mb-0">GCFO's Approval</h4> */}

                                {/* <p className="sub-header">

          Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s.

        </p> */}


                                <div className="row">

                                  {/* <div className="col-lg-4">

            <div className="mb-3">

              <label className="form-label text-dark font-14">Approved:</label> <br />

              <div className="form-check form-check-inline">

                <input

                  type="radio"

                  id="customRadio1"

                  name="customRadio"

                  className="form-check-input"

                  checked={approved === 'yes'}

                  onChange={() => setApproved('yes')}

                />

                <label className="form-check-label" htmlFor="customRadio1">Yes</label>

              </div>

              <div className="form-check form-check-inline">

                <input

                  type="radio"

                  id="customRadio2"

                  name="customRadio"

                  className="form-check-input"

                  checked={approved === 'no'}

                  onChange={() => setApproved('no')}

                />

                <label className="form-check-label" htmlFor="customRadio2">No</label>

              </div>

            </div>

          </div> */}


                                  {/* <div className="col-lg-4">

            <div className="mb-3">

              <label className="form-label text-dark font-14">Date:</label>

              <div><span className="text-muted font-14">01-10-2023</span></div>

            </div>

          </div> */}


                                  {/* <div className="col-lg-4">

            <div className="mb-3">

              <label className="form-label text-dark font-14">Signature:</label>

              <div><span className="text-muted font-14">XYX</span></div>

            </div>

          </div> */}


                                  {/* <div className="col-lg-12">

            <div className="mb-3">

              <label className="form-label text-dark font-14">GCFO's Remarks:</label>

              <div>

                <span className="text-muted font-14">

                  Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum is simply dummy text of the printing and typesetting industry.

                </span>

              </div>

            </div>

          </div> */}


                                  {

                                    currentItem.Status == "Submitted" && (<div className="col-lg-12">

                                      <div className="mb-0" >

                                        <label htmlFor="example-textarea" className="form-label text-dark font-14">Remarks:</label>

                                        <textarea className="form-control" id="example-textarea" rows={5} name="Remark" value={formData.Remark}

                                          onChange={(e) => onChange(e.target.name, e.target.value)}></textarea>

                                      </div>

                                    </div>)

                                  }


                                </div>

                                {


                                  currentItem.Status == "Submitted" && (

                                    <div className="row mt-3">

                                      <div className="col-12 text-center">

                                        <a href="my-approval.html">

                                          <button type="button" className="btn btn-success waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Approved')}>

                                            <i className="fe-check-circle me-1"></i> Approve

                                          </button>

                                        </a>

                                        <a href="my-approval.html">

                                          <button type="button" className="btn btn-warning waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Rework')}>

                                            <i className="fe-corner-up-left me-1"></i> Rework

                                          </button>

                                        </a>

                                        <a href="my-approval.html">

                                          <button type="button" className="btn btn-danger waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Reject')}>

                                            <i className="fe-x-circle me-1"></i> Reject

                                          </button>

                                        </a>

                                        <button type="button" className="btn btn-light waves-effect waves-light m-1" onClick={(e) => handleCancel()}>

                                          <i className="fe-x me-1"></i> Cancel

                                        </button>

                                      </div>

                                    </div>

                                  )

                                }


                              </div>

                            </div>

                          )

                        }




                      </div>

                    </div>

                  )

                }
              </div>
            )

            }
            {activeTab === "DMS" && (
              <div>
                {activeComponent === "" ?
                  (<div>
                    <div className="DMSMasterContainer">
                      <h4 className="page-title fw-bold mb-1 font-20">My Approvals 1</h4>
                      <div className="" style={{ backgroundColor: 'white', border: '1px solid #54ade0', marginTop: '20px', borderRadius: '20px', padding: '15px' }}>
                        <table className="mtbalenew">
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

                            {Mylistdata.length > 0 ? Mylistdata.map((item, index) => {
                              return (
                                <tr>
                                  <td style={{ minWidth: '40px', maxWidth: '40px' }}>
                                    <span style={{ marginLeft: '5px' }} className="indexdesign">{index}</span>
                                  </td>
                                  <td >{(truncateText(item.FileUID.FileUID, 22))}</td>
                                  <td >Capex Form</td>
                                  <td >{(truncateText(item.FileUID.RequestedBy, 22))}</td>
                                  <td >
                                    <div
                                      style={{
                                        padding: '5px',
                                        border: '1px solid #efefef',
                                        background: '#fff', fontSize: '14px',
                                        borderRadius: '30px',

                                      }}
                                      className="btn btn-light"
                                    >
                                      {item.FileUID.Created}
                                    </div>
                                  </td>
                                  <td style={{ minWidth: '80px', maxWidth: '80px', textAlign: 'center' }}>
                                    <div className="finish mb-0">Pending</div>
                                  </td>
                                  <td style={{ minWidth: '70px', maxWidth: '70px' }}>
                                    <a onClick={(e) => getTaskItemsbyID(e, item.FileUID.RequestNo)}>
                                      <FontAwesomeIcon icon={faEye} />
                                    </a>
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
                  </div>) : (
                    <div>
                      {activeComponent === 'Approval Action' && (
                        <div>
                          <button style={{ float: 'right' }} type="button" className="btn btn-secondary" onClick={() => handleReturnToMain('')}> Back </button>
                          <DMSMyApprovalAction props={currentItemID} />
                        </div>

                      )}

                    </div>


                  )
                }

              </div>

            )

            }
            {activeTab === "Automation" && (
              <div>Automation</div>

            )

            }


          </div>

        </div>

      </div>

    </div>

  );

};


const MyApproval: React.FC<IMyApprovalProps> = (props) => (

  <Provider>

    <MyApprovalContext props={props} />

  </Provider>

);

export default MyApproval;



