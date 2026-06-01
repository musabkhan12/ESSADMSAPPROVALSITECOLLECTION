import * as React from 'react'
import { useMediaQuery } from 'react-responsive';
import context from '../../../GlobalContext/context';
import UserContext from '../../../GlobalContext/context';
import { getSP } from '../loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss"
import Provider from '../../../GlobalContext/provider';
import "../components/Managemaster.scss";
import CustomBreadcrumb from '../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb';
import "../../horizontalNavBar/components/horizontalNavbar.scss";
import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
// @ts-ignore
import { getSettingAPI, getSettingAPImanagemaster, getSettingAPIPpowerappsmaster } from "../../../APISearvice/settingsService";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPaperclip } from '@fortawesome/free-solid-svg-icons';
import HorizontalNavbar from '../../horizontalNavBar/components/HorizontalNavBar';
//import styles from "../components/MasterManage.module.scss"
import { IMasterManageProps } from './IManageMasterProps';
import { useEffect, useState } from 'react';
// import { BusinessAppsComponent } from './BusinessAppsComponent';
// import { SettingsAppsComponent } from './MasterSettingsAppComponent';
const endsWith = (str: string, ending: string) => {
  console.log("strrrr", str, ending)
  return str.slice(-ending.length) === ending;
}
let siteID: any;
let response: any;
const specialGroups = ['Super Admin Group', 'Content Contributor Group', 'Intranet Member Group'];
export const MastersettingContext = ({ props }: any) => {
  const sp: SPFI = getSP();
  console.log(sp, 'sp');
  const [showIframe, setShowIframe] = useState(false);
  const [iframeUrl, setIframeUrl] = useState('');
  const handleCardClick = (url: any) => {
    if (url.Title === "Business Apps" || url.Title === "Settings Master") {
      setSelectedItem(url); // Set item to render BusinessAppsComponent
      setShowIframe(true);
    }
    else if(specialGroups.some(group => url.Title.includes(group))){
      setSelectedItem(url); // Set item to render BusinessAppsComponent
      setShowIframe(true);
      setIframeUrl(url);
      
    }
     else {
      window.location.href = url.LinkUrl; // Navigate for other items
    }
    setIframeUrl(url);
    // setShowIframe(true);
    // hideElementsInIframe()

  };
  useEffect(() => {
    if (showIframe && iframeUrl) {
      const iframe = document.querySelector('iframe') as HTMLIFrameElement | null;
      
      if (iframe) {
        const handleLoad = () => {
          const iframeDocument =
  iframe?.contentDocument ||
  iframe?.contentWindow?.document;
  if (!iframeDocument) return;
          const sideNavBox = iframeDocument.getElementById('sideNavBox');
          const s4TitleRow = iframeDocument.getElementById('s4-titlerow');
          const mainribbon = iframeDocument.getElementById('suiteBarDelta');

          if (sideNavBox) {
            // sideNavBox.style.display = 'none';
            sideNavBox?.remove()
          }
          if (mainribbon) {
            // sideNavBox.style.display = 'none';
            mainribbon?.remove()
          }
          if (s4TitleRow) {
            // s4TitleRow.style.display = 'none';
            s4TitleRow?.remove()
          }
        };

        iframe.addEventListener('load', handleLoad);

        // Cleanup the event listener
        return () => {
          iframe.removeEventListener('load', handleLoad);
        };
      }
    }
  }, [showIframe, iframeUrl]);
  const handleBackClick = () => {
    // alert("Back clicked"+ showIframe);
    setShowIframe(false);
  };
 function hideElementsInIframe() {

  const iframe = document.querySelector('iframe') as HTMLIFrameElement | null;

  if (!iframe) return;

  iframe.addEventListener('load', () => {

     const iframeDocument =
  iframe.contentDocument ||
  iframe.contentWindow?.document;

if (!iframeDocument) return;
      if (!iframeDocument) return;
      const sideNavBox = iframeDocument.getElementById('sideNavBox');
      const s4TitleRow = iframeDocument.getElementById('s4-titlerow');
      const s4TitleRow2 = iframeDocument.getElementById('s4-ribbonrow');

      if (sideNavBox) {
        sideNavBox?.remove()
        // sideNavBox.style.display = 'none';
      }
      if (s4TitleRow) {
        // s4TitleRow.style.display = 'none';
        s4TitleRow?.remove()
      }
      if (s4TitleRow2) {
        // s4TitleRow.style.display = 'none';
        s4TitleRow?.remove()
      }
    });
  }
  const isSpecialGroup = (linkUrl: any) => {

    const specialGroups = ['Super Admin Group', 'Content Contributor Group', 'Intranet Member Group'];
    //const specialGroups = ['IntranetAdmin', 'IntranetContentContributor', 'IntranetMembers'];
    return specialGroups.some(group => linkUrl?.includes(group));
  };
  // const { useHide }: any = React.useContext(UserContext);
  // const [isSidebarOpen, setIsSidebarOpen] = React.useState(false);
  const { useHide }: any = React.useContext(UserContext);
  console.log('This function is called only once', useHide);
  const elementRef = React.useRef<HTMLDivElement>(null);
  const [isSidebarOpen, setIsSidebarOpen] = React.useState(true);
  const { setHide }: any = context;
  const [isDarkMode, setIsDarkMode] = React.useState(false);
const [settingArray, setsettingArray] = React.useState<any[]>([]);
  const [IsUserAlllowed, setIsUserAlllowed] = React.useState(false);
  const SiteUrl = props.siteUrl;
  const Breadcrumb = [
    {
      "MainComponent": "Home",
      "MainComponentURl": ""
    },
    {
      "ChildComponent": "Settings",
      "ChildComponentURl": `${SiteUrl}/SitePages/ManageMaster.aspx`
    }
  ]
  const IsUserAllowedAccess = async () => {
    // Get groups for the current user
    const userGroups = await sp.web.currentUser.groups();
    let grptitle: String[] = [];
    for (var i = 0; i < userGroups.length; i++) {
      grptitle.push(userGroups[i].Title.toLowerCase());
    }

let sidebarnavitems = await sp.web.lists
  .getByTitle("ARGSidebarNavigation")
  .items
  .select("Title,Url,Icon,ParentId,ID,EnableAudienceTargeting,Audience/Title")
  .expand("Audience")
  .orderBy("Order0", true)();
    let securednavitems = sidebarnavitems.filter((nav: any) => {
      return (nav.EnableAudienceTargeting && nav.Audience && nav.Audience.some((nv1: any) => { return grptitle?.includes(nv1.Title.toLowerCase()); }))
    }
    )
    console.log("sidebarnavitems", sidebarnavitems, securednavitems)
    let access = securednavitems.some((navitm: any) => (navitm.Url) ? endsWith(navitm.Url.toLowerCase(), location.pathname.toLowerCase()) : false);
    console.log("acess", access);
    //return securednavitems.some((navitm:any)=> (navitm.Url)?endsWith(navitm.Url.toLowerCase(),location.pathname.toLowerCase()):false)
    return true;
  }
  React.useEffect(() => {
    ApiCall()
    console.log('This function is called only once', useHide);
    IsUserAllowedAccess().then(bAllowed => {

      console.log("%c Access allowed", "color:green,font-size:14px", bAllowed);
      setIsUserAlllowed(bAllowed);

    })
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
        toggle.addEventListener('click', () => {
          nav.classList.toggle('show');
          toggle.classList.toggle('bx-x');
          bodypd.classList.toggle('body-pd');
          headerpd.classList.toggle('body-pd');
        });
      }
    };

    showNavbar('header-toggle', 'nav-bar', 'body-pd', 'header');

    const linkColor = document.querySelectorAll('.nav_link');

    function colorLink(this: HTMLElement) {
      if (linkColor) {
        linkColor.forEach(l => l.classList.remove('active'));
        this.classList.add('active');
      }
    }

    linkColor.forEach(l => l.addEventListener('click', colorLink));
  }, [useHide]);
  // Media query to check if the screen width is less than 768px
  const isMobile = useMediaQuery({ query: '(max-width: 768px)' });
  const handleSidebarToggle = (bol: boolean) => {
    setIsSidebarOpen((prevState: any) => !prevState);
    setHide(!bol);
    document.querySelector(".sidebar")?.classList.toggle("close");
  };

  const ApiCall = async () => {
    let path1 = window.location.href;

    let listTitle = 'Settings'
    let CurrentsiteID = props.context.pageContext.site.id;
    siteID = CurrentsiteID;
    response = await sp.web.lists.getByTitle(listTitle).select('Id')();

    // this is code done by anamika
  //   console.log("resp", path1, path1.toLowerCase().includes("managemaster"), path1.toLowerCase().includes("powerappsmaster"));
  //   let settingsData: any;
  //   if (path1.toLowerCase().includes("managemaster")) {
  //     setsettingArray(await getSettingAPImanagemaster(sp))
  //   } else if (path1.toLowerCase().includes("powerappsmaster")) {
  //     setsettingArray(await getSettingAPIPpowerappsmaster(sp))
  //   }
  //   //let Managemasterdata = If(path1.includes("managemaster") ? await getSettingAPImanagemaster(sp)
  //   //let Powerappsmasterdata =
  //   //settingsData = setsettingArray(path1.includes("managemaster") ? await getSettingAPImanagemaster(sp) : await getSettingAPIPpowerappsmaster(sp))
  //   console.log(settingsData, 'settingsData');
  // }

  // this is code done by riya
  let pathsplit = path1.split('/');
  console.log("resp", response);
  // const settingsData = setsettingArray(pathsplit.indexOf("ManageMaster") > -1 ? await getSettingAPImanagemaster(sp) :await getSettingAPIPpowerappsmaster(sp))
  const cleanPath = pathsplit.map(part => part.replace(".aspx", "")); 

const settingsData = setsettingArray(
  cleanPath.some((x:any) => x.toLowerCase() === "managemaster")
      ? await getSettingAPImanagemaster(sp) 
      : await getSettingAPIPpowerappsmaster(sp)
);
  console.log(settingsData, 'settingsData');
}



  // Bannerthing
  const [showModal, setShowModal] = React.useState(false);
  const [showBannerModal, setShowBannerTable] = React.useState(false);
 const [BnnerImagepostArr, setBannerImagepostArr] = React.useState<any[]>([]);
  const [formData, setFormData] = React.useState({
    title: '',
    bannerImg: []
  })
  const setShowModalFunc = (bol: boolean, name: string) => {
    debugger
    if (name == "bannerimg") {
      setShowModal(bol)
      setShowBannerTable(true)
    }
  }
  const onFileChange = async (event: React.ChangeEvent<HTMLInputElement>, libraryName: string, docLib: string) => {
    debugger;
    event.preventDefault();
    let uloadDocsFiles: any[] = [];
    let uloadImageFiles: any[] = [];
    let uloadBannerImageFiles: any[] = [];

    if (event.target.files && event.target.files.length > 0) {
      const files = Array.from(event.target.files);


      const imageVideoFiles = files.filter(file =>
        file.type.startsWith('image/') ||
        file.type.startsWith('video/')
      );

      if (imageVideoFiles.length > 0) {
        const arr = {
          files: imageVideoFiles,
          libraryName: libraryName,
          docLib: docLib
        };

        uloadBannerImageFiles.push(arr);
        setBannerImagepostArr(uloadBannerImageFiles);
      }

    }

  };
  const onChange = (name: string, value: string) => {
    setFormData((prevData) => ({
      ...prevData,
      [name]: value,
    }));
  };

  // Function to check the URL and show an alert
  function checkUrlForMembershipGroupId() {
    const currentUrl = window.location.href;
    const pattern = /_layouts\/15\/people\.aspx\?MembershipGroupId=32/;

    if (pattern.test(currentUrl)) {
      //alert("Match found: MembershipGroupId=32");
    }
  }

  // Call the function
  checkUrlForMembershipGroupId();
  const [selectedItem, setSelectedItem] = useState<any>(null);

  // const handleCardClick = (item) => {
  //   if (item.Title === "Business Apps") {
  //     setSelectedItem(item); // Set item to render BusinessAppsComponent
  //   } else {
  //     window.location.href = item.LinkUrl; // Navigate for other items
  //   }
  // };

  // const handleBackClick = () => {
  //   setSelectedItem(null); // Reset state to go back
  // };

  
  return (

    <div id="wrapper" ref={elementRef}>
      <div
        className='app-menu'
        id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>
      <div className="content-page">
        <HorizontalNavbar _context={sp} siteUrl={SiteUrl} />
        <div className="content" style={{ marginLeft: `${!useHide ? '240px' : '80px'}` }}>
          <div className="container-fluid  paddb">
            <div className="row pt-0 mt-2" style={{ paddingLeft: '0.5rem' }}>
              <div className="col-lg-3">
                 <CustomBreadcrumb Breadcrumb={Breadcrumb} _context={sp}/>
              </div>
              <div className="row manage-master mt-3">
                {/* {console.log("IsUserAlllowed",IsUserAlllowed,settingArray)}
                {
                  IsUserAlllowed ?
                    settingArray.map((item: any) => {
                      const ImageUrl = item.ImageIcon == undefined || item.ImageIcon == null ? "" : JSON.parse(item.ImageIcon);
                      return (<div className="col-sm-3 col-md-3 mt-2">
                        <a href={item?.LinkUrl}>
                          <div className="card-master box1">
                            <div className="icon">
                              <img src={ImageUrl?.serverUrl + ImageUrl?.serverRelativeUrl} />
                            </div>
                            <p className="text-dark">{item.Title}</p>
                          </div>
                        </a>
                      </div>)
                    }) : (<div>Access Denied</div>)
                } */}

                {/* {!showIframe ? (
        IsUserAlllowed ? (
          settingArray.map((item) => {
            //const ImageUrl = item.ImageIcon == undefined || item.ImageIcon == null ? "" : JSON.parse(item.ImageIcon);
            const imageData = item.ImageIcon == undefined || item.ImageIcon == null ? "" : JSON.parse(item.ImageIcon);
            let siteId = siteID;
            let listID =response && response.Id;
            let img1 = imageData && imageData.fileName ? `${SiteUrl}/_api/v2.1/sites('${siteId}')/lists('${listID}')/items('${item.ID}')/attachments('${imageData.fileName}')/thumbnails/0/c400x400/content` : ""
            let img = imageData && imageData.serverRelativeUrl ? `https://officeindia.sharepoint.com${imageData.serverRelativeUrl}` : img1
            const imageUrl = imageData
              //? `${siteUrl}/SiteAssets/Lists/ea596702-57db-4833-8023-5dcd2bba46e3/${imageData.fileName}`
              //? `${imageData.serverUrl}${imageData.serverRelativeUrl}`
              ? img
              : require("../assets/news.png");
            return (
              <div className="col-sm-3 col-md-3 mt-2" key={item.Title}>
                {isSpecialGroup(item?.Title) ? (
                  <div
                    className="card-master box1"
                    onClick={() => handleCardClick(item?.LinkUrl)}
                  >
                    <div className="icon">
                      {/* <img src={ImageUrl?.serverUrl + ImageUrl?.serverRelativeUrl} alt="Icon" /> */}
                {/* <img src={imageUrl} alt="Icon" />
                    </div>
                    <p className="text-dark">{item.Title}</p>
                  </div>
                ) : (
                  <a href={item?.LinkUrl}>
                    <div className="card-master box1">
                      <div className="icon">
                        {/* <img src={ImageUrl?.serverUrl + ImageUrl?.serverRelativeUrl} alt="Icon" /> */}
                {/* <img src={imageUrl} alt="Icon" />
                      </div>
                      <p className="text-dark">{item.Title}</p>
                    </div>
                  </a>
                )}
              </div>
            );
          })
        ) : (
          <div>Access Denied</div>
        )
      ) : (
        <div>
          <button 
            style={{ margin: '10px' }} 
            className="btn btn-secondary" 
            onClick={handleBackClick}
          >
            Back
          </button>
          <iframe
            id='iframe'
            src={iframeUrl}
            style={{ width: '100%', height: '600px', border: 'none' }}
            title="Content"
          ></iframe>
        </div>
      )}  */}

                {!selectedItem
                  ? ( // Show cards when no selection
                    IsUserAlllowed ? (
                      settingArray.map((item) => {
                        //const ImageUrl =
                        // item.ImageIcon == undefined || item.ImageIcon == null
                        //   ? ""
                        //   : JSON.parse(item.ImageIcon);
                        const imageData = item.ImageIcon == undefined || item.ImageIcon == null ? "" : JSON.parse(item.ImageIcon);
                        let siteId = siteID;
                        let listID = response && response.Id;
                        let img1 = imageData && imageData.fileName ? `${SiteUrl}/_api/v2.1/sites('${siteId}')/lists('${listID}')/items('${item.ID}')/attachments('${imageData.fileName}')/thumbnails/0/c400x400/content` : ""
                        let img = imageData && imageData.serverRelativeUrl ? `https://officeindia.sharepoint.com${imageData.serverRelativeUrl}` : img1
                        const imageUrl = imageData
                          //? `${siteUrl}/SiteAssets/Lists/ea596702-57db-4833-8023-5dcd2bba46e3/${imageData.fileName}`
                          //? `${imageData.serverUrl}${imageData.serverRelativeUrl}`
                          ? img
                          : require("../assets/news.png");
                        return (
                          <div className="col-sm-3 col-md-3 mt-2" key={item.Title}>
                            <div
                              className="card-master box1"
                              onClick={() => handleCardClick(item)}
                            >
                            <img className="newopacity" src={imageUrl}/> 
                              <div className="icon">
                                <img
                                  src={imageUrl}
                                  alt="Icon"
                                />
                              </div>
                              <p className="text-dark">{item.Title}</p>
                            </div>
                          </div>
                        );
                      })
                    ) : (
                      <div>Access Denied</div>
                    )
                  ) : (
                    // Render BusinessAppsComponent when Business Apps is clicked
                    <div>
                      <button
                        style={{ margin: "10px" }}
                        className="btn btn-secondary"
                        onClick={handleBackClick}
                      >
                        Back
                      </button>
                      {/* {selectedItem?.Title === "Business Apps" ? (
                        <BusinessAppsComponent data={selectedItem} />
                      ) : 
                      selectedItem?.Title === "Settings Master" ? (
                        <SettingsAppsComponent data={selectedItem} />
                      ) :
                      //  <SettingsAppsComponent data={selectedItem} />} */}
                      <iframe
                      id='iframe'
                      src={selectedItem?.LinkUrl}
                      style={{ width: '100%', height: '600px', border: 'none' }}
                      title="Content"
                    ></iframe>
                       
                    </div>
                  )

                }


                <div id="iframeContainer" style={{ marginTop: '20px' }}></div>


                {/* {
  IsUserAlllowed ? (
    settingArray.map((item: any) => {
      const ImageUrl = item.ImageIcon == undefined || item.ImageIcon == null ? "" : JSON.parse(item.ImageIcon);
      return (
        <div className="col-sm-3 col-md-3 mt-2">
          <div
            className="card-master box1"
            onClick={() => {
              const iframe = document.createElement("iframe");
              iframe.src = item?.LinkUrl;
              iframe.style.width = "100%";
              iframe.style.height = "600px"; // Adjust height as needed
              iframe.style.border = "none";

              const container = document.getElementById("iframeContainer");
              container.innerHTML = ""; // Clear previous content
              container.appendChild(iframe);
            }}
          >
            <div className="icon">
              <img src={ImageUrl?.serverUrl + ImageUrl?.serverRelativeUrl} alt="Icon" />
            </div>
            <p className="text-dark">{item.Title}</p>
          </div>
        </div>
      );
    })
  ) : (
    <div>Access Denied</div>
  )
}


<div id="iframeContainer" style={{ marginTop: '20px' }}></div> */}

              </div>
              {/* <>

                <TableData/>
                </> */}
            </div>

          </div>
        </div>
      </div>
    </div>


  )
}

const MasterManage: React.FC<IMasterManageProps> = (props) => {
  return (
    <Provider>
      <MastersettingContext props={props} />
    </Provider>
  );
};

export default MasterManage;