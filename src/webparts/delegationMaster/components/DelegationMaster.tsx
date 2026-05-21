import * as React from 'react';
// import styles from './Delegation.module.scss';

import Provider from '../../../GlobalContext/provider';
import { faArrowLeft, faEdit, faPlusCircle, faSort, faEllipsisV, faFileExport } from '@fortawesome/free-solid-svg-icons';
import { SPFI } from '@pnp/sp/presets/all';
import UserContext from '../../../GlobalContext/context';
import { allowstringonly, getCurrentUser } from '../../../APISearvice/CustomService';
import { decryptId, encryptId } from '../../../APISearvice/CryptoService';
import { DeleteDelegateAPI, getDelegateList } from '../../../APISearvice/DelegateService';
import Swal from 'sweetalert2';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faPaperclip } from '@fortawesome/free-solid-svg-icons';
import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import "bootstrap/dist/css/bootstrap.min.css";
import "../../../CustomCss/mainCustom.scss";
import "../../verticalSideBar/components/VerticalSidebar.scss";
import VerticalSideBar from '../../verticalSideBar/components/VerticalSideBar';
import CustomBreadcrumb from '../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb';

import "../../../CustomJSComponents/CustomForm/CustomForm.scss"
import HorizontalNavbar from '../../horizontalNavBar/components/HorizontalNavBar';
import context from '../../../GlobalContext/context';
import { useRef } from 'react';
import { getSP } from '../loc/pnpjsConfig';
import moment from 'moment';
import * as XLSX from 'xlsx';
import type { IDelegationMasterProps } from './IDelegationMasterProps';
import "../../../CustomJSComponents/CustomTable/CustomTable.scss";



const DelegationMasterContext = ({ props }: any) => {
  const elementRef = React.useRef<HTMLDivElement>(null);
  const sp: SPFI = getSP(props.context);
  const Spurl = sp.web;
  const siteUrl = props.siteUrl;
  const { useHide }: any = React.useContext(UserContext);
  // console.log('This function is called only once', useHide);
  const [sortConfig, setSortConfig] = React.useState({ key: '', direction: 'ascending' });
  const [Delegatelistdata, setDelegateData] = React.useState([]);

 const ApiCall = async () => {
    try {
        let DelegateArr: any[] = [];
        const userGroups = await sp.web.currentUser.groups();
        let groupTitles: string[] = userGroups.map((group) => group.Title.toLowerCase());

        // Check if user is super admin
        const isSuperAdmin = groupTitles.includes("intranetadmin") ? "Yes" : "No";
        
        DelegateArr = await getDelegateList(sp, isSuperAdmin);
        console.log("DelegateArr", DelegateArr);
        setDelegateData(DelegateArr);
    } catch (error) {
        console.error("Error in ApiCall:", error);
        Swal.fire({
            icon: 'error',
            title: 'Error',
            text: 'Failed to load delegation data',
        });
    }
};
  const [filters, setFilters] = React.useState({
    SNo: '',
    Title: '',
    Overview: '',
    URL: '',
    Status: '',
    SubmittedDate: ''
  });
  const Breadcrumb = [
    {
      "MainComponent": "Settings",
      "MainComponentURl": `${siteUrl}/SitePages/Settings.aspx`
    },
    {
      "ChildComponent": "Delegation Master",
      "ChildComponentURl": `${siteUrl}/SitePages/DelegationMaster.aspx`
    }
  ]

  React.useEffect(() => {
    ApiCall();
  }, [siteUrl]);

  const handleFilterChange = (e: React.ChangeEvent<HTMLInputElement>, field: string) => {
    setFilters({
      ...filters,
      [field]: e.target.value,
    });
  };


  const handleSortChange = (key: string) => {
    let direction = 'ascending';
    if (sortConfig.key === key && sortConfig.direction === 'ascending') {
      direction = 'descending';
    }
    setSortConfig({ key, direction });
  };
  const applyFiltersAndSorting = (data: any[]) => {
    // debugger
    // Filter data
    const filteredData = data.filter((item, index) => {
      return (
        (filters.SNo === '' || String(index + 1).includes(filters.SNo)) &&
        (filters.Title === '' || item.Title.toLowerCase().includes(filters.Title.toLowerCase())) &&
        (filters.URL === '' || item.URL.toLowerCase().includes(filters.URL.toLowerCase())) &&
        (filters?.Status === '' || item?.Status?.toLowerCase().includes(filters?.Status?.toLowerCase())) &&
        (filters.SubmittedDate === '' || item.Created.toLowerCase().includes(filters.SubmittedDate.toLowerCase()))
      );
    });
    const sortedData = filteredData.sort((a, b) => {
      if (sortConfig.key === 'SNo') {
        // Sort by index
        const aIndex = data.indexOf(a);
        const bIndex = data.indexOf(b);

        return sortConfig.direction === 'ascending' ? aIndex - bIndex : bIndex - aIndex;
      } else if (sortConfig.key) {
        // Sort by other keys
        const aValue = a[sortConfig.key] ? a[sortConfig.key].toLowerCase() : '';
        const bValue = b[sortConfig.key] ? b[sortConfig.key].toLowerCase() : '';

        if (aValue < bValue) {
          return sortConfig.direction === 'ascending' ? -1 : 1;
        }
        if (aValue > bValue) {
          return sortConfig.direction === 'ascending' ? 1 : -1;
        }
      }
      return 0;
    });
    return sortedData;
  };

  const filteredDelegateData = applyFiltersAndSorting(Delegatelistdata);

  const [currentPage, setCurrentPage] = React.useState(1);
  const itemsPerPage = 10;
  const totalPages = Math.ceil(filteredDelegateData.length / itemsPerPage);

  const handlePageChange = (pageNumber: any) => {
    if (pageNumber > 0 && pageNumber <= totalPages) {
      setCurrentPage(pageNumber);
    }
  };

  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentData = filteredDelegateData.slice(startIndex, endIndex);
  const goToAddForm = () => {
    sessionStorage.removeItem("delegateId")
    window.location.href = `${siteUrl}/SitePages/Delegation.aspx`;
  }



  //#region 
  const EditDelegate = (id: any) => {
    // debugger
    // setUseId(id)
    const encryptedId = encryptId(String(id));
    sessionStorage.setItem("delegateId", encryptedId)
    window.location.href = `${siteUrl}/SitePages/Delegation.aspx`;
  }
  //#endregion

  //#region 
  const DeleteDelegate = (id: any) => {
    Swal.fire({
      title: "Are you sure?",
      text: "You won't be able to revert this!",
      icon: "warning",
      showCancelButton: true,
      confirmButtonColor: "#3085d6",
      cancelButtonColor: "#d33",
      confirmButtonText: "Yes, delete it!"
    }).then(async (result) => {
      if (result.isConfirmed) {
        await DeleteDelegateAPI(sp, id)
        setDelegateData(prevBanners => prevBanners.filter(item => item.ID !== id));
        Swal.fire({
          title: "Deleted!",
          text: "Item has been deleted.",
          icon: "success"
        });

      }
    })
  }

  const [isOpen, setIsOpen] = React.useState(false);

  const toggleDropdown = () => {

    setIsOpen(!isOpen);

  };

  const [isOpenNews, setIsOpenNews] = React.useState(false);

  const toggleDropdownNews = () => {

    setIsOpenNews(!isOpenNews);

  };


  const handleExportClick = () => {

    const exportData = currentData.map((item, index) => ({

      'S.No.': startIndex + index + 1,

      'DelegateName': item.DelegateName.Title,

      'StartDate': item.StartDate,

      'EndDate': item.EndDate,

      'ActingFor': item.ActingFor.Title,

      'Status': item.Status,



    }));


    exportToExcel(exportData, 'Delegates');

  };

  const exportToExcel = (data: any[], fileName: string) => {

    const workbook = XLSX.utils.book_new();

    const worksheet = XLSX.utils.json_to_sheet(data);

    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    XLSX.writeFile(workbook, `${fileName}.xlsx`);

  };






  return (

    <div id="wrapper" ref={elementRef}>
      <div
        className="app-menu"
        id="myHeader">
        <VerticalSideBar _context={sp} />
      </div>
      <div className="content-page">
        <HorizontalNavbar _context={sp} siteUrl={siteUrl} />
        <div className="content " style={{ marginLeft: `${!useHide ? '240px' : '80px'}` }}>
          <div className="container-fluid  paddb">
            <div className="row">
              <div className="col-lg-3">
                <CustomBreadcrumb Breadcrumb={Breadcrumb} _context={sp} />
              </div>

              <div className="col-lg-9">
                <div className="d-flex flex-wrap align-items-center justify-content-end mt-3">
                  <div className="d-flex flex-wrap align-items-center justify-content-start">
                    <a href={`${siteUrl}/SitePages/settings.aspx`}>
                      <button type="button" className="btn btn-secondary me-1 waves-effect waves-light">
                        <FontAwesomeIcon icon={faArrowLeft} className="me-1" />
                        Back
                      </button>
                    </a>
                    <a href={`${siteUrl}/SitePages/Delegation.aspx`} onClick={() => goToAddForm()}>
                      <button type="button" className="btn btn-primary waves-effect waves-light" style={{ background: '#1fb0e5' }}>
                        <FontAwesomeIcon icon={faPlusCircle} className="me-1" />
                        Add
                      </button>
                    </a>
                  </div>
                </div>
              </div>
            </div>
            <div className="card cardCss mt-4 mb-0">
              <div className="card-body">
                <div id="cardCollpase4" className="collapse show">
                  <div className="table-responsive pt-0">
                    <table className="mtbalenew mt-0 table-centered table-nowrap table-borderless mb-0">
                      <thead>
                        <tr>
                          <th style={{
                            borderBottomLeftRadius: '0px', minWidth: '40px',
                            maxWidth: '40px', borderTopLeftRadius: '0px'
                          }}>
                            <div className="d-flex pb-2"
                              style={{ justifyContent: 'space-between' }}>
                              <span>S.No.</span>
                              <span onClick={() => handleSortChange('SNo')}>
                                <FontAwesomeIcon icon={faSort} />
                              </span>
                            </div>
                            <div className="bd-highlight">
                              <input
                                type="text"
                                placeholder="index"
                                onKeyDown={(e) => {
                                  if (e.key === 'Enter' && !e.shiftKey) {
                                    e.preventDefault(); // Prevents the new line in textarea
                                  }
                                }}
                                onChange={(e) => handleFilterChange(e, 'SNo')}
                                className="inputcss"
                                style={{ width: '100%' }}
                              />
                            </div>
                          </th>
                          <th style={{ minWidth: '120px', maxWidth: '120px' }}>
                            <div className="d-flex flex-column bd-highlight ">
                              <div className="d-flex pb-2" style={{ justifyContent: 'space-evenly' }}>
                                <span >Employee Name</span>  <span onClick={() => handleSortChange('Title')}><FontAwesomeIcon icon={faSort} /> </span></div>
                              <div className=" bd-highlight">
                                <input type="text" placeholder="Filter by Title" onChange={(e) => handleFilterChange(e, 'Title')}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter' && !e.shiftKey) {
                                      e.preventDefault(); // Prevents the new line in textarea
                                    }
                                  }}
                                  className='inputcss' style={{ width: '100%' }} />
                              </div>
                            </div>
                          </th>

                          <th style={{ minWidth: '120px', maxWidth: '120px' }}>
                            <div className="d-flex flex-column bd-highlight ">
                              <div className="d-flex pb-2" style={{ justifyContent: 'space-evenly' }}>
                                <span >Delegate Access To</span>  <span onClick={() => handleSortChange('Title')}><FontAwesomeIcon icon={faSort} /> </span></div>
                              <div className=" bd-highlight">
                                <input type="text" placeholder="Filter by Title" onChange={(e) => handleFilterChange(e, 'Title')}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter' && !e.shiftKey) {
                                      e.preventDefault(); // Prevents the new line in textarea
                                    }
                                  }}
                                  className='inputcss' style={{ width: '100%' }} />
                              </div>
                            </div>
                          </th>


                          <th style={{ minWidth: '146px', maxWidth: '146px' }}>
                            <div className="d-flex flex-column bd-highlight ">
                              <div className="d-flex pb-2" style={{ justifyContent: 'space-evenly' }}>
                                <span >Start Date</span>  <span onClick={() => handleSortChange('SubmittedDate')}><FontAwesomeIcon icon={faSort} /> </span></div>
                              <div className=" bd-highlight">
                                <input type="text" placeholder="Filter by SubmittedDate" onChange={(e) => handleFilterChange(e, 'SubmittedDate')}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter' && !e.shiftKey) {
                                      e.preventDefault(); // Prevents the new line in textarea
                                    }
                                  }}
                                  className='inputcss' style={{ width: '100%' }} />
                              </div>
                            </div>
                          </th>
                          <th style={{ minWidth: '146px', maxWidth: '146px' }}>
                            <div className="d-flex flex-column bd-highlight ">
                              <div className="d-flex pb-2" style={{ justifyContent: 'space-evenly' }}>
                                <span >Finish Date</span>  <span onClick={() => handleSortChange('SubmittedDate')}><FontAwesomeIcon icon={faSort} /> </span></div>
                              <div className=" bd-highlight">
                                <input type="text" placeholder="Filter by SubmittedDate" onChange={(e) => handleFilterChange(e, 'SubmittedDate')}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter' && !e.shiftKey) {
                                      e.preventDefault(); // Prevents the new line in textarea
                                    }
                                  }}
                                  className='inputcss' style={{ width: '100%' }} />
                              </div>
                            </div>
                          </th>

                          <th style={{ minWidth: '80px', maxWidth: '80px' }}>
                            <div className="d-flex flex-column bd-highlight ">
                              <div className="d-flex pb-2" style={{ justifyContent: 'space-evenly' }}>
                                <span >Status</span>  <span onClick={() => handleSortChange('Status')}><FontAwesomeIcon icon={faSort} /> </span></div>
                              <div className=" bd-highlight">
                                <input type="text" placeholder="Filter by Status" onChange={(e) => handleFilterChange(e, 'Status')}
                                  onKeyDown={(e) => {
                                    if (e.key === 'Enter' && !e.shiftKey) {
                                      e.preventDefault(); // Prevents the new line in textarea
                                    }
                                  }}
                                  className='inputcss' style={{ width: '100%' }} />
                              </div>
                            </div>
                          </th>
                          <th style={{ textAlign: 'center', minWidth: '80px', maxWidth: '80px', borderBottomRightRadius: '0px', borderTopRightRadius: '0px' }}> <div className="d-flex flex-column bd-highlight pb-2">

                            <div className="d-flex  pb-2" style={{ justifyContent: 'space-evenly' }}>  <span >Action</span> <div className="dropdown">

                              <FontAwesomeIcon icon={faEllipsisV} onClick={toggleDropdown} size='xl' />

                            </div>

                            </div>

                            {/* <div className=" bd-highlight">   <div id="myDropdown" className={`dropdown-content ${isOpen ? 'show' : ''}`}>
                            
                                                              <div onClick={handleExportClick} className="" >
                            
                                                                <FontAwesomeIcon icon={faFileExport} />  Export
                            
                                                              </div>
                            
                                                            </div></div> */}


                          </div>

                            <div style={{ height: '32px' }}></div>

                          </th>
                          {/* <th style={{ borderBottomRightRadius: '0px', minWidth: '50px', maxWidth: '50px', borderTopRightRadius: '0px' }}>
                              <div className="d-flex flex-column bd-highlight pb-2">
                                <div className="d-flex  pb-2" style={{ justifyContent: 'space-evenly' }}>  <span >Action</span> <div className="dropdown">
                                  <FontAwesomeIcon icon={faEllipsisV} onClick={toggleDropdownNews} size='xl' />
                                </div>
                                </div>
                                <div className=" bd-highlight">   <div id="myDropdown" className={`dropdown-content ${isOpenNews ? 'showNews' : ''}`}>
                                  <div onClick={handleExportClick} className="" >
                                    <FontAwesomeIcon icon={faFileExport} />  Export
                                  </div>
                                </div></div>
  
                              </div>
                              <div style={{ height: '32px' }}></div>
                            </th> */}
                        </tr>
                      </thead>
                      {console.log("currentdataaa", currentData)}
                      <tbody style={{ maxHeight: '5000px' }}>
                        {currentData.length === 0 ?
                          (
                            <div className="no-results" style={{ display: 'flex', justifyContent: 'center' }}>No results found</div>
                          )
                          :
                          currentData.map((item, index) => {
                            // const ImageUrl = item.BannerImage == undefined || item.BannerImage == null ? "" : JSON.parse(item.BannerImage);
                            return (
                              <tr key={index}>
                                <td style={{ minWidth: '40px', maxWidth: '40px' }}><div style={{ marginLeft: '10px' }} className='indexdesign'> {index + 1}</div>  </td>
                                <td style={{ minWidth: '120px', maxWidth: '120px' }}>{item.DelegateName.Title}</td>
                                <td style={{ minWidth: '120px', maxWidth: '120px' }}>{item.ActingFor.Title}</td>

                                <td style={{ minWidth: '146px', maxWidth: '146px', textAlign: 'center' }}>
                                  {/* <div className='btn btn-light newlight'> {moment(item.StartDate).format("DD-MMM-YYYY")} </div> */}
                                  <div className='btn btn-light newlight'> {new Date(item.StartDate).toLocaleString('en-GB', {
                                    day: '2-digit',
                                    month: '2-digit',
                                    year: 'numeric',
                                    hour: '2-digit',
                                    minute: '2-digit',
                                    hour12: true
                                  })} </div>
                                </td>
                                <td style={{ minWidth: '146px', maxWidth: '146px', textAlign: 'center' }}>
                                  {/* <div className='btn btn-light newlight'> {moment(item.EndDate).format("DD-MMM-YYYY")} </div> */}
                                  <div className='btn btn-light newlight'> {new Date(item.EndDate).toLocaleString('en-GB', {
                                    day: '2-digit',
                                    month: '2-digit',
                                    year: 'numeric',
                                    hour: '2-digit',
                                    minute: '2-digit',
                                    hour12: true
                                  })} </div>
                                </td>
                                <td style={{ minWidth: '80px', maxWidth: '80px', textAlign: 'center' }}>  <div className='btn btn-status newlight'> {item.Status} </div> </td>
                                <td style={{ minWidth: '80px', maxWidth: '80px' }} className="ng-binding">
                                  <div className="d-flex  pb-0" style={{ justifyContent: 'center', gap: '5px' }}>
                                    <span > <a className="action-icon text-primary" onClick={() => EditDelegate(item.ID)}>
                                      <img src={require('../../../CustomAsset/edit.png')} />
                                    </a> </span >
                                    <span>   <a className="action-icon text-danger" onClick={() => DeleteDelegate(item.ID)}>
                                      <img src={require('../../../CustomAsset/del.png')} />
                                    </a> </span>
                                  </div>
                                </td>
                              </tr>
                            )
                          })
                        }
                      </tbody>
                    </table>


                    <nav className="pagination-container">
                      <ul className="pagination">
                        <li className={`page-item ${currentPage === 1 ? 'disabled' : ''}`}>
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
                            className={`page-item ${currentPage === num + 1 ? 'active' : ''}`}
                          >
                            <a
                              className="page-link"
                              onClick={() => handlePageChange(num + 1)}
                            >
                              {num + 1}
                            </a>
                          </li>
                        ))}
                        <li className={`page-item ${currentPage === totalPages ? 'disabled' : ''}`}>
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
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>
        {/* End table content */}
        {/* End container */}
      </div>
    </div>

  )

}


const DelegationMaster: React.FC<IDelegationMasterProps> = (props) => {
  return (
    <Provider>
      <DelegationMasterContext props={props} />
    </Provider>
  )
}


export default React.memo(DelegationMaster);
