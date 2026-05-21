import React, { useEffect, useState } from 'react'
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faChevronRight } from '@fortawesome/free-solid-svg-icons';
import "bootstrap/dist/css/bootstrap.min.css";
import "../CustomBreadcrumb/CustomBreadcrumb.scss"
import { SPFI } from "@pnp/sp/presets/all";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { getbreadcrumbhomeurlfunction } from '../../APISearvice/getbreadcrumbhomeurl';
import { getSP } from '../../webparts/essaApproval/loc/pnpjsConfig';
const CustomBreadcrumb = ({Breadcrumb , _context}) => {
const sp = getSP();
     const [urlbreadcrumbhomeurl, setURLbreadcrumbhomeurl] = useState('');
 React.useEffect( () => {
      const asyncFunction = async () => {
        //alert("asyncFunction");
    setURLbreadcrumbhomeurl(await getbreadcrumbhomeurlfunction(_context))
     //alert(urlbreadcrumbhomeurl + "urlbreadcrumbhomeurl")
         const url = await getbreadcrumbhomeurlfunction(_context);
         let url2 = url
    setURLbreadcrumbhomeurl(url); // This updates state
    //alert(url + " url from function"); // ✅ T
  };
  asyncFunction();
  }, []);
 
    return (
        <div className=''>
            <h4 className="page-title fw-bold mb-0 font-20">{Breadcrumb[1].ChildComponent}</h4>
            <ol className="breadcrumb mb-0">
                <li className="breadcrumb-item"><a href={urlbreadcrumbhomeurl}>{Breadcrumb[0].MainComponent}</a></li>
                <li className="breadcrumb-item pt-arr">
                    <FontAwesomeIcon
                        className="arrow-left"
                        icon={faChevronRight} size='xs' style={{color:'#6c757d'}}/>
                </li>
                <li className="breadcrumb-item active"><a href={Breadcrumb[1].ChildComponentURl}>{Breadcrumb[1].ChildComponent}</a></li>
            </ol>
        </div>
    )
}
 
export default CustomBreadcrumb

// import React from 'react'
// import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
// import { faChevronRight } from '@fortawesome/free-solid-svg-icons';
// import "bootstrap/dist/css/bootstrap.min.css";
// import "../CustomBreadcrumb/CustomBreadcrumb.scss"
// const CustomBreadcrumb = ({Breadcrumb}) => {
//     return (
//         <div className=''>
//             <h4 className="page-title fw-bold mb-0 font-20">{Breadcrumb[1].ChildComponent}</h4>
//             <ol className="breadcrumb mb-2">
//                 <li className="breadcrumb-item"><a href={Breadcrumb[0].MainComponentURl}>{Breadcrumb[0].MainComponent}</a></li>
//                 <li className="breadcrumb-item pt-arr">
//                     <FontAwesomeIcon
//                         className="arrow-left"
//                         icon={faChevronRight} size='xs' style={{color:'#6c757d'}}/>
//                 </li>
//                 <li className="breadcrumb-item active"><a href={Breadcrumb[1].ChildComponentURl}>{Breadcrumb[1].ChildComponent}</a></li>
//             </ol>
//         </div>
//     )
// }

// export default CustomBreadcrumb
