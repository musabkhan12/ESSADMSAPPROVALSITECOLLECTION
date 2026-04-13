import { escape } from "@microsoft/sp-lodash-subset";

import { spfi, SPFI } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { IWeb } from "@pnp/sp/webs/types";
import React, { useState } from "react";

import { updateItemApproval } from "../../APISearvice/ApprovalService";

import {getSP} from '../../webparts/essaApproval/loc/pnpjsConfig';

import { WebPartContext } from "@microsoft/sp-webpart-base";

import "bootstrap/dist/css/bootstrap.min.css";


// import "../../../CustomCss/mainCustom.scss";


// // import "../components/MyApproval.scss";


// import "bootstrap/dist/js/bootstrap.bundle.min.js";


// import "../../../CustomJSComponents/CustomTable/CustomTable.scss";


// import "../../verticalSideBar/components/VerticalSidebar.scss";


// import "./WorkflowAuditHistory.scss"


export interface IWorkflowAuditHistoryProps {

  ContentItemId: any;
  
  ContentType: string;
  currenttab?: string;
  ctx: WebPartContext;
  SiteBaseURL: string;
  listName: string;
}


export const WorkflowAuditHistory = (props: IWorkflowAuditHistoryProps) => {
   console.log("WorkflowAuditHistory", props);  
  //  console.log("WorkflowAuditHistory", JSON.stringify(props));  
  // alert(props.SiteBaseURL + "props.SiteBaseUR")
  // alert(props.listName + "props.listName")
  const [AuditHistoryRows, setAuditHistoryRows] = React.useState([]);

  const [Loading, setLoading] = React.useState(false);

  const [IsHistoryData, setIsHistoryData] = React.useState(false);

  const siteUrl = props.ctx.pageContext.site.absoluteUrl;


  const sp = getSP(props.ctx);


  React.useEffect(() => {

    // if (props.ContentItemId && AuditHistoryRows.length == 0 && !IsHistoryData) {
      const externalWebUrl = "https://officeindia.sharepoint.com/sites/DigitalServices";

      const externalSp: SPFI = spfi(externalWebUrl).using(SPFx(props.ctx));
      if(externalSp){
        console.log("externalSp", externalSp);
      }else{
        // alert("externalSp is null");
      }
  
      getAllAPI()

    // }

  },[])

  const getAllAPI = async () => {
    const externalWebUrl = "https://officeindia.sharepoint.com/sites/DigitalServices";

      const externalSp: SPFI = spfi(externalWebUrl).using(SPFx(props.ctx));
      console.log("externalSp", externalSp);
      
    // alert("getAllAPI ContentType"+props.ContentType);
    let listname = "";
    let columnname = "";
    let newlistname = await sp.web.lists.getByTitle("AllApprovalLists").items
      .filter(`IsActive eq 'Yes' and ProcessName eq '${props.ContentType}'`).orderBy("Created", false)();
    let finallistname = newlistname[0].Title;
    //  alert("finallistname"+finallistname);
    //  alert(props.ContentType + "props.ContentType");
    if (props.ContentItemId) {
      //  alert("ContentType AllApprovalLists"+ props.ContentType);
      switch (props.ContentType) {
        case "Payment Request Process":
          listname = finallistname;
          break;
          case "IT Demand Request":
            listname = finallistname;
            console.log("listname for IT Demand Request:", listname);
            break;
        case "IT Demand Proposal Process":
          listname = finallistname;
          break;
     
        case "Supplier Guarantee":
          listname = finallistname;
          break;
        case "InterOfficeApprovalList":
          listname = finallistname;
          break;
        case "Capex Requisition":
          listname = finallistname;
          break;
        case "IT Project Charter":
          listname = finallistname;
          break;
        case "PettyCash":
          listname = finallistname;
          break;
        case "Payment Certificate":
          listname = finallistname;
          break;
        case "HOTO Request Task":
          listname = finallistname;
          break;
        case "IT Demand Proposal Process":
          listname = finallistname;
          break;
        case "Sales Order Summary":
          listname = finallistname;
          break;
        case "IT Purchase Request":
          listname = finallistname;
          break;
        case "PAF Task":
          listname = finallistname;
          break;
        case "Return From Vacation":
          listname = finallistname;
          break;
        default:
      }
      // alert("listname last"+listname);
      setLoading(true);
      const normalizedListname = listname.trim().toLowerCase();
      console.log("Normalized listname:", normalizedListname);
      if (props.currenttab == "Automation") {
        if(listname === 'PaymentRequestApprovalList'){
        await sp.web.lists.getByTitle(listname).items
          .select("*,RequestedBy/Id,RequestedBy/Title,ApprovedBy/Id,ApprovedBy/Title" , "PaymentRequestID/Id" , "PaymentRequestID" , "AssignedTo/Title" , "Level" , "Status" , "Remark" , "Created" , "Modified" )
          .expand("AssignedTo", "ApprovedBy", "RequestedBy" , "PaymentRequestID")
          .filter(`PaymentRequestID/Id eq ${props.ContentItemId}`)
          .orderBy('Created')().then(datarows => {
            console.log("datarows", datarows);
            console.log("datarows", datarows);
            if (datarows.length == 0) {
              setIsHistoryData(true);
              setLoading(false);
            }
            if (datarows.length > 0) {
              setLoading(false);
              setIsHistoryData(true);
            }
            setAuditHistoryRows(datarows);
          })
        }
        // else if(listname === 'ITDemandProposalApprovalList'){
        //   await sp.web.lists.getByTitle(listname).items
        //   .select("*,RequestedBy/Id,RequestedBy/Title,ApprovedBy/Id,ApprovedBy/Title" , "DemandProposalRequestID/Id" , "DemandProposalRequestID" , "AssignedTo/Title" , "Step" , "Status" , "Remark" , "Created" , "Modified" )
        //   .expand("AssignedTo", "ApprovedBy", "RequestedBy" , "DemandProposalRequestID")
        //   .filter(`DemandProposalRequestID/Id eq ${props.ContentItemId}`)
        //   .orderBy('Created')().then(datarows => {
        //     const mappedRows = datarows.map(item => ({
        //       ...item,
        //       Level: item.Step // Copy Step value to a new Level property
        //     }));
        //     console.log("datarows", datarows);
        //     console.log("datarows", datarows);
        //     if (mappedRows.length === 0) {
        //       setIsHistoryData(true);
        //       setLoading(false);
        //     } else {
        //       setLoading(false);
        //       setIsHistoryData(true);
        //     }
        
        //     setAuditHistoryRows(mappedRows);

        //     // if (datarows.length == 0) {
        //     //   setIsHistoryData(true);
        //     //   setLoading(false);
        //     // }
        //     // if (datarows.length > 0) {
        //     //   setLoading(false);
        //     //   setIsHistoryData(true);
        //     // }
        //     // setAuditHistoryRows(datarows);
        //   })
        // }
        //  else  if(listname === 'ITDemandRequestApprovalList'){
        // await sp.web.lists.getByTitle(listname).items
        //   .select("*,RequestedBy/Id,RequestedBy/Title,ApprovedBy/Id,ApprovedBy/Title" , "ITDemandRequestID/Id" , "ITDemandRequestID" , "AssignedTo/Title" , "Level" , "Status" , "Remark" , "Created" , "Modified" )
        //   .expand("AssignedTo", "ApprovedBy", "RequestedBy" , "ITDemandRequestID")
        //   .filter(`ITDemandRequestID/Id eq ${props.ContentItemId}`)
        //   .orderBy('Created')().then(datarows => {
        //     console.log("datarows", datarows);
        //     console.log("datarows", datarows);
        //     if (datarows.length == 0) {
        //       setIsHistoryData(true);
        //       setLoading(false);
        //     }
        //     if (datarows.length > 0) {
        //       setLoading(false);
        //       setIsHistoryData(true);
        //     }
        //     setAuditHistoryRows(datarows);
        //   })
        // }
         else  if(listname === 'SupplierGuaranteeApprovalList'){
        await sp.web.lists.getByTitle(listname).items
          .select("*,RequestedBy/Id,RequestedBy/Title,ApprovedBy/Id,ApprovedBy/Title" , "SupplierGuaranteeRequestListID/Id" , "SupplierGuaranteeRequestListID" , "AssignedTo/Title"  , "Status" , "Remarks" , "Created" , "Modified" )
          .expand("AssignedTo", "ApprovedBy", "RequestedBy" , "SupplierGuaranteeRequestListID")
          .filter(`SupplierGuaranteeRequestListID/Id eq ${props.ContentItemId}`)
          .orderBy('Created')().then(datarows => {
            console.log("datarows", datarows);
            console.log("datarows", datarows);
            if (datarows.length == 0) {
              setIsHistoryData(true);
              setLoading(false);
            }
            if (datarows.length > 0) {
              setLoading(false);
              setIsHistoryData(true);
            }
            setAuditHistoryRows(datarows);
          })
        }
//          else if(normalizedListname === 'itdemandrequestapprovallist'){
//           alert("ITDemandProposalApprovalList or CapexRequisitionApprovalList"  + listname);


//           alert("Fetching data for: " + listname);

//           // 🌐 Define dynamic lookup columns per list
//           const lookupColumnMap: { [key: string]: string } = {
         
//              ITDemandRequestApprovalList:"ITDemandRequestID",
//              // "Return From Vacation":"MasterID",

//           };
        
//           const lookupColumn = lookupColumnMap[listname];
        
//           // ✅ Define reusable select/expand fields
//           const selectFields = [
//             "*",
//             "RequestedBy/Id", "RequestedBy/Title",
//             "ApprovedBy/Id", "ApprovedBy/Title",
//             `ITDemandRequestID/Id`, `ITDemandRequestID/Title`,
//             "AssignedTo/Title",
//             "Created", "Modified"
//           ].join(",");
        
//           const expandFields = [
//             "RequestedBy",
//             "ApprovedBy",
//             "ITDemandRequestID",
//             "AssignedTo"
//           ].join(",");
//         // 🔁 Reusable fetch function
// const fetchApprovalListData = async ({
//  siteUrl,
//  listName,
//  filterLookupColumn,
//  filterLookupId,
//  selectFields,
//  expandFields,
// }: {
//  siteUrl: string;
//  listName: string;
//  filterLookupColumn: string;
//  filterLookupId: number;
//  selectFields: string;
//  expandFields: string;
// }) => {
//  alert("Fetching data for: " + siteUrl +"siteUrl" + listName + "listName" + filterLookupColumn + "filterLookupColumn" + filterLookupId + "filterLookupId" + selectFields + "selectFields" + expandFields + "expandFields");
//  console.log("Fetching data for:", listName);
//  console.log("Filter Lookup Column:", filterLookupColumn);
//  console.log("Filter Lookup ID:", filterLookupId);
//  console.log("Select Fields:", selectFields);
//  console.log("Expand Fields:", expandFields);
//  // Construct the API URL
//  const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items` +
//    `?$select=${selectFields}` +
//    `&$expand=${expandFields}` +
//    `&$filter=${filterLookupColumn}/Id eq ${filterLookupId}` +
//    `&$orderby=Created asc`;

//  console.log("Fetching data from:", apiUrl);

//  try {
//    const response = await fetch(apiUrl, {
//      method: 'GET',
//      headers: {
//        'Accept': 'application/json',
//      },
//    });

//    if (!response.ok) {
//      throw new Error(`HTTP error! Status: ${response.status}`);
//    }

//    const data = await response.json();
//    return data.value;
//  } catch (error) {
//    console.error(`Error fetching data for ${listName}:`, error);
//    return [];
//  }
// };
//           const result = await fetchApprovalListData({
//             siteUrl: props.SiteBaseURL,
//             listName: listname,
//             filterLookupColumn: lookupColumn,
//             filterLookupId: props.ContentItemId,
//             selectFields,
//             expandFields
//           });
        
//           if (result.length === 0) {
//             setIsHistoryData(true);
//             setLoading(false);
//           } else {
//             const mappedRows = result.map((item: any) => ({
//               ...item,
//               Level: item.Step || "" // optional mapping
//             }));
//             console.log('Mapped Rows:', mappedRows);
//             setAuditHistoryRows(mappedRows);
//             setIsHistoryData(true);
//             setLoading(false);
//           }
//         }
        else if(
          normalizedListname === 'itdemandrequestapprovallist' ||
           listname === 'ITDemandProposalApprovalList' || 
          listname === 'CapexRequisitionApprovalList' ||
          listname === 'ITPurcahseApprovalList' ||
          listname === 'InterOfficeApprovalList' ||
          listname === 'ITprojectcharterApprovalList' ||
          listname === 'PettyCashApprovalList' ||
          listname === 'PaymentCertificateFormApprovalList' ||
          listname === 'HOTO Request Task' ||
          listname === 'SalesOrderSummaryApprovalList' ||
          listname === 'PAF Task' 
      
         ){

          //  alert("ITDemandProposalApprovalList or CapexRequisitionApprovalList"  + listname);


          //  alert("Fetching data for: " + listname);

           // 🌐 Define dynamic lookup columns per list
           const lookupColumnMap: { [key: string]: string } = {
             ITDemandProposalApprovalList: "DemandProposalRequestID",
             CapexRequisitionApprovalList: "CapexRequestID",
             ITPurcahseApprovalList: "ITPurchaseRequestID",
             InterOfficeApprovalList: "InterOfficeRequestListID",
             ITprojectcharterApprovalList:"projectcharterRequestID",
             PettyCashApprovalList:"PettyCashRequestListID",
             PaymentCertificateFormApprovalList :"PaymentCertificateFormReqId",
             "HOTO Request Task":"MasterID",
             SalesOrderSummaryApprovalList:"SalesOrderSummaryRequestListID",
              "PAF Task":"MasterID",
              itdemandrequestapprovallist:"ITDemandRequestID",
              ITDemandRequestApprovalList:"ITDemandRequestID",
              // "Return From Vacation":"MasterID",

           };
            let lookupColumn = "";
           if(normalizedListname === 'itdemandrequestapprovallist'){
             lookupColumn = lookupColumnMap[normalizedListname];
           }else{
             lookupColumn = lookupColumnMap[listname];
           }
          
         
           // ✅ Define reusable select/expand fields
           const selectFields = [
             "*",
             "RequestedBy/Id", "RequestedBy/Title",
             "ApprovedBy/Id", "ApprovedBy/Title",
             `${lookupColumn}/Id`, `${lookupColumn}/Title`,
             "AssignedTo/Title",
             "Created", "Modified"
           ].join(",");
         
           const expandFields = [
             "RequestedBy",
             "ApprovedBy",
             lookupColumn,
             "AssignedTo"
           ].join(",");
         // 🔁 Reusable fetch function
const fetchApprovalListData = async ({
  siteUrl,
  listName,
  filterLookupColumn,
  filterLookupId,
  selectFields,
  expandFields,
}: {
  siteUrl: string;
  listName: string;
  filterLookupColumn: string;
  filterLookupId: number;
  selectFields: string;
  expandFields: string;
}) => {
  // alert("Fetching data for: " + siteUrl +"siteUrl" + listName + "listName" + filterLookupColumn + "filterLookupColumn" + filterLookupId + "filterLookupId" + selectFields + "selectFields" + expandFields + "expandFields");
  console.log("Fetching data for:", listName);
  console.log("Filter Lookup Column:", filterLookupColumn);
  console.log("Filter Lookup ID:", filterLookupId);
  console.log("Select Fields:", selectFields);
  console.log("Expand Fields:", expandFields);
  // Construct the API URL
  const apiUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')/items` +
    `?$select=${selectFields}` +
    `&$expand=${expandFields}` +
    `&$filter=${filterLookupColumn}/Id eq ${filterLookupId}` +
    `&$orderby=Created asc`;

  console.log("Fetching data from:", apiUrl);

  try {
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
      },
    });

    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }

    const data = await response.json();
    return data.value;
  } catch (error) {
    console.error(`Error fetching data for ${listName}:`, error);
    return [];
  }
};
           const result = await fetchApprovalListData({
             siteUrl: props.SiteBaseURL,
             listName: listname,
             filterLookupColumn: lookupColumn,
             filterLookupId: props.ContentItemId,
             selectFields,
             expandFields
           });
         
           if (result.length === 0) {
             setIsHistoryData(true);
             setLoading(false);
           } else {
             const mappedRows = result.map((item: any) => ({
               ...item,
               Level: item.Step || "" // optional mapping
             }));
             console.log('Mapped Rows:', mappedRows);
             setAuditHistoryRows(mappedRows);
             setIsHistoryData(true);
             setLoading(false);
           }
          //  this below code was working fine but need some updates
                     // const apiUrl = `${props.SiteBaseURL}/_api/web/lists/getbytitle('${listname}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail&$expand=Requestor_x0020_Name&$filter=`;
  //         const apiUrl = `${props.SiteBaseURL}/_api/web/lists/getbytitle('${listname}')/items` +
  // `?$select=*,RequestedBy/Id,RequestedBy/Title,ApprovedBy/Id,ApprovedBy/Title,` +
  // `DemandProposalRequestID/Id,DemandProposalRequestID/Title,AssignedTo/Title,Created,Modified` +
  // `&$expand=RequestedBy,ApprovedBy,DemandProposalRequestID,AssignedTo` +
  // `&$filter=DemandProposalRequestID/Id eq ${props.ContentItemId}` +
  // `&$orderby=Created asc`;
  //             //  get data 
  //             try {
  //               const response = await fetch(apiUrl, {
  //                 method: 'GET',
  //                 headers: {
  //                   'Accept': 'application/json',
  //                 },
  //               });
            
  //               if (!response.ok) {
  //                 console.log('Error fetching data:', response.statusText);
  //                 throw new Error(`HTTP error! Status: ${response.status}`);
  //               }
            
  //               const data = await response.json();
  //               console.log('Fetched ITDemandProposalApprovalList Data:', data.value);
            
  //               if (data.value.length === 0) {
  //                 setIsHistoryData(true);
  //                 setLoading(false);
  //               } else {
  //                 // Optional: map Step to Level if needed
  //                 const mappedRows = data.value.map((item:any) => ({
  //                   ...item,
  //                   Level: item.Step // If `Step` exists and you want to map it to `Level`
  //                 }));
  //                 console.log('Mapped Rows:', mappedRows);
  //                 setAuditHistoryRows(mappedRows);
  //                 setIsHistoryData(true);
  //                 setLoading(false);
  //               }
  //             } catch (error) {
  //               console.error('Error fetching ITDemandProposalApprovalList data:', error);
  //               setLoading(false);
  //             }
        }
//         else if(listname === 'PAF Task' ||
//            listname === 'PAF Task List' || 
//            listname === 'DAR Task' ||
//             listname === 'InterOfficeApprovalList' ||
//              listname === 'CapexRequisitionApprovalList' ||
//               listname === 'ITDemandRequestApprovalList' ||
//                listname === 'ITProjectCharterApprovalList' ||
//                 listname === 'PettyCashApprovalList' || 
//                 listname === 'PaymentCertificateFormApprovalList' ||
//                  listname === 'HOTO Request Task' ||
//                   listname === 'ITDemandProposalApprovalList' ||
//                     listname === 'SalesOrderSummaryApprovalList' || 
//                     listname === 'ITPurcahseApprovalList' || 
//                     listname === 'Return From Vacation' 
//                 ){
//                   alert('listname'+listname + "DigitalServices");
//                   const lookupColumnMap = {
//                     "PAF Task": "PaymentRequestID",
//                     "PAF Task List": "InvoiceRequestID",
//                     "DAR Task": "TravelRequestID",
//                     "InterOfficeApprovalList": "TravelRequestID",
//                     "CapexRequisitionApprovalList": "TravelRequestID",
//                     "ITDemandRequestApprovalList": "TravelRequestID",
//                     "ITProjectCharterApprovalList": "TravelRequestID",
//                     "PettyCashApprovalList": "TravelRequestID",
//                     "PaymentCertificateFormApprovalList": "TravelRequestID",
//                     "HOTO Request Task": "TravelRequestID",
//                     "ITDemandProposalApprovalList": "TravelRequestID",
//                     "SalesOrderSummaryApprovalList": "TravelRequestID",
//                     "ITPurcahseApprovalList": "TravelRequestID",
//                     "Return From Vacation": "TravelRequestID",
//                     // Add more mappings here as needed
//                   };
                  
//                   // Get the dynamic column name based on the list
//                   const isLookupColumnKey = (key: string): key is keyof typeof lookupColumnMap => {
//                     console.log("lookupColumnMap", lookupColumnMap);
//                     console.log("lookupColumnMap key", key);
//                     return key in lookupColumnMap;
//                   }
//                   console.log("lookupColumnMap func", isLookupColumnKey(listname));
//           // const apiUrl = `${props.SiteBaseURL}/_api/web/lists/getbytitle('${listname}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail&$expand=Requestor_x0020_Name&$filter=`;
// `${props.SiteBaseURL}/_api/web/lists/getbytitle('${listname}')/items` +
//   `?$select=*,RequestedBy/Id,RequestedBy/Title,ApprovedBy/Id,ApprovedBy/Title,` +
//   `PaymentRequestID/Id,PaymentRequestID/Title,AssignedTo/Title,Created,Modified` +
//   `&$expand=RequestedBy,ApprovedBy,PaymentRequestID,AssignedTo` +
//   `&$filter=PaymentRequestID/Id eq ${props.ContentItemId}` +
//   `&$orderby=Created asc`;
//         }
      } else {
        sp.web.lists.getByTitle("ARGMyRequest").items
          .select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title")
          .expand("Approver,Requester")
          .filter('ContentId eq ' + props.ContentItemId + "and ProcessName eq '" + props.ContentType + "'")
          .orderBy('Created')().then(datarows => {

            if (datarows.length == 0) {

              setIsHistoryData(true);

              setLoading(false);

            }

            if (datarows.length > 0) {

              setLoading(false);

              setIsHistoryData(true);

            }

            setAuditHistoryRows(datarows);


          })
      }
    }

  }


  return (


    <div className="card cardCss mb-0 mt-3">


      <div className="card-body">


        <div id="cardCollpase4" className="collapse show">

          <h3 className="font-16 mb-1 fw-bold text-dark">Audit History</h3>
          <p className="font-14 mb-3 text-muted">Below table describes the status of approval at various level </p>

          {console.log("AuditHistoryRows", AuditHistoryRows)}

          <div className="table-responsive pt-0">


            <table


              className="mtbalenew  table-centered  thead-light mb-0"


              style={{ position: "relative" }}


            >

              <thead>

                <tr>

                  <th style={{ minWidth: '60px', maxWidth: '60px' }}>

                    S.No.

                  </th>

                  <th style={{ minWidth: '60px', maxWidth: '60px' }}>

                    Level

                  </th>

                  <th style={{ minWidth: '110px', maxWidth: '110px' }}>

                    Assigned To

                  </th>

                  <th style={{ minWidth: '110px', maxWidth: '110px' }}>

                    Requester Name

                  </th>

                  <th style={{ minWidth: '110px', maxWidth: '110px' }}>

                    Requested Date

                  </th>

                  <th style={{ minWidth: '110px', maxWidth: '110px' }}>

                    Action Taken By

                  </th>

                  <th style={{ minWidth: '110px', maxWidth: '110px' }}>

                    Action Taken On

                  </th>


                  <th style={{ minWidth: '110px', maxWidth: '110px' }}>

                    Remarks

                  </th>

                  <th style={{ minWidth: '70px', maxWidth: '70px' }}>

                    Status

                  </th>


                </tr>

              </thead>

              {Loading && AuditHistoryRows.length == 0 ? (//chhaya

                // Show loader when loading is true

                (!IsHistoryData ?

                  <div className="spinner-border text-primary" role="status">

                    <span className="sr-only">Loading...</span>

                  </div> :

                  <div >

                    <span className="sr-only">No Record has been found</span>

                  </div>)

              ) : (

                <tbody>


                  {


                    AuditHistoryRows.map((row: any, index: number) =>

                      <tr>

                        <td style={{ minWidth: '60px', maxWidth: '60px' }}> {index + 1}</td>

                        {/* <td> {(row.LevelId == 0) ? "Initiator" : `Level ${row.LevelId}`}</td> */}
                        <td style={{ minWidth: '60px', maxWidth: '60px' }}> {row.Level}</td>

                        <td style={{ minWidth: '110px', maxWidth: '110px' }}> {row.AssignedTo.Title}</td>

                        <td style={{ minWidth: '110px', maxWidth: '110px' }}> {row?.RequestedBy?.Title}</td>

                        <td  style={{ minWidth: '110px', maxWidth: '110px' }} title={
    row?.Created ? 
    new Date(row.Created).toLocaleString('en-GB', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    }) : ""
  }>
                           {/* {(new Date(row.Created)).toLocaleString()} */}
                           {new Date(row?.Created).toLocaleString('en-GB', {
                                                  month: '2-digit',
                                                  day: '2-digit',
                                                  year: 'numeric',
                                                  hour: '2-digit',
                                                  minute: '2-digit',
                                                  //second: '2-digit',
                                                  hour12: true
                                                })}

                        </td>

                        <td style={{ minWidth: '110px', maxWidth: '110px' }}> {(row.Status != 'Pending') ? (row.AssignedTo.Title) : ""}</td>

                        <td  style={{ minWidth: '110px', maxWidth: '110px' }} title={
    row?.Modified ? 
    new Date(row.Modified).toLocaleString('en-GB', {
      month: '2-digit',
      day: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
      hour12: true
    }) : ""
  }>{(row.Status != 'Pending') ? 
                        (
                          // (new Date(row.Modified)).toLocaleString()
                          (new Date(row?.Modified).toLocaleString('en-GB', {
                            month: '2-digit',
                            day: '2-digit',
                            year: 'numeric',
                            hour: '2-digit',
                            minute: '2-digit',
                            //second: '2-digit',
                            hour12: true
                          }))
                      ) 

                        
                        : ""}</td>

                        <td style={{ minWidth: '110px', maxWidth: '110px' }}> {row.Remark}</td>

                        <td style={{ minWidth: '70px', maxWidth: '70px' }}> <div className="btn  btn-status">{row.Status}</div> </td>


                      </tr>


                    )


                  }

                </tbody>

              )

              }

            </table>

          </div>

        </div>

      </div>

    </div>


  )


}
