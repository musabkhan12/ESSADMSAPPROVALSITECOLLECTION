import { escape } from "@microsoft/sp-lodash-subset";

import React, { useState } from "react";
import { updateItemApproval } from "../../APISearvice/ApprovalService";
import {getSP} from '../../webparts/essaApproval/loc/pnpjsConfig';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import Swal from "sweetalert2";

export interface IWorkflowActionProps
{
    currentItem:any;
    ctx:WebPartContext;
    DisableApproval?:boolean;
    DisableRework?:boolean;
    DisableReject?:boolean;
    DisableCancel?:boolean;

}

export const WorkflowAction=(props: IWorkflowActionProps) => {

    
    const siteUrl = props.ctx.pageContext.site.absoluteUrl;

    const sp=getSP(props.ctx);

    const [formData, setFormData] = React.useState({
        Remark: '',    
      })

    const onChange = (name: string, value: string) => {

        debugger
    
        setFormData((prevData) => ({
    
          ...prevData,
    
          [name]: value,
    
        }));
    
      };
      const handleCancel = () => {   
        window.location.href = `${siteUrl}/SitePages/MyApprovals.aspx`;
      }
      const handleFromSubmit = async (e: React.MouseEvent<HTMLButtonElement, MouseEvent>, Status: string) => {

        e.preventDefault();
    
      
        const postPayload = {
    
          Remark: formData.Remark,
    
          Status: Status
    
        };
    
        console.log(postPayload);
       let confirmation,resultmessage="";

       if(Status=='Approved') 
        {
          confirmation="Do you want to approve this request?";
          resultmessage='Approved successfully.'

        }
       else if(Status=='Rejected') {
         confirmation="Do you want to reject this request?";
         resultmessage='Rejected successfully.'}

       else if(Status=='Rework') { 
        confirmation="Do you want to rework this request?";
        resultmessage='Sent for rework.'
       }

       Swal.fire({
        // title: 'Do you want to save?',
        title: confirmation ,
        showConfirmButton: true,
        showCancelButton: true,
        confirmButtonText: "Yes",
        cancelButtonText: "No"
        
      }).then(async (result) => {
        console.log(result)
        if (result.isConfirmed) {

          const postResult = await updateItemApproval(postPayload, sp, props.currentItem.Id);
    
          if (postResult) {
              Swal.fire(resultmessage, '', 'success');
              setTimeout(() => {
        
                // window.location.reload()
    
                location.href=`${siteUrl}/SitePages/MyApprovals.aspx`;
        
              }, 1000);
        
        
            }

        }}) 
       
    
      }

    return(
         
        <div className="card">

                    <div className="card-body">

                      <div className="row">
                        {

                        (props.currentItem.Status == "Pending" || props.currentItem.IsRework) && (<div className="col-lg-12">

                            <div className="mb-0" >

                              <label htmlFor="example-textarea" className="form-label text-dark font-14">Remarks:</label>

                              <textarea style={{height:'80px'}} className="form-control" id="example-textarea" rows={5} name="Remark" value={formData.Remark}

                                onChange={(e) => onChange(e.target.name, e.target.value)}></textarea>

                            </div>

                          </div>)

                        }


                      </div>

                      {


                        (props.currentItem.Status == "Pending" || props.currentItem.IsRework)  && (

                          <div className="row mt-3">

                            <div className="col-12 text-center">

                              {!props.DisableApproval? (<a href="my-approval.html">

                                <button type="button" className="btn btn-success waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Approved')}>

                                  <i className="fe-check-circle me-1"></i> Approve

                                </button>

                              </a>):(<div></div>)}

                              { !props.DisableApproval?(<a href="my-approval.html">

                                <button type="button" className="btn btn-warning waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Rework')}>

                                  <i className="fe-corner-up-left me-1"></i> Rework

                                </button>

                              </a>):(<div></div>)}

                              {!props.DisableApproval?(<a href="my-approval.html">

                                <button type="button" className="btn btn-danger waves-effect waves-light m-1" onClick={(e) => handleFromSubmit(e, 'Rejected')}>

                                  <i className="fe-x-circle me-1"></i> Reject

                                </button>

                              </a>):(<div></div>)}

                              {!props.DisableCancel?(<button type="button" className="btn cancel-btn waves-effect waves-light m-1" onClick={(e) => handleCancel()}>

                                <i className="fe-x me-1"></i> Cancel

                              </button>):(<div></div>)}

                            </div>

                          </div>

                        )

                      }


                    </div>

                  </div>
    )

}