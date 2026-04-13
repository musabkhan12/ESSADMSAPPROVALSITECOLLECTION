import { faHeart, faThumbsUp } from "@fortawesome/free-solid-svg-icons";

import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import React, { useRef } from "react";
import { useState } from "react";
import { Heart, MessageSquare, ThumbsUp, User, MoreVertical } from "react-feather";
import "../../CustomJSComponents/CustomCommentCard/Commentscard.scss";
import "../../CustomCss/mainCustom.scss"
import moment from "moment";
// import {getSP} from '../../webparts/blogDetails/loc/pnpjsConfig';
import {getSP} from '../../webparts/essaApproval/loc/pnpjsConfig';
import { SPFI } from "@pnp/sp/presets/all";
import Swal from "sweetalert2";
import Avatar from "@mui/material/Avatar";
// Define types for reply and comment structures
interface Reply {
  Id: number;
  AuthorId: number,
  UserName: string;
  UserEmail: string;
  SPSPicturePlaceholderState:string;
  Comments: string;
  Created: string;
  UserProfile: string;
}

interface Like {
  ID: number;
  like: string;
  UserName: string;
  AuthorId: number,
  Count: number;
  Created: string;
}

interface Comment {
  Id: number
  UserName: string;
  AuthorId: number,
  Comments: string;
  Created: string;
  UserLikesJSON: Like[];
  UserCommentsJSON: Reply[];
  userHasLiked: boolean; // New property to track if the user liked this comment
  UserProfile: string;
}
// CommentCard component to render individual comments and their replies
export const CommentCard: React.FC<{
  commentId: number;
  username: string;
  Commenttext: string;
  Comments: any;
  Created: string;
  likes: Like[];
  replies: Reply[];
  userHasLiked: boolean;
  CurrentUserProfile: string;
  loadingLike:boolean;
  Action: string;
  onAddReply: (text: string) => void;
  loadingReply:boolean;
  onLike: () => void;
  mainArray:any;
  CurrentUserEmail:string;
  CurrSPSPicturePlaceholderState:string;
  siteUrl:string;
}> = ({ commentId, username, Commenttext, Comments, Created, likes, replies, userHasLiked, CurrentUserProfile,loadingLike, Action, onAddReply, onLike ,loadingReply ,mainArray,CurrentUserEmail,CurrSPSPicturePlaceholderState,siteUrl}) => {
  const sp: SPFI = getSP();
  const [newReply, setNewReply] = useState('');
  const [loading, setLoading] = useState(false); // Loading state for replies
  console.log(Comments, 'Comments');
  console.log(Action, 'Action');
  const menuRef = useRef(null);
  const [openMenuIndex, setOpenMenuIndex] = useState(null);
  const toggleMenu=(index:any)=>{
    setOpenMenuIndex(openMenuIndex === index ? null : index);
  }

  const handleAddReply = async (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
    e.preventDefault()
    if (newReply.trim() === '') return;

    // setLoading(true); // Set loading to true
    await new Promise((resolve) => setTimeout(resolve, 1000)); // Simulate an async operation
    onAddReply(newReply);
    setNewReply('');
    setLoading(false); // Reset loading state
  };
  // const handleKeyDown = (e: any) => {
  //   if (e.key === "Enter") {
  //     e.prevent.default();
  //     console.log("Enter key pressed");
  //     handleAddReply();
  //   }
  // }
  // const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
  //   if (e.key === "Enter") {
  //     e.preventDefault(); // Prevent the default form submission
  //     handleAddReply(e);
  //   }
  // };
   
   React.useEffect(() => {
      const handleClickOutside = (event:any) => {
        if (menuRef.current && !menuRef.current.contains(event.target)) {
          setOpenMenuIndex(null);
        }
      };
      document.addEventListener("mousedown", handleClickOutside);
      return () => {
        document.removeEventListener("mousedown", handleClickOutside);
      };
    }, []);

  const handleReportClick = async (e:any,commentRepliesObject: any,flag:string) => {
      console.log("Report Clicked");
      e.preventDefault()
      try {
           const currentUser = await sp.web.currentUser();
                const reportListName=flag === "replies" ? "ARGBlogUserComments": "ARGBlogComments";
                const eventReportData=await sp.web.lists.getByTitle("ReportedIssueList").items.select("*").filter(`ProcessName eq 'Blog' and ReportedById eq ${currentUser.Id} and ListName eq '${reportListName}' and ListItemId eq ${commentRepliesObject.Id}`)();
                console.log("eventReportData",eventReportData);
                      
                if (eventReportData.length >0 ) {
                    Swal.fire("Already Reported", "You have already reported this content.", "info");
                    return;
                } 
        // Create the popup container
      const popupDiv = document.createElement("div");
      popupDiv.id = "report-issue";
      popupDiv.style.position = "fixed";
      popupDiv.style.top = "50%";
      popupDiv.style.left = "50%";
      popupDiv.style.transform = "translate(-50%, -50%)";
      popupDiv.style.padding = "20px";
      popupDiv.style.backgroundColor = "#fff";
      popupDiv.style.boxShadow = "0px 4px 6px rgba(0,0,0,0.1)";
      popupDiv.style.borderRadius = "8px";
      popupDiv.style.zIndex = "1000";
      popupDiv.style.width = "300px";
 
      // Create a wrapper div inside the popup
      const wrapperDiv = document.createElement("div");
      wrapperDiv.className="report-Issue-Wrapper-Div"
      wrapperDiv.style.padding = "20px";
      wrapperDiv.style.display = "flex";
      wrapperDiv.style.flexDirection = "column";
      wrapperDiv.style.gap = "10px"; // Adds spacing between child elements
      popupDiv.appendChild(wrapperDiv);
   
      // Add a heading
      const heading = document.createElement("h2");
      heading.innerText = "Report Reason";
      heading.style.margin = "0 0 10px 0";
      // popupDiv.appendChild(heading);
      wrapperDiv.appendChild(heading);
   
      // Add a close button
      const closeButton = document.createElement("span");
      closeButton.innerText = "x";
      closeButton.style.position = "absolute";
      closeButton.style.top = "20px";
      closeButton.style.right = "10px";
      closeButton.style.lineHeight = "28px";
      closeButton.style.textAlign = "center";
      closeButton.style.border = "1px solid #ccc";
    closeButton.style.width = "32px";
    closeButton.style.height = "32px";
    closeButton.style.background = "transparent";
    closeButton.style.fontSize = "22px";
    closeButton.style.borderRadius = "1000px";
    closeButton.style.cursor = "pointer";
      closeButton.style.color="#3c3c3c"
      closeButton.onclick = () => {
        document.body.removeChild(popupDiv);
      };
      // popupDiv.appendChild(closeButton);
      wrapperDiv.appendChild(closeButton);
   
      // Add the textarea
      const textAreaElement = document.createElement("textarea");
      textAreaElement.placeholder = "Why are you reporting this comment?";
      textAreaElement.style.width = "100%";
      textAreaElement.style.height = "80px";
      textAreaElement.style.padding = "8px";
      textAreaElement.style.marginBottom = "10px";
      textAreaElement.style.border = "1px solid #ccc";
      textAreaElement.style.borderRadius = "4px";
      // popupDiv.appendChild(textAreaElement);
      wrapperDiv.appendChild(textAreaElement);
   
      // Add a submit button
      const submitButton = document.createElement("button");
      submitButton.innerText = "Submit";
      submitButton.style.padding = "8px 16px";
      submitButton.style.backgroundColor = "#007BFF";
      submitButton.style.color = "#fff";
      submitButton.style.border = "none";
      submitButton.style.borderRadius = "4px";
      submitButton.style.cursor = "pointer";
      submitButton.onclick = async () => {
        const issueValue = textAreaElement.value.trim();
        if (!issueValue) {
          Swal.fire("Error", "Please provide a reason for reporting.", "error");
          return;
        }
        // const currentUser = await sp.web.currentUser();
        console.log("currentUser",currentUser);
        try {
         
          const commentObject = Comments[commentId];
          console.log("flag",flag);
          console.log("commentRepliesObject",commentRepliesObject);
          console.log("mainArray",mainArray);
          let listName="";
          let reportedContentAddedOn;
          let reportedContent;
          let listItemID;
          let reportedContentAddedBy;
          let title;
          let processName;
          if(Action === "Blog"){
            processName ="Blog";
            listName = flag === "replies" ? "ARGBlogUserComments": "ARGBlogComments";
            reportedContentAddedOn = flag === "replies" ? commentRepliesObject.Created : Created;
            reportedContent = flag === "replies" ? commentRepliesObject.Comments: Commenttext;
            listItemID = flag === "replies" ? commentRepliesObject.Id : commentObject.Id;
            reportedContentAddedBy =flag === 'replies' ? commentRepliesObject.AuthorId: commentObject.AuthorId;
            title=mainArray[0].Title;
          }else if(Action === "Announcement"){
            processName="Announcement"
            listName = flag === "replies" ? "ARGAnnouncementAndNewsUserComments": "ARGAnnouncementandNewsComments";
            reportedContentAddedOn = flag === "replies" ? commentRepliesObject.Created : Created;
            reportedContent = flag === "replies" ? commentRepliesObject.Comments: Commenttext;
            listItemID = flag === "replies" ? commentRepliesObject.Id : commentObject.Id;
            reportedContentAddedBy =flag === 'replies' ? commentRepliesObject.AuthorId: commentObject.AuthorId;
            title=mainArray[0].Title;
          }else if(Action === "Group"){
            processName="Group And Teams"
            listName = flag === "replies" ? "ARGGroupandTeamUserComments": "ARGGroupandTeamComments";
            reportedContentAddedOn = flag === "replies" ? commentRepliesObject.Created : Created;
            reportedContent = flag === "replies" ? commentRepliesObject.Comments: Commenttext;
            listItemID = flag === "replies" ? commentRepliesObject.Id : commentObject.Id;
            reportedContentAddedBy =flag === 'replies' ? commentRepliesObject.AuthorId: commentObject.AuthorId;
            title=mainArray[0].GroupName;
          }
          const payload={
            ReportReason: issueValue,
            ProcessName: processName,
            ReportedDate: new Date(),
            Status: "Pending",
            ListName: listName,
            ReportedContentAddedOn: reportedContentAddedOn,
            ReportedContent:reportedContent,
            ReportedById: currentUser.Id,
            ListItemId: listItemID,
            ReportedContentAddedById: reportedContentAddedBy,
            Title: title,
            Action:"Active",
            MainListColumnName:flag === "replies" ? "UserCommentsJSON" :"",
            MainListName:flag === "replies" ? "ARGBlogComments" :"",
            MainListItemId:flag === "replies" ? commentObject.Id:0,
            MainListStatus:flag === "replies" ? "Available":"NA",
          }
          // const insertData = await sp.web.lists.getByTitle("ReportedIssueList").items.add({
          //   ReportReason: issueValue,
          //   ProcessName: "Event",
          //   ReportedDate: new Date(),
          //   Status: "Pending",
          //   ListName: "ARGEventsComments",
          //   ReportedContentAddedOn: Created,
          //   ReportedContent:Commenttext,
          //   ReportedById: currentUser.Id,
          //   ListItemId: commentObject.Id,
          //   ReportedContentAddedById: commentObject.AuthorId,
          //   Title: EventArray[0].EventName,
          //   Action:"Active"
          // });
          const insertData = await sp.web.lists.getByTitle("ReportedIssueList").items.add(payload);
          console.log("payload",payload)
          console.log("Items added successfully");
          document.body.removeChild(popupDiv);
          Swal.fire("Success", "Reported successfully", "success");
        } catch (error) {
          console.log("Error adding the data into the list", error);
        }
      };
      // popupDiv.appendChild(submitButton);
      wrapperDiv.appendChild(submitButton);
   
      // Append the popup to the body
      document.body.appendChild(popupDiv);
      } catch (error) {
        console.log("Error in report popup",error);
      }
      
    };
  return (
    <div className="card team-fedd p-4" style={{  boxShadow: '0 3px 20px #1d26260d' }}>
      <div className="">
        <div className="row">
          <div className="d-flex align-items-start">
            {/* <img
              className="me-2 mt-1 avatar-sm rounded-circle"
              src={Comments[0].UserProfile}
              alt="User"
            /> */}
            {Number(Comments[commentId].SPSPicturePlaceholderState) == 0 ?
              <img
                
                src={

                  `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${Comments[commentId].AuthorEmail}`

                }
                className="me-2 mt-0 avatar-sm rounded-circle"
                alt="User"
              />
              :
              Comments[commentId].AuthorEmail !== null &&Comments[commentId].AuthorEmail !== "" &&
              <Avatar sx={{ bgcolor: 'primary.main' }} className="me-2 mt-0 rounded-circle avatar-xl fontl">
                {`${Comments[commentId].AuthorEmail?.split('.')[0].charAt(0)}${Comments[commentId].AuthorEmail?.split('.')[1].charAt(0)}`.toUpperCase()}
              </Avatar>
            }

            {/* <User /> */}
            <div className="w-100 mt-0">
              <h5 className="mt-0 font-16 fw600 mb-0 text-dark fw-bold">
                {/* <a href="#" className="text-dark fw-bold font-14"> */}
                {username}
                {/* </a> */}
              </h5>
              <p className="text-muted font-12 mt-0">
                {/* <small>{Created}</small> */}
                <small> {moment(Created).format("DD-MMM-YYYY")}</small>
              </p>
            </div>

            <div className="post-content">
                              <div className="post-actions">
                                      <div className="menu-toggle"
                                      onClick={()=>toggleMenu(Comments[commentId].Id)}
                                      >
                                          <MoreVertical size={20} />
                                      </div>
                                      {openMenuIndex === Comments[commentId].Id && (
                                          <div className="dropdown-menucsspost" ref={menuRef}>
                                              <button
                                              // onClick={(e) => handleEditClick(e)} disabled={post.AutherId != CurrentUser.Id}
                                              onClick={(e) => handleReportClick(e,Comments[commentId],"MainComment")}
                                             
                                              >Report</button>
                                          </div>
                                      )}
                            </div>
                        </div>
          </div>

          <p style={{ whiteSpace: 'pre-wrap'}} className="mt-2 font-16">{Commenttext}</p>

          <div className="mt-0 mb-2 d-flex" style={{ gap: '2rem' }}>
            <div  onClick={!loadingLike ? onLike : undefined}  className="btn btn-sm btn-link text-muted hovertext ps-0" style={{
              fontSize: '0.765625rem',display: 'flex', gap:'5px', alignItems:'center',  pointerEvents: loadingLike ? 'none' : 'auto',  opacity: loadingLike ? 0.5 : 1,
              textDecoration: 'unset' 
            }} >
              {
                likes.length > 0 ? <FontAwesomeIcon icon={faThumbsUp} className="text-primary" size='xl' /> : <ThumbsUp size={15} />
              }
              {/* Add heart icon here */}
              {/* <FontAwesomeIcon icon={faHeart} className="text-primary" />  */}
              {` ${likes.length} Likes`}
            </div>
            <div style={{ display: 'flex', fontSize: '0.765625rem',gap:'5px', alignItems:'center'}}>
              <MessageSquare size={15} />  {replies.length} Replies
            </div>
            {/* <button onClick={onLike} className="btn btn-sm btn-link text-muted ps-0">
              <i className="mdi mdi-heart text-primary"></i> {likes.length} Likes
            </button> */}
          </div>
        </div>
        {/* Render replies */}
        <div style={{ display: 'block' }} >
          <div className="row commentheight">
            <div className=" ">
              {replies.map((reply, index) => (
                <div key={index} className="UserReplycss p-2 d-flex " style={{ width: '100%', display: 'flex' }}>
                  <div>
                    {/* <img
                      className="me-2 mt-0 avatar-sm rounded-circle"
                      src={reply.UserProfile}
                      alt="User"
                    /> */}
                    {Number(reply.SPSPicturePlaceholderState) == 0 ?
                      <img

                        src={

                          `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${reply.UserEmail}`

                        }
                        className="me-2 mt-0 avatar-sm rounded-circle"
                        alt="User"
                      />
                      :
                      reply.UserEmail !== null && reply.UserEmail !== undefined &&reply.UserEmail !== "" &&
                      <Avatar sx={{ bgcolor: 'primary.main' }} className="me-2 mt-0 rounded-circle avatar-xl fontl">
                        {`${reply.UserEmail?.split('.')[0].charAt(0)}${reply.UserEmail?.split('.')[1].charAt(0)}`.toUpperCase()}
                      </Avatar>
                    }
                  </div>
                  <div className="w-100 mt-0">
                    <h6 className="font-16 fw600">{reply.UserName}</h6>
                    <p className="mb-2 para-width  text-muted ng-binding" style={{ wordBreak: 'break-all', fontSize:'15px', paddingRight:'22px' }}>{reply.Comments}</p>
                    <p className="text-muted font-14 mt-0 mb-0">
                      <small> {moment(reply.Created).format("DD-MMM-YYYY")}</small>
                    </p>

                  </div>

                  {/* Three-dot menu */}
                            <div className="post-actions" style={{ marginLeft: "auto", position: "relative" }}>
                              <div className="menu-toggle" onClick={() => toggleMenu(reply.Id)}>
                                <MoreVertical size={20} />
                              </div>
                              {openMenuIndex === reply.Id && (
                                <div
                                  className="dropdown-menucsspost"
                                  ref={menuRef}
                                  style={{
                                    position: "absolute",
                                    top: "30px",
                                    right: "0",
                                    background: "white",
                                    boxShadow: "0 2px 8px rgba(0,0,0,0.1)",
                                    borderRadius: "4px",
                                    zIndex: 1000,
                                  }}
                                >
                                  <button
                                    onClick={(e) => handleReportClick(e,reply,"replies")}
                                    style={{
                                      background: "none",
                                      border: "none",
                                      
                                      width: "100%",
                                      textAlign: "left",
                                      cursor: "pointer",
                                      color:"black"
                                    }}
                                  >
                                    Report
                                  </button>
                                </div>
                              )}
                            </div>
                </div>

              ))}
            </div>
          </div>
          <div className="row mt-4" style={{ display: 'block' }}>
            <div className=" align-items-start mt-1">
              <div className="w-100">
                <div className="d-flex align-items-start">
                  <div className="al nice me-2">

                    {/* <img src={CurrentUserProfile} className="w30 avatar-sm rounded-circle" alt="user" /> */}
                    {Number(CurrSPSPicturePlaceholderState) == 0 ?
                      <img

                        src={

                          `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${CurrentUserEmail}`

                        }
                        className="w30 avatar-sm rounded-circle" alt="user"
                      />
                      :
                      CurrentUserEmail !== null && CurrentUserEmail !== "" &&
                      <Avatar sx={{ bgcolor: 'primary.main' }} className="w30 rounded-circle avatar-xl fontl">
                        {`${CurrentUserEmail?.split('.')[0].charAt(0)}${CurrentUserEmail?.split('.')[1].charAt(0)}`.toUpperCase()}
                      </Avatar>
                    }
                  </div>
                  <textarea
                    className="form-control ht form-control-sm"
                    placeholder="Reply to comment..."
                    value={newReply}
                    onChange={(e) => setNewReply(e.target.value)}
                    disabled={loadingReply}
                    rows={3}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && !e.shiftKey) {
                        e.preventDefault(); // Prevents a new line
                        handleAddReply(e); // Calls the function to add a reply
                      }
                    }}
                  />
                  {/* <textarea
                    className="form-control ht form-control-sm"
                    placeholder="Reply to comment..."
                    value={newReply}
                    onChange={(e) => setNewReply(e.target.value)}
                    disabled={loading}
                  /> */}
                </div>

                {/* <button
              className="btn btn-primary mt-2"
              onClick={handleAddReply}
              disabled={loading}
            >
              {loading ? 'Submitting...' : 'Reply'}
            </button> */}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};