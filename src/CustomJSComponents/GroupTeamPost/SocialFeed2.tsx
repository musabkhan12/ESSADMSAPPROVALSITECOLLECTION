import React, { useEffect, useState } from 'react'
import HorizontalNavbar from '../../webparts/horizontalNavBar/components/HorizontalNavBar'
import VerticalSideBar from '../../webparts/verticalSideBar/components/VerticalSideBar'
import { ISocialFeedProps } from '../GroupTeamPost/ISocialFeedProps'
import Provider from '../../GlobalContext/provider'
import { getSP } from '../../webparts/essaApproval/loc/pnpjsConfig'
import { SPFI } from '@pnp/sp/presets/all';
import UserContext from '../../GlobalContext/context';
import {

  addItem,
  additemtoFollowedGroup,

  fetchgrouppandteammaybeInterested,

  fetchNotFollowedGroupdata,
  getGroupTeam,

  getGroupTeamByID,

  getGroupTeamDetailsById,

  getType,

  updateItem,

} from "../../APISearvice/GroupTeamService";
import "../../socialFeed/components/SocialFeed.scss";
import "../../groupandTeamDetails/components/SocialFeed2.scss";
import "../../groupandTeamDetails/components/GroupandTeamDetails.scss";
import "../../../CustomCss/mainCustom.scss";
import "../../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import "bootstrap/dist/css/bootstrap.min.css";
// import CustomBreadcrumb from '../../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb'
import CustomBreadcrumb from '../../CustomJSComponents/CustomBreadcrumb/CustomBreadcrumb'
import { Link, Plus, PlusCircle, Rss, TrendingUp, User, UserPlus, Users } from 'react-feather'
import { getCurrentUser, getCurrentUserName, getCurrentUserNameId, getCurrentUserProfileEmail, getuserprofilepic } from '../../APISearvice/CustomService'
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome'
import { faPaperPlane } from "@fortawesome/free-solid-svg-icons"
import classNames from 'classnames'
import { uploadFileToLibrary } from '../../APISearvice/AnnouncementsService'
import { PostComponent } from '../../CustomJSComponents/SocialFeedPost/PostComponent'
import { fetchBlogdatatop } from '../../APISearvice/BlogService'
import AvtarComponents from '../../CustomJSComponents/AvtarComponents/AvtarComponents'
import { fetchUserInformationList } from '../../APISearvice/Dasborddetails'
import { getDiscussion } from '../../APISearvice/DiscussionForumService'
import moment from 'moment'
import { fetchMediaGallerydata } from '../../APISearvice/MediaDetailsServies'
import { GroupPostComponent } from '../../CustomJSComponents/GroupTeamPost/GroupPostComponent'
import Swal from 'sweetalert2'
import { colors } from '@mui/material'
import Avatar from "@mui/material/Avatar";

export interface IGroupAndTeamPosts {
  Id: any;
  Created: any;
  Comments: string;
  CommentsCount: any;
  GroupTeamCommentsId: any;
  GroupTeamImagesJson: any;
  GroupTeamLikesId: any;
  GroupTeamsImagesId: any;
  GroupandTeamId: any;
  LikesCount: any;
  Title: any;
  UserCommentsJSON: any; UserLikesJSON: any;
  UserName: any;
  UserProfile: any;
  userHasLiked: any;
  GroupTeamComments: { Id: any; Comments: any; };
  GroupTeamsImages: { Id: any; }
  GroupTeamLikes: { Id: any; }
  Author: { Id: any; Title: any;EMail?:any }
}


let userJobTitle: any
let userEmail: any
let userDepartment: any
let userWorkPhone: any
let userOfficeLocation: any
let GroupName: any
let userjobtitle = "";
let GroupDescription: any
interface Post {
  text: string;
  images: string[];
}
const SocialFeedContext = ({ props }: any) => {
  const siteUrl = props.siteUrl;
  const elementRef = React.useRef<HTMLDivElement>(null);
  const sp: SPFI = getSP();
  const { useHide }: any = React.useContext(UserContext);
  const context = React.useContext(UserContext);
  const { setHide }: any = context;
  console.log(sp, "sp");
  const [posts, setPosts] = useState([]);
  const [postsME, setPostsME] = useState([]);
  const [groupsData, setGroupsData] = React.useState([]);
  const [Contentpost, setContent] = useState('');
  const [SocialFeedImagesJson, setImages] = useState<string[]>([]);
  const [hideCreatePost, setHideCreatePost] = useState(true)
  const [HideShowPost, setHideShowPost] = useState(false)
  const [currentEmail, setCurrentEmail] = useState("")
  const [currentId, setCurrentID] = useState<any>(0)
  const [filterCriteria, setFilterCriteria] = useState<string | null>(null);
  const [currentUsername, setCurrentUserName] = useState("")
   const [SPSPicturePlaceholderState, setSPSPicturePlaceholderState]= useState(null);
  const [currentUserId, setCurrentUserId] = useState(0)
  const [UploadedFile, setUploadFile] = useState([])
  const [ImageIds, setImageIds] = useState([])
  const [blogdata, setblogdata] = useState([])
  const [groupmembers, setGroupMembersArr] = useState<any[]>([]);
  const [usersitem, setUsersArr] = useState<any[]>([]);
  const [DiscussionData, setDiscussion] = useState([])
  const [isSubmitting, setIsSubmitting] = useState(false);
  const [copySuccess, setCopySuccess] = useState('');
  const [activeMainTab, setActiveMainTab] = useState("feed");
  const [ArrDetails, setArrDetails] = useState([]);
  const [getAllgroup, setgetAllgroup] = useState([]);
  const [activeTab, setActiveTab] = useState("allgroups");
  const [gallerydata, setGalleryData] = useState<any[]>([]);
  const [Loading1, setLoading1] = useState(false);
  const [IsEdit, setIsEdit] = useState<boolean>(false);
  // const [formData, setFormData] = React.useState({
  //   Post: "",

  // });
  const [argcurrentgroupuser, setArgcurrentgroupuser] = useState([])
  useEffect(() => {
    // Load posts from localStorage when the component mounts
    // const storedPosts = JSON.parse(localStorage.getItem('posts') || '[]');
    // console.log(storedPosts);
    fetchDetailbyProjectName()
    // setPosts(storedPosts);
    getAllAPI()
    // getuserprofile()

  }, [props]);
  async function fetchDetailbyProjectName() {
    console.log(GroupName, "GroupName")
    const ids = window.location.search;
    const originalString = ids;
    const idNum = originalString.substring(1);
    try {


      // alert(GroupName)
      const getargmember = await sp.web.lists.getByTitle('ARGGroupandTeam').items.filter(`GroupName eq '${GroupName}'`).select("*,InviteMemebers/ID,InviteMemebers/Title ,Author/ID,Author/Title").expand("InviteMemebers ,Author")();
      console.log(getargmember, "getargmember")

      setArgcurrentgroupuser(getargmember)
    } catch (error) {

    }


  }

  console.log(argcurrentgroupuser, "ARGGroupandTeam member")
  useEffect(() => {
    getuserprofile()
  }, []);
  const getownerjobtitle = async (mail: String) => {
    debugger

    let userdetails = await sp.web.lists.getByTitle("User Information List").items
      .select("ID", "Title", "EMail", "Department", "JobTitle", "Picture", "MobilePhone", "WorkPhone")
      .filter(`EMail eq '${mail}'`)
      ().then((x) => {
        if (x.length > 0) {
          userjobtitle = x[0].JobTitle !== null ? x[0].JobTitle : "NA"
        }
        console.log("testytfghv", x, userJobTitle)
      })
    return userjobtitle
  }
  const getuserprofile = async () => {
    //alert("hi")
    try {
      // Fetch the current user's profile properties
      const userProfile = await sp.profiles.myProperties();

      // Display a few selected properties
      const userDisplayName = userProfile.DisplayName;
      userEmail = userProfile.Email;
      userJobTitle = userProfile.Title; // or use 'JobTitle' if available
      userDepartment = userProfile.Department;
      userWorkPhone = userProfile.WorkPhone;
      userOfficeLocation = userProfile.Office;
      console.log("User Name:", userDisplayName);
      console.log("User Email:", userEmail);
      console.log("User Job Title:", userJobTitle);
      console.log("userDepartment ", userDepartment);
      console.log("userWorkPhone:", userWorkPhone);
      console.log("userOfficeLocation:", userOfficeLocation);
      setGroupsData(await getGroupTeam(sp));
      // Return data if needed for component display
      return {
        name: userDisplayName,
        email: userEmail,
        jobTitle: userJobTitle,
      };

    } catch (error) {
      console.error("Error fetching user profile:", error);
    }

  }




  useEffect(() => {
    // alert
    getGroup()
  }, []);

  const filteredPosts = posts.filter((post) => {
    if (!filterCriteria) return true; // Show all posts if no filter is applied
    return post.Contentpost.includes(filterCriteria); // Adjust based on filter logic
  });
  let resfiltergroup: any = [];
  const getGroup = async () => {
    debugger
    const userList = await sp.web.lists.getByTitle("User Information List").items.select("ID", "Title", "EMail", "Department", "JobTitle", "Picture", "MobilePhone", "WorkPhone")();
    setGroupMembersArr(userList);
    const ids = window.location.search;
    //  alert(ids)
    const originalString = ids;
    // alert(originalString)
    const idNum2: any = originalString.substring(1);
    // alert(idNum2)
    debugger;
    const getgroup1 = await sp.web.lists
      .getByTitle("ARGGroupandTeam")
      .items.getById(idNum2).select("*,GroupFollowers/Id,GroupFollowers/Title,InviteMemebers/Id,InviteMemebers/Title,GroupType , Author/ID,Author/Title").expand("InviteMemebers , GroupFollowers , Author")()
      //.then(async (res) => {
        // arr=res;
        GroupName = getgroup1.GroupName
        GroupDescription = getgroup1.GroupDescription
        console.log(getgroup1, ":response")
        // debugger
        
        let arrr: any[] = [];
        if(getgroup1?.GroupFollowersId !== null){
          for (var j = 0; j < getgroup1.GroupFollowers.length; j++) {
            //console.log("hjhj",getgroup1.GroupFollowers[j].ID)
            var user = await sp.web.getUserById(getgroup1.GroupFollowersId[j])();
            console.log("userrrrrrr",user,getgroup1.GroupFollowersId[j])
            if (user) {
            var profile = await sp.profiles.getPropertiesFor(`i:0#.f|membership|${user.Email}`);
            getgroup1.GroupFollowers[j].EMail = user.Email;
        
            getgroup1.GroupFollowers[j].SPSPicturePlaceholderState = profile.UserProfileProperties?profile.UserProfileProperties[profile.UserProfileProperties.findIndex((obj: { Key: string })=>obj.Key === "SPS-PicturePlaceholderState")].Value:"1";
            }  
          }
        }
       
        
        arrr.push(getgroup1)
        setArrDetails(arrr)
        debugger

        userList.forEach(function (element, index, array) {
          // Code to execute for each element
          if (element.Id == getgroup1.GroupFollowersId) {
            resfiltergroup.push(element);
          }
          setGroupMembersArr(resfiltergroup);
        });
        // const resgroup=groupmembers.filter(x=>x.Id.includes(res.GroupFollowersId));

      //})
      // .catch((error) => {
      //   console.log("Error fetching data: ", error);
      // });
   
  }

  console.log(ArrDetails, "ArrDetails data success")
  console.log(ArrDetails, "ArrDetails data success")
  const getAllAPI = async () => {
    const ids = window.location.search;
    const originalString = ids;
    const idNum = originalString.substring(1);
    const currentUser = await getCurrentUserNameId(sp);
    let getAllgroups = await fetchgrouppandteammaybeInterested(sp);
    console.log("getAllgroup------", getAllgroups)
    setgetAllgroup(getAllgroups)
    setCurrentUserId(currentUser);
    let currentGroup = await getGroupTeamDetailsById(sp, Number(idNum));
    userjobtitle = await getownerjobtitle(currentGroup[0]?.Author?.EMail);
    setArrDetails(currentGroup);
    console.log(currentGroup, "currentGroup", userjobtitle);
    if (currentGroup[0]?.GroupType === "Selected Members" &&
      currentGroup[0]?.InviteMemebers?.some((invitee: any) => invitee.Id === currentUser) || currentGroup[0].Author.ID === currentUser) {
      setIsEdit(true);
    } else if (currentGroup[0]?.GroupType === "All" && currentGroup[0]?.GroupFollowers &&
      currentGroup[0]?.GroupFollowers?.some((invitee: any) => invitee.Id === currentUser) || currentGroup[0].Author.ID === currentUser) {
      setIsEdit(true);
    }
    else if (currentGroup[0]?.GroupType === "All" && currentGroup[0]?.GroupFollowers.length == 0 &&
      currentGroup[0].Author.ID === currentUser) {
      setIsEdit(true);
    }

    const galleryItemsone = await fetchMediaGallerydata(sp);
    setGalleryData(galleryItemsone);
    setCurrentEmail(await getCurrentUserProfileEmail(sp))


    setCurrentUserName(await getCurrentUserName(sp))
    setblogdata(await fetchBlogdatatop(sp))
    setUsersArr(await fetchUserInformationList(sp))
    setDiscussion(await getDiscussion(sp))
     const profileemail = await getCurrentUserProfileEmail(sp);
    setSPSPicturePlaceholderState( await getuserprofilepic(sp,profileemail));
    fetchPosts();


  }

  // const handleImageChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
  //   const files = Array.from(e.target.files || []); 
  //   let uploadedImages: any[] = [];
  //   let ImagesIds: any[] = [];
  //   for (const file of files) {
  //     try {

  //       const uploadedImage = await uploadFileToLibrary(file, sp, "SocialFeedImages");

  //       uploadedImages.push(uploadedImage); 
  //       console.log(uploadedImage, 'Uploaded file data');
  //     } catch (error) {
  //       console.log("Error uploading file:", file.name, error);
  //     }
  //   }





  //   setImages(flatArray(uploadedImages)); 

  //   setUploadFile(flatArray(uploadedImages)); 

  // };


  const handleImageChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(e.target.files || []); // Ensure files is an array of type File[]
    let uploadedImages: any[] = [];
    let ImagesIds: any[] = [];
    for (const file of files) {
      setLoading1(true);
      try {
        // Assuming uploadFileToLibrary is your custom function to upload files
        const uploadedImage = await uploadFileToLibrary(
          file,
          sp,
          "SocialFeedImages"
        );

        uploadedImages.push(uploadedImage); // Store uploaded image data
        console.log(uploadedImage, "Uploaded file data");
      } catch (error) {
        setLoading1(false);
        console.log("Error uploading file:", file.name, error);
      } finally {
        setLoading1(false); // Enable the button after the function completes
      }
    }
    // Set state after uploading all images

    setImages(flatArray(uploadedImages)); // Store all uploaded images

    setUploadFile(flatArray(uploadedImages)); // Optional: Track the uploaded file(s) in another state
  };

  //#region flatArray

  const flatArray = (arr: any[]): any[] => {

    return arr.reduce((acc, val) => acc.concat(val), []);

  };

  //#endregion

  console.log(SocialFeedImagesJson, 'SocialFeedImagesJson');
  const validateForm = () => {

    let valid = true;

    if (!Contentpost) {
      Swal.fire('Error', 'Post is required!', 'error');
      valid = false;
    }
    return valid;
  };

  const [errorMessage, setErrorMessage] = useState('');
  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    //if (validateForm()) {
    if (!Contentpost) {
      event.preventDefault()
      setErrorMessage('Please write something in the post.');
      return;
    }

    setErrorMessage('');
    let ImagesIdss: any[] = [];
    ImagesIdss = ImagesIdss.concat(SocialFeedImagesJson.map((item: any) => item.ID));
    setImageIds(ImagesIdss)
    e.preventDefault();
    let newPostss: any
    const ids = window.location.search;
    const originalString = ids;
    const idNum = originalString.substring(1);

    await sp.web.lists.getByTitle("ARGGroupandTeamComments")
      .items.add({
        Comments: Contentpost,
        GroupTeamImagesJson: JSON.stringify(SocialFeedImagesJson),
        UserProfile: `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentEmail}`,
        LikesCount: 0,
        CommentsCount: 0,
        GroupandTeamId: idNum,
        GroupTeamsImagesId: ImagesIdss
      });
    //useState('');
    getAllAPI()
    setContent('');
    setImages([]);
    fetchPosts();
    //}
  };



  const ShowPost = () => {

    //debugger

    setHideCreatePost(false)

    setHideShowPost(true)
    fetchPostsMe();

  }

  const CreatePostTab = () => {

    //debugger

    setHideCreatePost(true)

    setHideShowPost(false)
    fetchPosts()
  }

  const [post, setPost] = useState("");

  const [storedPosts, setStoredPosts] = useState([]);



  // Fetch posts from SharePoint when the component loads

  // Fetch posts from SharePoint list

  const fetchPosts = async () => {
    debugger
    const ids = window.location.search;
    const originalString = ids;
    const idNum = originalString.substring(1);


    try {
      sp.web.lists.getByTitle("ARGGroupandTeamComments")
        .items.select("*,GroupTeamComments/Id,GroupTeamComments/Comments,GroupTeamsImages/Id,GroupTeamLikes/Id,Author/Id,Author/Title,Author/EMail,Author/SPSPicturePlaceholderState")
        .expand("GroupTeamComments,GroupTeamsImages,GroupTeamLikes,Author")
        .orderBy("Created", false)
        .filter(`GroupandTeamId eq ${idNum}`)().then((results: IGroupAndTeamPosts[]) => {
          console.log("currentpost",results);
          if (results) {
            debugger;
            const PostItems = results.map((ele: IGroupAndTeamPosts) => {
              return {
                Contentpost: ele.Comments,
                GroupTeamImagesJson: ele.GroupTeamImagesJson,
                Created: ele.Created,
                userName: ele.Author?.Title,
                Author: ele.Author,
                userAvatar: ele.UserProfile,
                likecount: ele.LikesCount,
                commentcount: ele.CommentsCount,
                Id: ele.Id,
                SocialFeedUserLikesJson: ele.UserLikesJSON ? JSON.parse(ele.UserLikesJSON) : [],
                gcomments: ele.UserCommentsJSON ? JSON.parse(ele.UserCommentsJSON) : [],
              }
            });
            console.log("setpostss",PostItems)
            setPosts(PostItems);

          }
        });
    } catch (error) {
      console.log("Error fetching posts:", error);
    }
  };

  const fetchPostsMe = async () => {
    const ids = window.location.search;
    const originalString = ids;
    const idNum = originalString.substring(1);

    try {
      const cuurentID = await getCurrentUserNameId(sp);
      await sp.web.lists.getByTitle("ARGGroupandTeamComments")
        .items.select("*,GroupTeamComments/Id,GroupTeamComments/Comments,GroupTeamsImages/Id,GroupTeamLikes/Id,Author/Id,Author/Title,Author/EMail,Author/SPSPicturePlaceholderState")
        .expand("GroupTeamComments,GroupTeamsImages,GroupTeamLikes,Author")
        .orderBy("Created", false)
        .filter(`GroupandTeamId eq ${idNum} and AuthorId eq ${cuurentID}`)().then((results: any) => {
          console.log("me result",results);
          if (results) {
            const PostItems = results.map((ele: IGroupAndTeamPosts) => {
              return {
                Contentpost: ele.Comments,
                GroupTeamImagesJson: ele.GroupTeamImagesJson,
                Created: ele.Created,
                userName: ele.Author?.Title,
                Author:ele.Author,
                userAvatar: ele.UserProfile,
                likecount: ele.LikesCount,
                commentcount: ele.CommentsCount,
                Id: ele.Id,
                SocialFeedUserLikesJson: ele.UserLikesJSON ? JSON.parse(ele.UserLikesJSON) : [],
                gcomments: ele.UserCommentsJSON ? JSON.parse(ele.UserCommentsJSON) : []
              }
            });
            console.log("setpostss me", PostItems)
            setPostsME(PostItems);
          }
        })
    } catch (error) {
      console.log("Error fetching posts:", error);
    }
  };


  console.log(postsME, 'postsME');


  // Handle input change

  const handlePostChange = (e: { target: { value: React.SetStateAction<string> } }) => {

    setPost(e.target.value);

  };



  // Handle form submission and save to SharePoint list

  const handlePostSubmit = async () => {

    const newPost = post.trim();

    if (newPost) {

      try {
        // Add new post to SharePoint list
        debugger;
        await sp.web.lists
          .getByTitle("Posts") // SharePoint list name
          .items.add({
            Title: newPost, // Adding the post content to the "Title" field
          });
        // Refresh the list of posts after adding
        fetchPosts();
        setPost(""); // Clear the input field
      } catch (error) {
        console.log("Error adding post:", error);
      }
    }
  };

  const Breadcrumb = [

    {

      "MainComponent": "Home",

      "MainComponentURl": `${siteUrl}/SitePages/Dashboard.aspx`

    },

    {

      "ChildComponent": "Groups",

      "ChildComponentURl": `${siteUrl}/SitePages/GroupandTeam.aspx`

    }

  ]



  const copyToClipboard = (Id: number) => {

    const link = `${siteUrl}/SitePages/SocialFeed.aspx?${Id}`;

    navigator.clipboard.writeText(link)

      .then(() => {

        setCopySuccess('Link copied!');

        setTimeout(() => setCopySuccess(''), 2000); // Clear message after 2 seconds

      })

      .catch(err => {

        setCopySuccess('Failed to copy link');

      });

  };



  const mergeAndRemoveDuplicates = (str: string, str1: string) => {

    //debugger;

    let url = str1;



    // Find the position of the third occurrence of "/"



    let thirdSlashIndex = url.indexOf(

      "/",

      url.indexOf("/", url.indexOf("/") + 1) + 1

    );



    // Get the substring after the third occurrence of "/"



    let updatedUrl = url.substring(thirdSlashIndex);



    console.log("check the url--->>", updatedUrl); // Output: /SocialFeedImages/announcement-5.jpg





    return str + updatedUrl; // Concatenate directly if str1 starts with a slash

    //}

  };





  const GotoNextPageone = (item: any, pagename: string) => {

    console.log("item-->>>>", item)



    window.location.href = `${siteUrl}/SitePages/${pagename}.aspx`;

  };



  const handleTabClick = (tab: React.SetStateAction<string>) => {

    setActiveTab(tab);

  };



  const truncateText = (text: string, maxLength: number) => {

    if (text != null) {

      return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;



    }

  };

  const addFollowedGroup = async (item: any) => {
    const currentUser = await sp.web.currentUser();
    const Itemdata = {

      GroupIDId: Number(item.ID),
      FollowedById: currentUser.Id
    };
    console.log("Itemdata", Itemdata);

    const postResult = await additemtoFollowedGroup(sp, Itemdata);
    const postId = postResult?.data?.ID;
    Swal.fire("Group followed successfully", "", "success");
    setgetAllgroup(await fetchgrouppandteammaybeInterested(sp));
    debugger
    if (!postId) {
      console.error("Post creation failed.");
      return;
    }
  }


  return (

    <div id="wrapper" ref={elementRef}>

      <div

        className="app-menu"

        id="myHeader">

        <VerticalSideBar _context={sp} />

      </div>

      <div className="content-page">

        <HorizontalNavbar _context={sp} siteUrl={siteUrl} />

        <div className="content" style={{ marginLeft: `${!useHide ? '240px' : '80px'}`, marginTop: '1rem' }}>

          <div className="container-fluid  paddb">

            <div className="row" style={{paddingLeft:'0.5rem'}}>

              <div className="col-lg-3">

                 <CustomBreadcrumb Breadcrumb={Breadcrumb} _context={sp}/>

              </div>

            </div>




            {activeTab === "allgroups" && (<div className="row mt-3">

              <div className="col-md-3">
                <div className="row">

                  <div style={{ gap: '1rem' }}>

                    {/* <img src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentEmail}`}

                      className="rounded-circlecss img-thumbnail avatar-xl" style={{

                        borderRadius: '5rem', height: '3rem',

                        width: '3rem'

                      }} /> */}

                    <span style={{
                      fontSize: '28px',
                      color: 'black',
                      whiteSpace: 'nowrap',

                      display: 'flex',
                      fontWeight: '700',
                      textWrap: 'inherit',
                      alignItems: 'center'

                    }}>{GroupName}</span>
                    <br />
                    <span style={{
                      fontSize: '15px',
                      color: 'black',


                      display: 'flex',

                      alignItems: 'center',
                      lineHeight: '22px',
                      textWrap: 'inherit'

                    }}>{GroupDescription}</span>
                    {/* <div className='font-14 text-muted mt-2'>{GroupDescription}</div> */}

                    <br />
                    <div className="row mt-0">
                      <div >
                        {(ArrDetails[0]?.InviteMemebers?.length > 0 || ArrDetails[0]?.GroupFollowers?.length > 0) &&
                          <div  className="card" style={{ borderRadius: "1rem",position: 'sticky', top: '90px' }}>
                            <div className="card-body pb-0 gheight">

                              <h4 className="header-title font-16 text-dark fw-bold mb-0">
                                Group Members

                              </h4>

                              {/* {console.log("ArrDetails[0]ArrDetails[0]", ArrDetails[0])} */}
                              {ArrDetails[0]?.GroupType === "All" &&
                                ArrDetails[0]?.GroupFollowers?.length > 0 && ArrDetails[0].GroupFollowers.map((follower: any, idx: any) => (

                                  <div className="projectmemeber" key={idx}>
                                    <div className="itemalign">
                                      {/* <img
                                        src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${follower?.EMail}`}
                                        className="rounded-circlenu img-thumbnail avatar-xl"
                                        alt="profile-image"
                                      /> */}

                                      {follower?.SPSPicturePlaceholderState == 0 ?
                                        <img
                                          src={

                                            `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${follower?.EMail}`

                                          }
                                          className="rounded-circle"
                                          width="50"
                                          alt={follower?.Title}
                                        />
                                        :
                                        follower?.EMail !== null &&
                                        <Avatar sx={{ bgcolor: 'primary.main' }} className="rounded-circle font-14 avatar-xl">
                                          {`${follower?.EMail?.split('.')[0].charAt(0)}${follower?.EMail?.split('.')[1].charAt(0)}`.toUpperCase()}
                                        </Avatar>
                                      }
                                      <p>
                                        {follower?.Title}
                                        <img
                                          src={require("../assets/calling.png")}
                                          className="alignright"
                                          onClick={() => window.open(`https://teams.microsoft.com/l/call/0/0?users=${follower.EMail}`, "_blank")}
                                          alt="Call"
                                        />
                                      </p>
                                    </div></div>
                                ))
                              }

                              {/* <p>{GroupName}</p> */}
                              {/* <>{ArrDetails[0].GroupType}</> */}
                              {ArrDetails[0]?.InviteMemebers?.length > 0 &&
                                ArrDetails[0]?.InviteMemebers.map((member: any, idx: any) => (
                                  <div className="projectmemeber">
                                    <div className="itemalign">
                                      {/* <img
                                        
                                        src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${member?.EMail}`}
                                        className="rounded-circlenu img-thumbnail avatar-xl"
                                        alt="profile-image"
                                      /> */}
                                      {member?.SPSPicturePlaceholderState == 0 ?
                                        <img
                                          src={

                                            `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${member?.EMail}`

                                          }
                                          className="rounded-circle"
                                          width="50"
                                          alt={member?.Title}
                                        />
                                        :
                                        member?.EMail !== null &&
                                        <Avatar sx={{ bgcolor: 'primary.main' }} className="rounded-circle font-14 avatar-xl">
                                          {`${member?.EMail?.split('.')[0].charAt(0)}${member?.EMail?.split('.')[1].charAt(0)}`.toUpperCase()}
                                        </Avatar>
                                      }
                                      <p className='mb-0'>{member?.Title} </p>
                                    </div>
                                    {/* {item.Author.EMail === currentUserEmailRef.current && (
        <div>
        <button onClick={()=>handleRemoveUser(id?.ID)}>Remove</button>
        </div>
 
        )
       
        } */}

                                    <img

                                      src={require("../assets/calling.png")}

                                      className="alignright"

                                      onClick={() =>

                                        window.open(

                                          `https://teams.microsoft.com/l/call/0/0?users=${member.EMail}`,

                                          "_blank"

                                        )

                                      }

                                      alt="Call"

                                    />
                                  </div>

                                ))}
                              {/* <div className="inbox-widget" style={{ marginTop: '1rem' }}>
                              {groupmembers.map((user, index) => (
                                <div
                                  key={index}
                                  className="d-flex border-bottom heit8 align-items-start w-100 justify-content-between mb-1"
                                >
                                  <div style={{ display: 'flex' }}>
                                    <a>
                                      <img
                                        src={user.Picture != null ? `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${user.EMail}` : require("../assets/users.jpg")}
                                        className="rounded-circlenu"
                                        width="50"
                                        alt={user.name}
                                      />
                                    </a>
                                    <a href="contacts-profile.html" style={{ marginTop: '10px', marginLeft: '15px' }}>
                                      <p className="fw-bold font-14 mb-0 text-dark">
                                        {user.Title} | {user.Department != null ? user.Department : 'NA'}
                                      </p>
                                    </a>
                                  </div>
 
 
                                </div>
                              ))}
                            </div> */}
                            </div>
                          </div>
                        }

                      </div>


                    </div>

                  </div>

                </div>


              </div>


              <div className="col-md-6 mobile-w2">

                {activeMainTab === "feed" && (

                  <>

                    <div className="card">

                      <div className="post-form">

                        <ul className="navcss nav-tabs nav-bordered">

                          <li className="nav-itemcss">

                            <a className={classNames('nav-linkcss px-4 py-3', { active: hideCreatePost && !HideShowPost })} onClick={CreatePostTab}>Create Post</a>

                          </li>

                          <li className="nav-itemcss">

                            <a className={classNames('nav-linkcss px-4 py-3', { active: HideShowPost && !hideCreatePost })} onClick={ShowPost}>My Post</a>

                          </li>

                        </ul>
                        {console.log("IsEditIsEdit", IsEdit)}
                        {hideCreatePost &&

                          <div className="tab-content pt-0">

                            <div className="tab-pane p-4 active show">

                              <div className="border rounded">

                                <form onSubmit={handleSubmit} className="comment-area-box">

                                  <textarea

                                    className="form-control border-0 resize-none textareacss"

                                    placeholder="Write something in this group..."

                                    value={Contentpost}

                                    rows={4}

                                    onChange={(e) => setContent(e.target.value)}
                                    disabled={!IsEdit}

                                  />

                                  <div className="p-2 bg-light d-flex justify-content-between align-items-center">



                                    <label>

                                      <div>
                                        {IsEdit &&
                                        <div className='btn btn-default'>
                                        <svg  onClick={() => handleImageChange}  xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect><circle cx="8.5" cy="8.5" r="1.5"></circle><polyline points="21 15 16 10 5 21"></polyline></svg>
                                          </div>
                                          // <Link style={{ width: "20px", height: "16px" }} onClick={() => handleImageChange} />
                                        }
                                        <input

                                          type="file"

                                          multiple

                                          accept="image/*"

                                          onChange={handleImageChange}

                                          className="fs-6 w-50" aria-rowspan={5} style={{ display: 'none' }}

                                        />

                                      </div>

                                    </label>

                                    {/* <div className="image-preview mt-2">



                                      {SocialFeedImagesJson.map((image: any, index) => {

                                        var imageUrl = mergeAndRemoveDuplicates(siteUrl, image.fileUrl)


                                        return (
                                          image.fileType === "image/jpeg" ?

                                            <img key={index} src={imageUrl} alt={`preview-${index}`} style={{ width: '100px', marginRight: '10px', marginLeft:'10px' }} />
                                            : 

                                            <a id={"fileLink"+index} href={imageUrl} target="_blank" style={{marginLeft:'10px'}}>{image.fileName}</a>

                                            

                                          )
                                      }





                                      )}

                                    </div> */}



                                    <div className="image-preview mt-2">
                                      {Loading1 ? (
                                        <div
                                          className="spinner-border text-primary"
                                          role="status"
                                        >
                                          <span className="sr-only">
                                            Loading...
                                          </span>
                                        </div>
                                      ) : (
                                        SocialFeedImagesJson.map(
                                          (image: any, index) => {
                                            var imageUrl =
                                              mergeAndRemoveDuplicates(
                                                siteUrl,
                                                image.fileUrl
                                              );
                                            console.log(imageUrl);

                                            return (
                                              <img
                                                key={index}
                                                src={imageUrl}
                                                alt={`preview-${index}`}
                                                style={{
                                                  width: "100px",
                                                  marginRight: "10px",
                                                }}
                                              />
                                            );
                                          }
                                        )
                                      )}
                                    </div>
                                    <p style={{ color: 'red', fontSize: '14px', marginTop: '10px' }}>{errorMessage}</p>
                                    <button type="submit" className="btn btn-sm btn-priamry primary1  font-121" disabled={!IsEdit}>

                                    <FontAwesomeIcon style={{float:'left',margin:"7px 6px 0px 0px"}} icon={faPaperPlane} /> Post

                                    </button>

                                  </div>

                                </form>

                              </div>

                            </div>

                          </div>

                        }

                      </div>

                    </div>
                    {console.log("posts.length", posts.length, hideCreatePost, HideShowPost)}
                    {posts.length > 0 && hideCreatePost && !HideShowPost && (
                      <div className="feed">
                        {posts.length > 0 ? (
                          posts.map((post, index) => (
                            <GroupPostComponent
                              key={index}
                              sp={sp}
                              siteUrl={siteUrl}
                              isedit={IsEdit}
                              currentUserName={currentUsername}
                              currentUser={currentUserId}
                              currentEmail={currentEmail}
                              SPSPicturePlaceholderState ={SPSPicturePlaceholderState}
                              editload={fetchPosts}
                              post={{
                                userName: post.userName,
                                Created: post.Created,
                                Contentpost: post.Contentpost,
                                AuthorDetails:post.Author,
                                Author: post.Author.EMail,
                                AuthorId:post.Author.Id,
                                SocialFeedImagesJson: post.GroupTeamImagesJson,
                                userAvatar: post.userAvatar,
                                likecount: post.likecount,
                                commentcount: post.commentcount,
                                gcomments: post?.gcomments != null ? post.gcomments : [],
                                postId: post.Id,
                                SocialFeedUserLikesJson: post.SocialFeedUserLikesJson,
                              }}
                            groupsArray={ArrDetails}
                            />
                          ))
                        ) : (
                          <div className="no-posts-message">No posts found for the selected filter.</div>
                        )}
                      </div>
                    )}

                    {postsME.length > 0 && !hideCreatePost && HideShowPost &&
                      <div className="feed">
                        {postsME.map((post, index) => {
                          console.log(postsME, 'postsME');
                          return (
                            <GroupPostComponent
                              key={index}
                              sp={sp}
                              siteUrl={siteUrl}
                              isedit={IsEdit}
                              currentUserName={currentUsername}
                              currentUser={currentUserId}
                              currentEmail={currentEmail}
                              SPSPicturePlaceholderState ={SPSPicturePlaceholderState}
                              editload={fetchPosts}
                              post={{
                                userName: post.userName,
                                AuthorDetails:post.Author,
                                Author:post.Author.EMail,
                                AuthorId:post.Author.Id,
                                Created: post.Created,
                                Contentpost: post.Contentpost,
                                SocialFeedImagesJson: post.GroupTeamImagesJson,
                                userAvatar: post.userAvatar,
                                likecount: post.likecount,
                                commentcount: post.commentcount,
                                gcomments: post?.gcomments != null ? post.gcomments : [],
                                postId: post.Id,
                                SocialFeedUserLikesJson: post.SocialFeedUserLikesJson
                              }}
                              groupsArray={ArrDetails}
                            />
                          )
                        }
                        )}
                      </div>

                    }

                  </>

                )}

                {activeMainTab === "followers" && (

                  <div className="row card-view">

                    {usersitem.map((item) => (

                      <div className="col-lg-6 col-md-6" key={item.Title}>

                        <div

                          style={{ border: "1px solid #54ade0" }}

                          className="text-center card mb-3"

                        >

                          <div className="card-body">


                            <div className="pt-2 pb-2">

                              <a style={{ position: "relative" }}>

                                <img

                                  src={require("../assets/calling.png")}

                                  className="alignright"

                                  onClick={() =>

                                    window.open(

                                      "https://teams.microsoft.com",

                                      "_blank"

                                    )

                                  }

                                  alt="Call"

                                />

                                {/* <img  src={

                                    item.Picture != null

                                      ? `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${item.EMail}`

                                      : require("../assets/users.jpg")

                                  }

                                  className="rounded-circlecss img-thumbnail

                                 avatar-xl"

                                  alt="profile-image"

                                  style={{ cursor: "pointer" }}

                                /> */}

                                 {item?.SPSPicturePlaceholderState == 0 ?
                                        <img
                                          src={

                                            `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${item?.EMail}`

                                          }
                                          className="rounded-circle"
                                          width="50"
                                          alt={item?.Title}
                                        />
                                        :
                                        item?.EMail !== null &&
                                        <Avatar sx={{ bgcolor: 'primary.main' }} className="rounded-circle avatar-xl">
                                          {`${item?.EMail?.split('.')[0].charAt(0)}${item?.EMail?.split('.')[1].charAt(0)}`.toUpperCase()}
                                        </Avatar>
                                      }

                              </a>

                              <h4 className="mt-2 mb-1">

                                <a

                                  className="text-dark font-16 fw-bold"

                                  style={{

                                    textDecoration: "unset",

                                    fontSize: "20px",

                                  }}

                                >

                                  <strong>

                                    {truncateText(item.Title, 15)}

                                  </strong>

                                </a>

                              </h4>



                              <p

                                className="text-muted"

                                style={{ fontSize: "14px" }}

                              >

                                <span data-tooltip={item.EMail}>

                                  {truncateText(item.EMail, 15)} |

                                </span>

                                <span

                                  className="pl-2"

                                  style={{ color: "pink" }}

                                >



                                  {truncateText(

                                    item.Department != null

                                      ? item.Department

                                      : " NA ",

                                    10

                                  )}




                                </span>

                              </p>

                              <div

                                style={{

                                  display: "flex",

                                  justifyContent: "center",

                                  gap: "0.5rem",

                                }}

                              >

                                <button

                                  type="button"

                                  onClick={() =>

                                    window.open(

                                      "https://teams.microsoft.com",

                                      "_blank"

                                    )

                                  }

                                  className="btn btn-primary btn-sm waves-effect waves-light"

                                >

                                  Message

                                </button>

                                <button

                                  type="button"

                                  className="btn btn-light btn-sm waves-effect"

                                >

                                  Follow

                                </button>

                              </div>

                              <div className="row mt-2">

                                <div className="col-4">

                                  <div className="mt-3">

                                    <h4

                                      className="fw-bold font-14"

                                      style={{

                                        fontSize: "0.80rem",

                                        color: "#343a40",

                                      }}

                                    >

                                      NA

                                    </h4>

                                    <p className="mb-0 text-muted text-truncate">

                                      Post

                                    </p>

                                  </div>

                                </div>

                                <div className="col-4">

                                  <div className="mt-3">

                                    <h4

                                      className="fw-bold font-14"

                                      style={{

                                        fontSize: "0.80rem",

                                        color: "#343a40",

                                      }}

                                    >

                                      NA

                                    </h4>

                                    <p className="mb-0 text-muted text-truncate">

                                      Followers

                                    </p>

                                  </div>

                                </div>

                                <div className="col-4">

                                  <div className="mt-3">

                                    <h4

                                      className="fw-bold font-14"

                                      style={{

                                        fontSize: "0.80rem",

                                        color: "#343a40",

                                      }}

                                    >

                                      NA

                                    </h4>

                                    <p className="mb-0 text-muted text-truncate">

                                      Followings

                                    </p>

                                  </div>

                                </div>

                              </div>



                            </div>



                          </div>

                        </div>



                      </div>

                    ))}

                  </div>

                )}

                {activeMainTab === "following" && (

                  <div className="row card-view">

                    {usersitem.map((item) => (

                      <div className="col-lg-6 col-md-6" key={item.Title}>

                        <div

                          style={{ border: "1px solid #54ade0" }}

                          className="text-center card mb-3"

                        >

                          <div className="card-body">



                            <div className="pt-2 pb-2">

                              <a style={{ position: "relative" }}>

                                <img

                                  src={require("../assets/calling.png")}

                                  className="alignright"

                                  onClick={() =>

                                    window.open(

                                      "https://teams.microsoft.com",

                                      "_blank"

                                    )

                                  }

                                  alt="Call"

                                />

                                {/* <img

                                  src={

                                    item.Picture != null

                                      ? `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${item.EMail}`

                                      : require("../assets/users.jpg")

                                  }

                                  className="rounded-circlecss img-thumbnail

                                 avatar-xl"

                                  alt="profile-image"

                                  style={{ cursor: "pointer" }}

                                /> */}
                                {item?.SPSPicturePlaceholderState == 0 ?
                                        <img
                                          src={

                                            `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${item?.EMail}`

                                          }
                                          className="rounded-circle"
                                          width="50"
                                          alt={item?.Title}
                                        />
                                        :
                                        item?.EMail !== null &&
                                        <Avatar sx={{ bgcolor: 'primary.main' }} className="rounded-circle avatar-xl">
                                          {`${item?.EMail?.split('.')[0].charAt(0)}${item?.EMail?.split('.')[1].charAt(0)}`.toUpperCase()}
                                        </Avatar>
                                      }

                              </a>

                              <h4 className="mt-2 mb-1">

                                <a

                                  className="text-dark font-16 fw-bold"

                                  style={{

                                    textDecoration: "unset",

                                    fontSize: "20px",

                                  }}

                                >

                                  <strong>

                                    {truncateText(item.Title, 15)}

                                  </strong>

                                </a>

                              </h4>



                              <p

                                className="text-muted"

                                style={{ fontSize: "14px" }}

                              >

                                <span data-tooltip={item.EMail}>

                                  {truncateText(item.EMail, 15)} |

                                </span>

                                <span

                                  className="pl-2"

                                  style={{ color: "pink" }}

                                >


                                  {truncateText(

                                    item.Department != null

                                      ? item.Department

                                      : " NA ",

                                    10

                                  )}





                                </span>

                              </p>

                              <div

                                style={{

                                  display: "flex",

                                  justifyContent: "center",

                                  gap: "0.5rem",

                                }}

                              >

                                <button

                                  type="button"

                                  onClick={() =>

                                    window.open(

                                      "https://teams.microsoft.com",

                                      "_blank"

                                    )

                                  }

                                  className="btn btn-primary btn-sm waves-effect waves-light"

                                >

                                  Message

                                </button>

                                <button

                                  type="button"

                                  className="btn btn-light btn-sm waves-effect"

                                >

                                  Follow

                                </button>

                              </div>

                              <div className="row mt-2">

                                <div className="col-4">

                                  <div className="mt-3">

                                    <h4

                                      className="fw-bold font-14"

                                      style={{

                                        fontSize: "0.80rem",

                                        color: "#343a40",

                                      }}

                                    >

                                      NA

                                    </h4>

                                    <p className="mb-0 text-muted text-truncate">

                                      Post

                                    </p>

                                  </div>

                                </div>

                                <div className="col-4">

                                  <div className="mt-3">

                                    <h4

                                      className="fw-bold font-14"

                                      style={{

                                        fontSize: "0.80rem",

                                        color: "#343a40",

                                      }}

                                    >

                                      NA

                                    </h4>

                                    <p className="mb-0 text-muted text-truncate">

                                      Followers

                                    </p>

                                  </div>

                                </div>

                                <div className="col-4">

                                  <div className="mt-3">

                                    <h4

                                      className="fw-bold font-14"

                                      style={{

                                        fontSize: "0.80rem",

                                        color: "#343a40",

                                      }}

                                    >

                                      NA

                                    </h4>

                                    <p className="mb-0 text-muted text-truncate">

                                      Followings

                                    </p>

                                  </div>

                                </div>

                              </div>


                            </div>



                          </div>

                        </div>



                      </div>

                    ))}

                  </div>

                )}

              </div>

              <div className="col-md-3 mobile-w3">

                <div className="card mobile-5" style={{ borderRadius: "1rem" }}>

                  <div className="card-body pb-3 gheight">



                    <h4 className="header-title font-16 text-dark fw-bold mb-0" style={{ fontSize: '20px' }}>

                      Group Owner

                      {/* <a

                        style={{ float: "right" }}

                        className="font-11 btn btn-primary  waves-effect waves-light view-all"

                        onClick={(e) => GotoNextPageone(e, "DiscussionForum")}

                      >

                        View All

                      </a> */}

                    </h4>
                    {/* 
                    {

                      DiscussionData.length > 0 ? DiscussionData.map(x => {

                        return (

                          <div style={{ margin: '15px 0px 10px 0px', padding: '0px 0px 10px 0px' }} className="d-flex align-items-start border-bottom ng-scope">

                       <h1>
                        {currentUsername} is Owner of this Group | 
                       </h1>

                            <TrendingUp size={18} style={{ marginTop: '5px' }} color='#1faee3' /> &nbsp;

                            <div className="w-100" style={{ fontWeight: '100' }}>

                              <a className="font-14" style={{ fontSize: '14px' }}>

                                <strong className="text-dark" style={{ fontWeight: '700' }}>{x?.DiscussionForumCategory?.CategoryName}:</strong> &nbsp;

                                <span className="text-muted" style={{ color: '#6b6b6b' }}>

                                  {x.Topic}

                                </span>

                              </a>

 

                            </div> 

                          </div>

                        )

                      }

                      ) : null

                    } */}
                    <div className="displcenter">
                      {/* <img
                        src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${ArrDetails[0]?.Author.EMail}`}
                        className="rounded-circlecss6 img-thumbnail avatar-xl"
                        alt="profile-image"
                      /> */}
                      {ArrDetails[0]?.Author.SPSPicturePlaceholderState == 0 ?
                                        <img
                                          src={

                                            `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${ArrDetails[0]?.Author.EMail}`

                                          }
                                          className="rounded-circle"
                                          width="50"
                                          alt={ArrDetails[0]?.Author.Title}
                                        />
                                        :
                                        ArrDetails[0]?.Author.EMail !== null &&ArrDetails[0]?.Author.EMail !== undefined &&ArrDetails[0]?.Author.EMail !== "" &&
                                        <Avatar sx={{ bgcolor: 'primary.main' }} className="rounded-circle avatar-xl">
                                          {`${ArrDetails[0]?.Author.EMail?.split('.')[0].charAt(0)}${ArrDetails[0]?.Author.EMail?.split('.')[1].charAt(0)}`.toUpperCase()}
                                        </Avatar>
                                      }
                      </div>

                    <h1 className='text-muted font-14 mt-3'>
                      <p className='text-dark font-16 text-center mb-2'> {ArrDetails[0]?.Author.Title}</p>
                      <p className='text-muted font-14 text-center mb-1'>{userjobtitle}</p>
                      <p className='text-muted font-12 text-center'>{ArrDetails[0]?.Author.EMail}  </p>

                    </h1>
                  </div>

                </div>



                <div className="card mobile-6" style={{ borderRadius: "1rem", position: 'sticky', top: '90px' }}>

                  <div className="card-body pb-0 gheight">

                    <h4 className="header-title font-16 text-dark fw-bold mb-0" style={{ fontSize: '20px' }}>

                      Group you might <br /> be interested in
                      <a className="font-11 view-all  btn btn-primary  waves-effect waves-light" style={{ float: 'right', lineHeight: '21px', right: '10px' }}>View All</a>
                      {/* <a

                        style={{ float: "right" }}

                        className="font-11 view-all  btn btn-primary  waves-effect waves-light"

                        onClick={(e) => GotoNextPageone(e, 'CorporateDirectory')}

                      >

                        View All

                      </a> */}

                    </h4>

                    <div className="inbox-widget mt-4">
                      {console.log("getAllgroup getAllgroup", getAllgroup)}
                      {getAllgroup.map((user: any, index: 0) => (

                        <div

                          key={index}

                          className="d-flex border-bottom pb-3 heit8 align-items-start w-100 justify-content-between mb-3"

                        >

                          {/*  <div className="col-sm-2 ">

                            <a>

                              <img

                                src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${user.EMail}`}

                                className="rounded-circle"

                                width="50"

                                alt={user.name}

                              /> 

                            </a>

                          </div> */}

                          <div className="col-sm-10">

                            <a>
                              <div style={{ width: '160px' }} className="gfg_tooltip">

                                <p className="fw-bold mt-1 font-14 mb-0 text-dark namcss" style={{ fontSize: '14px' }}>

                                  {user.GroupName}

                                </p>
                                <span style={{ left: '50%' }} className="gfg_text">
                                  {user.GroupName}
                                </span>

                              </div>

                            </a>



                          </div>

                          <div className="col-sm-2 txtr">

                            <PlusCircle onClick={() => addFollowedGroup(user)} className='newhoverzom' size={20} color='#008751' />

                          </div>

                        </div>

                      ))}

                    </div>

                  </div>

                </div>

              </div>

            </div>)}

            {activeTab === "groupsyoucreated" && (<div className="row mt-3">

              {groupsData.map((groupItem: any, index: number) => (

                <div

                  key={index}

                  className="col-lg-6 col-xl-3 position-relative ng-scope"

                >

                  <div

                    style={{ background: "#12a8de", color: "#fff" }}

                    className="card heightcard"

                  >

                    <div className="btn-light font-12 rounded-pill text-primary waves-effect fw-bold waves-light btn-lightcss">

                      <Users size={16} /> Follow Group

                    </div>

                    <div className="card-body">

                      <a key={index}>

                        <div className="bg">
                          <img src={require("../assets/backteam.png")} />



                        </div>

                        <h4

                          style={{ lineHeight: "26px", width: "79%", float: "left", color: "#fff", textTransform: 'capitalize' }}

                          className="card-title fw-bold font-20 mb-1 mt-0"

                        >

                          {groupItem.GroupName} ({groupItem?.GroupType})

                        </h4>

                      </a>

                    </div>

                  </div>

                </div>

              ))}

            </div>)}

          </div>

        </div>

      </div>

    </div >



  )

}




const SocialFeed: React.FC<ISocialFeedProps> = (props) => {

  return (

    <Provider>

      <SocialFeedContext props={props} />

    </Provider>

  );

};

export default SocialFeed;