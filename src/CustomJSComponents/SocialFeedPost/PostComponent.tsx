import React, { useEffect, useRef } from "react";
import { useState } from "react";

import "../../CustomJSComponents/SocialFeedPost/PostComponent.scss"

import { addActivityLeaderboard, addNotification, getCurrentUserNameId, getuserprofilepic, getUserProfilePicture, getUserSPSPicturePlaceholderState } from "../../APISearvice/CustomService";

import { Heart, Menu, MessageSquare, MoreHorizontal, MoreVertical, Share, Share2, ThumbsUp } from "react-feather";

import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";

import { faHeart, faThumbsUp } from "@fortawesome/free-solid-svg-icons";

import { Carousel, Modal, Spinner } from "react-bootstrap";

import moment from "moment";

import Swal from "sweetalert2";
import Avatar from "@mui/material/Avatar";

export const PostComponent = ({ key, sp, siteUrl, currentUsername, CurrentUser, currentEmail, editload, post }: any) => {
    const [loadingReply, setLoadingReply] = useState<boolean>(false);
    const [loadingLike, setLoadingLike] = useState<boolean>(false);

    const [liked, setLiked] = useState(post.userHasLiked);

    const [likesCount, setLikesCount] = useState(post.likecount || 0);

    const [CommentsCount, setCommentsCount] = useState(post.commentcount || 0);
    // const [gcomments, setGComments] = useState(post.gcomments || []);

    const [posts, setPosts] = useState([post]);

    const [comments, setComments] = useState(post.comments || []);

    const [SocialFeedImagesJson, setImages] = useState([]);

    const [authorId, setAuthorId] = useState(0);

    const [editedContent, setEditedContent] = useState(post.Contentpost);

    const [newComment, setNewComment] = useState('');

    const [isEditing, setIsEditing] = useState(false);

    const isPostAuthor = post.userName === currentUsername;

    const [copySuccess, setCopySuccess] = useState('');

    const [IsCall, setIsCall] = useState(true)
    const [showMore, setShowMore] = useState(false);
    const menuRef = useRef(null);
    const [showModal, setShowModal] = useState(false);
    const [CurrentuserPicturePlaceholderState, setCurrentuserPicturePlaceholderState] = useState("");
    const [CurrenuserProfilepic, SetCurrenuserProfilepic] = useState(null);
    const [currentImageIndex, setCurrentImageIndex] = useState(0);

    const [loading, setLoading] = useState(false); // To show loading spinner


    const [displayedCount, setDisplayedCount] = useState();
    useEffect(() => {
        //initializeData();

        GetId()

        if (typeof (post.SocialFeedImagesJson) == 'string') {

            if (post.SocialFeedImagesJson != null && post.SocialFeedImagesJson != undefined && post.SocialFeedImagesJson != "") {

                setImages(JSON.parse(post.SocialFeedImagesJson))

            }
        }
        const initializeData = async () => {
            if (IsCall) {
                // const userId = await getCurrentUserNameId(sp);

                setAuthorId(CurrentUser.Id);
                setIsCall(false)
                fetchInitialLikeData(CurrentUser.Id);
                SetCurrenuserProfilepic(await getUserProfilePicture(CurrentUser.Id, sp));
                setCurrentuserPicturePlaceholderState(await getuserprofilepic(sp, currentEmail))
                // setCurrentuserPicturePlaceholderState(await getUserSPSPicturePlaceholderState(CurrentUser.Id, sp))
            }


        };

        initializeData();

        const handleClickOutside = (event: { target: any; }) => {
            if (menuRef.current && !menuRef.current.contains(event.target)) {
                setIsMenuOpenshare(false);
                setIsMenuOpen(false);
            }
        };

        document.addEventListener('mousedown', handleClickOutside);

        return () => {
            document.removeEventListener('mousedown', handleClickOutside);
        };

    }, [post, sp, siteUrl])

    const [isMenuOpen, setIsMenuOpen] = useState(false);

    const [isMenuOpenshare, setIsMenuOpenshare] = useState(false);


    const toggleMenu = (e: any) => {

        e.preventDefault()

        setIsMenuOpen(!isMenuOpen);

    };

    const toggleMenushare = (e: any) => {

        e.preventDefault()

        setIsMenuOpenshare(!isMenuOpen);

    };


    const GetId = async () => {

        setAuthorId(await getCurrentUserNameId(sp))

        // fetchInitialLikeData();

    }



    const handleEditClick = (e: any) => {



        setIsMenuOpen(!isMenuOpen);

        setIsEditing(true);

    };



    const handleCancelEdit = (e: { preventDefault: () => void; }) => {

        e.preventDefault()

        setIsEditing(false);

        setEditedContent(post.Contentpost);

    };

    const handleSaveEdit = async (e: any) => {

        //  e.preventDefault();

        const postId = post.postId;

        try {

            await sp.web.lists.getByTitle('ARGSocialFeed').items.getById(post.postId).update({

                Contentpost: editedContent

            });

            setIsEditing(false);

            // window.location.reload()
            editload && editload()
            setEditedContent(editedContent);



        } catch (error) {

            console.error('Error updating post:', error);

        }

    };

    const fetchInitialLikeData = async (userId: number) => {

        const postId = post.postId;



        await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.select("*,Author/Id").expand("Author")

            .filter(`ARGSocialFeedPostsId eq ${postId} and AuthorId eq ${userId}`)().then(async (ele: any) => {

                if (ele.length > 0)

                    setLiked(ele[0].userHasLiked);

                // setLikesCount(ele.length);

            })



        await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.select("*,Author/Id").expand("Author")

            .filter(`ARGSocialFeedPostsId eq ${postId} and userHasLiked eq 1`)().then(async (ele: any) => {

                if (ele.length > 0)

                    setLikesCount(ele.length);

            })



        await sp.web.lists.getByTitle('ARGSocialFeedComments').items.select("*,Author/Id,Author/SPSPicturePlaceholderState").expand("Author")

            .filter(`ARGSocialFeedId eq ${postId}`)().then(async (ele: any) => {

                if (ele.length > 0)

                    setCommentsCount(ele.length);

            })

        await sp.web.lists.getByTitle('ARGSocialFeedComments').items.select("*,Author/Id,Author/Title,Author/EMail,Author/SPSPicturePlaceholderState").expand("Author")

            .filter(`ARGSocialFeedId eq ${postId}`)().then(async (ele: any) => {

                if (ele.length > 0)

                    setComments(ele);

            })


    };

    const fetchInitialGroupLikeData = async (userId: number) => {

        const postId = post.postId;
        await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.select("*,Author/Id").expand("Author")
            .filter(`ARGSocialFeedPostsId eq ${postId} and AuthorId eq ${userId}`)().then(async (ele: any) => {
                if (ele.length > 0)
                    setLiked(ele[0].userHasLiked);
                // setLikesCount(ele.length);
            })

        await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.select("*,Author/Id").expand("Author")
            .filter(`ARGSocialFeedPostsId eq ${postId} and userHasLiked eq 1`)().then(async (ele: any) => {
                if (ele.length > 0)
                    setLikesCount(ele.length);
            })

        await sp.web.lists.getByTitle('ARGSocialFeedComments').items.select("*,Author/Id").expand("Author")
            .filter(`ARGSocialFeedId eq ${postId}`)().then(async (ele: any) => {
                if (ele.length > 0)
                    setCommentsCount(ele.length);
            })

        // await sp.web.lists.getByTitle('ARGGroupandTeamComments').items.select("*,Author/Id,Author/Title").expand("Author")
        //     .filter(`GroupandTeamId eq ${postId}`)().then(async (ele: any) => {
        //         if (ele.length > 0)
        //             setGComments(ele);
        //     })



    };
    const handleShowMore = () => {
        // setDisplayedCount((prevCount) => prevCount + 3); // Show 3 more comments
    };

    // Handle "Show Less" button click
    const handleShowLess = () => {
        // setDisplayedCount(3); // Reset to initial 3 comments
    };
    const fetchPosts = async () => {

        try {

            const fetchedPosts = await sp.web.lists.getByTitle('ARGSocialFeed').items.getAll();

            setPosts(fetchedPosts);

        } catch (error) {

            console.error("Error fetching posts:", error);

        }

    };

    const handleLike = async (e: any, liked: boolean) => {
        debugger
        e.preventDefault();
        setLoadingLike(true)
        try {
            const updatedLikeStatus = !liked;
            const updatedLikesCount = updatedLikeStatus ? likesCount + 1 : likesCount - 1;

            // Optimistically update the UI state
            setLiked(updatedLikeStatus);
            setLikesCount(updatedLikesCount);

            // setAuthorId(await getCurrentUserNameId(sp))
            const postId = post.postId;
            debugger
            await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.select("*,Author/Id").expand("Author").filter(`ARGSocialFeedPostsId eq ${postId} and AuthorId eq ${authorId}`).top(1)().then(async (ele: any) => {
                console.log(ele, 'ele');
                debugger
                if (ele.length > 0) {
                    var likePosts = {
                        ARGSocialFeedPostsId: postId,
                        userHasLiked: !ele[0].userHasLiked
                    }
                    await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.getById(ele[0].Id).delete().then(async (ele1: any) => {
                        console.log(ele1);

                        await addActivityLeaderboard(sp, "Unlike on Post");

                    })
                }
                else {
                    const likePostsJson = {
                        ARGSocialFeedPostsId: postId,
                        userHasLiked: true
                    }
                    let likePostsJson1 = Array.isArray(likePostsJson) ? likePostsJson : [likePostsJson];
                    debugger
                    await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.add(likePostsJson).then(async (ele1: any) => {
                        console.log(ele1);
                        await addActivityLeaderboard(sp, "Likes on Post");
                        let notifiedArr = {

                            ContentId: post.postId,

                            NotifiedUserId: post.postAuthorId,

                            ContentType0: "Likes on Post",

                            ContentName: post.Contentpost,

                            ActionUserId: CurrentUser.Id,

                            DeatilPage: "SocialFeed",

                            ReadStatus: false,

                            ContentCommentId: postId,

                            ContentComment: post.Contentpost

                        }
                        const nofiArr = await addNotification(notifiedArr, sp)
                        console.log(nofiArr, 'nofiArr');
                    })
                }
            })
        }
        catch (error) {
            console.error('Error toggling like:', error);
        }
        finally {
            setLoadingLike(false); // Enable the button after the function completes
        }
    };



    const handleSaveOnEnter = (e: React.KeyboardEvent<HTMLTextAreaElement>) => {
        if (e.key === "Enter") {
            e.preventDefault();
            handleSaveEdit(e)
        }
    };


    const handleAddComment = async (e: { preventDefault: () => void; }) => {
        setLoadingReply(true);
        try {
            e.preventDefault();

            if (newComment.trim()) {
                const updatedComments = [...comments, { text: newComment, user: "You" }];
                // Update local state with new comment
                setComments(updatedComments);
                setNewComment('');

                try {
                    const postId = post.postId;// Assuming postId is correctly passed as a prop
                    console.log("Post ID:", postId);

                    const existingPost = await sp.web.lists.getByTitle('ARGSocialFeed').items.getById(postId)();

                    if (!existingPost) {
                        throw new Error("Post not found");
                    }
                    // Parse existing comments (if any)
                    const existingComments = existingPost.SocialFeedCommentsJson ? JSON.parse(existingPost.SocialFeedCommentsJson) : [];

                    // Create a new comment object

                    const newComments = [...existingComments, { text: newComment, user: post.userName }];

                    const commentsBody = {
                        Comments: newComment,  // Assuming Comments is the correct field name
                        ARGSocialFeedId: postId,
                        UserImage: `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentEmail}`
                    };

                    const commentResponse = await sp.web.lists.getByTitle('ARGSocialFeedComments').items.add(commentsBody);
                    console.log('Comment added:', commentResponse);
                    await sp.web.lists.getByTitle('ARGSocialFeedComments').items.select("*,Author/Id,Author/Title").expand("Author")
                        .filter(`ARGSocialFeedId eq ${postId}`)().then(async (ele: any) => {
                            if (ele.length > 0)
                                await addActivityLeaderboard(sp, "Comments on Post");
                            setCommentsCount(ele.length)
                            setComments(ele);
                            if (CurrentUser.Id != existingPost.AuthorId) {

                                let notifiedArr = {

                                    ContentId: postId,

                                    NotifiedUserId: existingPost.AuthorId,

                                    ContentType0: "Comment on post",

                                    ContentName: existingPost.Contentpost,

                                    ActionUserId: CurrentUser.Id,

                                    DeatilPage: "SocialFeed",

                                    ReadStatus: false,

                                    ContentCommentId: commentResponse.data.Id,

                                    ContentComment: newComment,

                                    CommentOnReply: ""

                                }

                                const nofiArr = await addNotification(notifiedArr, sp)

                                console.log(nofiArr, 'nofiArr');

                            }
                        })
                } catch (error) {
                    console.log("Error updating comments in SharePoint: ", error);
                }
            }
        }
        catch (error) {
            console.error('Error toggling Reply:', error);
        }
        finally {
            setLoadingReply(false); // Enable the button after the function completes
        }
    };



    const mergeAndRemoveDuplicates = (str: string, str1: string) => {
        let url = str1;

        let thirdSlashIndex = url.indexOf(
            "/",
            url.indexOf("/", url.indexOf("/") + 1) + 1
        );
        let updatedUrl = url.substring(thirdSlashIndex);

        console.log("check the url--->>", updatedUrl); // Output: /SocialFeedImages/announcement-5.jpg

        return str + updatedUrl; // Concatenate directly if str1 starts with a slash

    };

    const handleDeletePost = async (e: any, postId: number) => {

        e.preventDefault();
        Swal.fire({
            title: 'Do you want to delete?',
            showConfirmButton: true,
            showCancelButton: true,
            confirmButtonText: "Yes",
            cancelButtonText: "No",
            icon: 'warning'
        }).then(async (result) => {
            if (result.isConfirmed) {
                try {
                    // Assuming you are using SharePoint API to delete the post
                    await sp.web.lists.getByTitle('ARGSocialFeed').items.getById(postId).delete();
                    // Remove post from UI

                    setPosts(prevPosts => prevPosts.filter(post => post.postId !== postId));
                    //chhaya
                    editload && editload()

                } catch (error) {

                }
            }
        }).catch()
        {
        }
    };

    const copyToClipboard = (e?: React.MouseEvent<HTMLButtonElement>) => {
        e?.preventDefault();
        const link = `${siteUrl}/SitePages/SocialFeed.aspx`;
        navigator.clipboard.writeText(link)
            .then(() => {
                setCopySuccess('Link copied!');
                setTimeout(() => setCopySuccess(''), 2000); // Clear message after 2 seconds
            })
            .catch(err => {
                setCopySuccess('Failed to copy link');
            });
    };
    const sendanEmail = (item: any) => {

        // window.open("https://outlook.office365.com/mail/deeplink/compose?subject=Share%20Info&body=");
        //const subject = "Post Title -" + item.Contentpost;
        //const body = 'Here is the link to the Post:' + `${siteUrl}/SitePages/SocialFeed.aspx`;
        //const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        const subject = "Thought You’d Find This Interesting!";
        const body = 'Hi,' +
            'I came across something that might interest you: ' +
            `<a href="${siteUrl}/SitePages/SocialFeed.aspx"></a>`
        const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${subject}&body=${body}`;

        window.open(office365MailLink, '_blank');
        //const mailtoLink = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

        // Open the link to launch the default mail client (like Outlook)
        //window.location.href = mailtoLink;

    };

    const handleToggleImages = () => {
        // setShowMore((prevShowMore) => !prevShowMore);
        setShowModal(true); // Open the modal
        setCurrentImageIndex(0); // Reset the carousel to the first image
    };
    const sendanEmailStop = () => {
        setIsMenuOpenshare(false);
    }
    const displayedImages = showMore
        ? SocialFeedImagesJson
        : SocialFeedImagesJson.slice(0, 3);

    console.log(post.userHasLiked, '{liked}');


    const handleReportClick = async (e: any, commentRepliesObject: any, flag: string) => {
        console.log("Report Clicked");
        e.preventDefault()
        try {

            const currentUser = await sp.web.currentUser();
            const reportListName=flag === "replies" ? "ARGSocialFeedComments" : "ARGSocialFeed";
            const reportedListItemId=flag === "replies" ? commentRepliesObject.Id : post.postId;
            const eventReportData=await sp.web.lists.getByTitle("ReportedIssueList").items.select("*").filter(`ProcessName eq 'Social Feed' and ReportedById eq ${currentUser.Id} and ListName eq '${reportListName}' and ListItemId eq ${reportedListItemId}`)();
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
        wrapperDiv.className = "report-Issue-Wrapper-Div"
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
        closeButton.style.top = "10px";
        closeButton.style.right = "10px";
        closeButton.style.border = "none";
        closeButton.style.background = "transparent";
        closeButton.style.fontSize = "16px";
        closeButton.style.cursor = "pointer";
        closeButton.style.color = "Black"
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

            try {
                // const currentUser = await sp.web.currentUser();
                // const commentObject = Comments[commentId];
                console.log("flag", flag);
                console.log("commentRepliesObject", commentRepliesObject);
                const payload = {
                    ReportReason: issueValue,
                    ProcessName: "Social Feed",
                    ReportedDate: new Date(),
                    Status: "Pending",
                    ListName: flag === "replies" ? "ARGSocialFeedComments" : "ARGSocialFeed",
                    ReportedContentAddedOn: flag === "replies" ? commentRepliesObject.Created : commentRepliesObject.Created,
                    ReportedContent: flag === "replies" ? commentRepliesObject.Comments : commentRepliesObject.Contentpost,
                    ReportedById: currentUser.Id,
                    ListItemId: flag === "replies" ? commentRepliesObject.Id : post.postId,
                    ReportedContentAddedById: flag === 'replies' ? commentRepliesObject.AuthorId : commentRepliesObject.AutherId,
                    Title: post.Contentpost,
                    Action: "Active",
                    MainListColumnName: "",
                    MainListName: "",
                    MainListItemId: 0,
                    MainListStatus: "NA",
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
                console.log("payload", payload)
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

    const [openMenuIndex, setOpenMenuIndex] = useState(null);
    const menuRef1 = useRef(null);
    const toggleMenu1 = (index: any) => {
        setOpenMenuIndex(openMenuIndex === index ? null : index);
    }

    React.useEffect(() => {
        const handleClickOutside = (event: any) => {
            if (menuRef1.current && !menuRef1.current.contains(event.target)) {
                setOpenMenuIndex(null);
            }
        };
        document.addEventListener("mousedown", handleClickOutside);
        return () => {
            document.removeEventListener("mousedown", handleClickOutside);
        };
    }, []);
    return (
        <div className="post p-4">
            <div className="post-header">
                <div className="d-flex align-items-center" style={{ width: '100%' }}>
                    <div className="flex-shrink-0">
                        {/* <img src={post.userAvatar} alt="user avatar" className="avatar" /> */}

                        {post.SPSPicturePlaceholderState == 0 ?
                            <img
                                src={
                                    `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${post.AuthorEmail}`
                                }
                                alt="user avatar" className="avatar"

                            />
                            :
                            (post.AuthorEmail !== null || post.AuthorEmail !== "") &&
                            <Avatar sx={{ bgcolor: 'primary.main' }} className="rounded-circlecss img-thumbnail
                                  avatar-xl">
                                {`${post.AuthorEmail?.split('.')[0]?.charAt(0)}${post.AuthorEmail?.split('.')[1]?.charAt(0)}`.toUpperCase()}
                            </Avatar>
                        }
                    </div>
                    <div className="flex-grow-1 ">
                        <p className="pt-2 mb-1" style={{ marginBottom: 'unset' }}>
                            <strong>{post.userName}</strong>  </p>
                        <p style={{
                            fontSize: '0.75rem',
                            fontWeight: '400',
                            color: '#6c757d'
                        }}>
                            {moment(post.Created).fromNow()}
                        </p>
                    </div>
                    {/* Post Content */}
                    <div className="post-content">
                        {(
                            <><>
                            </>
                                {CurrentUser.Id === post.AutherId && (
                                    <div className="post-actions">
                                        <div className="menu-toggle" onClick={toggleMenu}>
                                            <MoreVertical size={20} />
                                        </div>
                                        {isMenuOpen && (
                                            <div className="dropdown-menucsspost" ref={menuRef}>
                                                <button onClick={(e) => handleEditClick(e)} disabled={post.AutherId != CurrentUser.Id}>Edit</button>
                                                <button onClick={(e) => handleDeletePost(e, post.postId)} disabled={post.AutherId != CurrentUser.Id}>Delete</button>
                                                <button onClick={(e) => handleReportClick(e, post, "Post")} disabled={post.AutherId != CurrentUser.Id}>Report</button>
                                            </div>
                                        )}
                                    </div>
                                )}
                            </>
                        )}
                    </div>
                    <div className="post-content">
                        {(
                            <><>
                            </>
                                {CurrentUser.Id !== post.AutherId && (
                                    <div className="post-actions">
                                        <div className="menu-toggle" onClick={toggleMenu}>
                                            <MoreVertical size={20} />
                                        </div>
                                        {isMenuOpen && (
                                            <div className="dropdown-menucsspost" ref={menuRef}>
                                                {/* <button onClick={(e) => handleEditClick(e)} disabled={post.AutherId != CurrentUser.Id}>Edit</button>
                                            <button onClick={(e) => handleDeletePost(e, post.postId)} disabled={post.AutherId != CurrentUser.Id}>Delete</button> */}
                                                <button onClick={(e) => handleReportClick(e, post, "Post")}>Report</button>
                                            </div>
                                        )}
                                    </div>
                                )}
                            </>
                        )}
                    </div>
                </div>
            </div>

            {/* Post Content */}
            <div className="post-content" onClick={() => sendanEmailStop()}>
                {isEditing ? (
                    <textarea
                        value={editedContent}
                        onChange={(e) => setEditedContent(e.target.value)}
                        rows={10}
                        className="edit-post-textarea"
                        onKeyPress={(e) => {
                            if (e.key === 'Enter' && !e.shiftKey) {
                                e.preventDefault(); // Prevents the new line in textarea
                                handleSaveOnEnter(e); // Calls the function to add comment
                            }
                        }}
                    />
                ) : (
                    // <p>{post.Contentpost}   {/* Edit Button */}
                    // </p>
                    // Code by Priyanka for Spaceing In post comment in Social feed
                    <p className="post-text">{post.Contentpost}</p>
                )}

                <div className="image-preview mt-2">
                    <div className="grid-container">
                        {displayedImages.map((image: any, index) => {
                            // Merging and cleaning up the image URL if needed
                            const imageUrl = mergeAndRemoveDuplicates(siteUrl, image.fileUrl);
                            // Assign a class for larger or smaller images based on the index
                            const className = index === 0 ? "large-image" : "small-image";
                            // Check if the current image is the third one and if there are more than 3 images
                            const isThirdImage = index === 2 && SocialFeedImagesJson.length > 3 && !showMore;
                            return (
                                <div
                                    key={index}
                                    className={`grid-item ${className}`}
                                    style={{ position: "relative" }}
                                >
                                    {/* Render the image */}
                                    <img
                                        src={imageUrl}
                                        alt={`Social feed ${index}`}
                                        style={{
                                            width: "100%",
                                            height: "auto",
                                            cursor: "pointer",
                                        }}
                                        onClick={() => {
                                            setShowModal(true); // Open the modal
                                            setCurrentImageIndex(index); // Pass the clicked image index
                                        }}
                                    />
                                    {/* Show +X overlay if it is the third image and there are more images to show */}
                                    {isThirdImage && (
                                        <div
                                            className="more-images-overlay"
                                            style={{
                                                position: "absolute",
                                                top: 0,
                                                left: 0,
                                                right: 0,
                                                bottom: 0,
                                                backgroundColor: "rgba(0, 0, 0, 0.5)",
                                                color: "white",
                                                fontSize: "24px",
                                                display: "flex",
                                                alignItems: "center",
                                                justifyContent: "center",
                                                cursor: "pointer",
                                            }}
                                            onClick={handleToggleImages} // Toggle between more or fewer images
                                        >
                                            +{SocialFeedImagesJson.length - 3}
                                        </div>
                                    )}
                                </div>
                            );
                        })}
                    </div>

                    {/* Only display a "Show Less" message if more images are shown */}

                    {showMore && (
                        <div
                            className="show-less"
                            style={{ marginTop: "10px", cursor: "pointer" }}
                            onClick={handleToggleImages}
                        >
                            Show Less
                        </div>
                    )}
                </div>

                {/* <div className="grid-container">
                    {SocialFeedImagesJson.map((image: any, index) => {
                        const imageUrl = mergeAndRemoveDuplicates(siteUrl, image.fileUrl);
                        const className = index === 0 ? 'large-image' : 'small-image'; // First image as large, rest as small
                        return (
                            <div key={index} className={`grid-item ${className}`}>
                                <img src={imageUrl} alt={`Social feed ${index}`} />
                            </div>
                        );
                    })}
                </div> */}
            </div>
            <Modal show={showModal} onHide={() => setShowModal(false)} centered>
                <Modal.Header closeButton></Modal.Header>
                <Modal.Body>
                    <Carousel
                        activeIndex={currentImageIndex} // Show the clicked image first
                        onSelect={(selectedIndex) => setCurrentImageIndex(selectedIndex)}
                    >
                        {SocialFeedImagesJson.map((item: any, index: number) => (
                            <Carousel.Item key={index}>
                                <img
                                    className="d-block w-100"
                                    src={mergeAndRemoveDuplicates(siteUrl, item.fileUrl)} // Use your image URL merge function
                                    alt={`Slide ${index}`}
                                    style={{ height: "auto", objectFit: "contain" }}
                                />
                            </Carousel.Item>
                        ))}
                    </Carousel>
                </Modal.Body>
            </Modal>

            {/* Post Interactions */}

            <div className="post-interactions mt-3 mb-3">
                <div className="likes hovertext" onClick={!loadingLike ? (e) => handleLike(e, liked) : undefined}  >
                    {liked ? <FontAwesomeIcon icon={faThumbsUp} fontSize={25} color="#1fb0e5" /> : <ThumbsUp size={20} color="gray" />}
                    <span>{likesCount} Likes</span>{liked}
                </div>
                <span className="likes"><MessageSquare size={20} /> {CommentsCount} Comments</span>
                <div className="post-actions likes hovertext">
                    <div className="menu-toggle" onClick={toggleMenushare}>
                        <Share2 size={20} /><span className="sahrenew"> Share</span>
                    </div>
                    {isMenuOpenshare && (
                        <div className="dropdown-menucsspost" ref={menuRef}>
                            <button onClick={(e) => sendanEmail(post)}>Share by email</button>
                            <button onClick={(e) => copyToClipboard(e)}>Copy link</button>
                            {copySuccess && <span className="text-success">{copySuccess}</span>}
                        </div>
                    )}
                </div>
            </div>


            {comments.length > 0
                ? comments.slice(0, displayedCount).map((comment: any, index: React.Key) => (
                    <div key={index} className="d-flex align-items-start commentss">
                        <div className="flex-shrink-0">
                            {/* <img
                                src={comment.UserImage}
                                alt="user avatar"
                                className="commentsImg"
                            /> */}

                            {comment?.Author?.SPSPicturePlaceholderState == 0 ?
                                <img
                                    src={
                                        `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${comment?.Author?.EMail}`
                                    }
                                    className="commentsImg"
                                    alt="profile-image"
                                    style={{ cursor: "auto", borderRadius: '1000px', width: "6rem", height: '6rem' }}
                                />
                                :
                                (comment?.Author?.EMail !== null || comment?.Author?.EMail !== "") &&
                                <Avatar sx={{ bgcolor: 'primary.main' }} className="rounded-circlecss img-thumbnail
                                  avatar-xl">
                                    {`${comment?.Author?.EMail?.split('.')[0]?.charAt(0)}${comment?.Author?.EMail?.split('.')[1]?.charAt(0)}`.toUpperCase()}
                                </Avatar>
                            }
                        </div>
                        <div className="flex-grow-1 ms-2">
                            <p className="mb-1 fw-bold">{comment?.Author?.Title}</p>
                            <p
                                style={{
                                    fontSize: "0.9rem",
                                    fontWeight: "400",
                                    color: "#6c757d",
                                    marginBottom: "0px",
                                }}
                            >
                                {comment.Comments}
                            </p>
                        </div>

                        <div className="post-content">
                            {(
                                <><>
                                </>
                                    {/* {CurrentUser.Id !== post.AutherId && ( */}
                                    <div className="post-actions">
                                        <div className="menu-toggle" onClick={() => toggleMenu1(comment.Id)}>
                                            <MoreVertical size={20} />
                                        </div>
                                        {openMenuIndex === comment.Id && (
                                            <div className="dropdown-menucsspost" ref={menuRef1}>
                                                {/* <button onClick={(e) => handleEditClick(e)} disabled={post.AutherId != CurrentUser.Id}>Edit</button>
                                            <button onClick={(e) => handleDeletePost(e, post.postId)} disabled={post.AutherId != CurrentUser.Id}>Delete</button> */}
                                                <button onClick={(e) => handleReportClick(e, comment, "replies")}>Report</button>
                                            </div>
                                        )}
                                    </div>
                                    {/* )} */}
                                </>
                            )}
                        </div>
                    </div>
                ))
                : ""}

            {/* <div className="mt-3">
              
                {comments.length > displayedCount && (
                    <div
                        
                        onClick={handleShowMore}
                    
                    >
                        Show More
                    </div>
                )}

                
                {displayedCount > 3 && (
                    <div
                        onClick={handleShowLess}
                    
                    >
                        Show Less
                    </div>
                )}
            </div> */}
            {/* <div className="post-comments">
                {comments.length > 0 ? comments.map((comment: any, index: React.Key) => (
                    <div className="comment" key={index}>
                        <img src={comment.UserImage} alt="user avatar" />
                        <p>
                            <strong>{comment?.Author?.Title}</strong>  </p>
                        <p>
                            {comment.Comments}
                        </p>
                    </div>
                )) : ""}
            </div> */}
            {/* Add a New Comment */}
            <form onSubmit={(e) => handleAddComment(e)} className="add-comment" style={{ gap: '1rem' }}>
                {console.log("CurrenuserProfilepicnmpostcompo", CurrenuserProfilepic, CurrentuserPicturePlaceholderState, currentEmail)}
                {currentEmail !== "" && CurrenuserProfilepic != null && Number(CurrentuserPicturePlaceholderState) == 0 ?
                    <img src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentEmail}`} alt="user avatar" className="commentsImg" />
                    :
                    currentEmail !== "" &&
                    <Avatar sx={{ bgcolor: 'primary.main' }} className="commentsImg img-thumbnail
                                  avatar-xl">
                        {`${currentEmail?.split('.')[0]?.charAt(0)}${currentEmail?.split('.')[1]?.charAt(0)}`.toUpperCase()}
                    </Avatar>
                }

                <textarea
                    value={newComment}
                    onChange={(e) => setNewComment(e.target.value)}
                    placeholder="Write a comment..."
                    disabled={loadingReply}
                    onKeyPress={(e) => {
                        if (e.key === 'Enter' && !e.shiftKey) {
                            e.preventDefault(); // Prevents the new line in textarea
                            handleAddComment(e); // Calls the function to add comment
                        }
                    }}
                />
                {/* <textarea
                    value={newComment}
                    onChange={(e) => setNewComment(e.target.value)}
                    placeholder="Write a comment..."
                /> */}
            </form>
        </div>

    );

};