import React, { useEffect, useRef } from "react";
import { useState } from "react";

import "../GroupTeamPost/GroupPostComponent.scss"

import { addActivityLeaderboard, addNotification, getCurrentUserNameId, getCurrentUserProfileEmail, getuserprofilepic, getUserProfilePicture, getUserSPSPicturePlaceholderState } from "../../APISearvice/CustomService";

import { Heart, Menu, MessageSquare, MoreHorizontal, MoreVertical, Share, Share2, ThumbsUp } from "react-feather";

import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";

import { faHeart, faThumbsUp } from "@fortawesome/free-solid-svg-icons";

import moment from "moment";

import Swal from "sweetalert2";

import SocialFeed, { IGroupAndTeamPosts } from "../../CustomJSComponents/GroupTeamPost/SocialFeed2";
import { Carousel, Modal } from "react-bootstrap";
import Avatar from "@mui/material/Avatar";

export const GroupPostComponent = ({ key, sp, siteUrl, isedit, currentUsername, CurrentUser, currentEmail, fetchPosts, post,groupsArray,SPSPicturePlaceholderState }: any) => {
    const [loadingReply, setLoadingReply] = useState<boolean>(false);
    const [loadingLike, setLoadingLike] = useState<boolean>(false);
    const [liked, setLiked] = useState(post.userHasLiked);
    const [likesCount, setLikesCount] = useState(post.LikesCount || 0);
    const [CurrentuserPicturePlaceholderState, setCurrentuserPicturePlaceholderState] = useState(null);
    const [CurrenuserProfilepic, SetCurrenuserProfilepic] = useState(null);
    const [CommentsCount, setCommentsCount] = useState(post.CommentsCount || 0);
    const [posts, setPosts] = useState([post]);
    const [comments, setComments] = useState(post.comments || []);
    // const [gcomments, setGComments] = useState(post.gcomments || []);
    const [SocialFeedImagesJson, setImages] = useState([]);
    const [authorId, setAuthorId] = useState(0);
    const [editedContent, setEditedContent] = useState(post.Contentpost);
    const [newComment, setNewComment] = useState('');
    const [isEditing, setIsEditing] = useState(false);
    const isPostAuthor = post.userName === currentUsername;
    const [copySuccess, setCopySuccess] = useState('');
    const [showMore, setShowMore] = useState(false);
    const menuRef = useRef(null);
    const [showModal, setShowModal] = useState(false);

    const [currentImageIndex, setCurrentImageIndex] = useState(0);
    let editedContentnew: string = "";

    useEffect(() => {
        GetId()
        if (typeof (post.SocialFeedImagesJson) == 'string') {
            if (post.SocialFeedImagesJson != null && post.SocialFeedImagesJson != undefined && post.SocialFeedImagesJson != "") {
                setImages(JSON.parse(post.SocialFeedImagesJson))
            }
        }

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


    useEffect(() => {
        initializeData();
    }, []);

    const [isMenuOpen, setIsMenuOpen] = useState(false);

    const [isMenuOpenshare, setIsMenuOpenshare] = useState(false);

    const initializeData = async () => {
        const userId = await getCurrentUserNameId(sp);
        setAuthorId(userId);
        fetchInitialGroupLikeData(userId);
        SetCurrenuserProfilepic(await getUserProfilePicture(userId, sp));
        const profileemail = await getCurrentUserProfileEmail(sp);
        setCurrentuserPicturePlaceholderState( await getuserprofilepic(sp,profileemail));
        // setCurrentuserPicturePlaceholderState(await getUserSPSPicturePlaceholderState(userId, sp))
    };

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
    }



    const handleEditClick = (e: any) => {
        setIsMenuOpen(!isMenuOpen);
        setIsEditing(true);
    };

    const fetchPosts1 = async () => {
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
                    console.log(results);
                    if (results) {
                        debugger;
                        const PostItems = results.map((ele: IGroupAndTeamPosts) => {
                            return {
                                Contentpost: ele.Comments,
                                GroupTeamImagesJson: ele.GroupTeamImagesJson,
                                Created: ele.Created,
                                userName: ele.Author?.Title,
                                userAvatar: ele.UserProfile,
                                likecount: ele.LikesCount,
                                commentcount: ele.CommentsCount,
                                Id: ele.Id,
                                SocialFeedUserLikesJson: ele.UserLikesJSON ? JSON.parse(ele.UserLikesJSON) : [],
                                gcomments: ele.UserCommentsJSON ? JSON.parse(ele.UserCommentsJSON) : [],
                            }
                        });
                        setPosts(PostItems);

                    }
                });
        } catch (error) {
            console.log("Error fetching posts:", error);
        }
    };

    const handleCancelEdit = (e: { preventDefault: () => void; }) => {

        e.preventDefault()

        setIsEditing(false);

        setEditedContent(post.Comments);

    };

    const handleSaveEdit = async (e: any) => {

        //  e.preventDefault();

        const postId = post.postId;

        try {

            await sp.web.lists.getByTitle('ARGGroupandTeamComments').items.getById(post.postId).update({

                Comments: editedContent

            });

            setIsEditing(false);
            post.Contentpost = editedContent;
            // window.location.reload()
            editedContentnew = editedContent;
            setEditedContent(editedContent);
            debugger
            //await fetchPosts()
            fetchPosts && fetchPosts();



        } catch (error) {

            console.error('Error updating post:', error);

        }

    };

    // const fetchInitialLikeData = async (userId: number) => {

    //     const postId = post.postId;
    //     await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.select("*,Author/Id").expand("Author")
    //         .filter(`ARGSocialFeedPostsId eq ${postId} and AuthorId eq ${userId}`)().then(async (ele: any) => {
    //             if (ele.length > 0)
    //                 setLiked(ele[0].userHasLiked);
    //             // setLikesCount(ele.length);
    //         })

    //     await sp.web.lists.getByTitle('ARGSocialFeedPostsUserLikes').items.select("*,Author/Id").expand("Author")
    //         .filter(`ARGSocialFeedPostsId eq ${postId} and userHasLiked eq 1`)().then(async (ele: any) => {
    //             if (ele.length > 0)
    //                 setLikesCount(ele.length);
    //         })

    //     await sp.web.lists.getByTitle('ARGSocialFeedComments').items.select("*,Author/Id").expand("Author")
    //         .filter(`ARGSocialFeedId eq ${postId}`)().then(async (ele: any) => {
    //             if (ele.length > 0)
    //                 setCommentsCount(ele.length);
    //         })

    //     await sp.web.lists.getByTitle('ARGSocialFeedComments').items.select("*,Author/Id,Author/Title").expand("Author")
    //         .filter(`ARGSocialFeedId eq ${postId}`)().then(async (ele: any) => {
    //             if (ele.length > 0)
    //                 setComments(ele);
    //         })

    // };

    const fetchInitialGroupLikeData = async (userId: number) => {

        const postId = post.postId;
        await sp.web.lists.getByTitle('ARGGroupandTeamUserLikes').items.select("*,Author/Id,Author/EMail").expand("Author")
            .filter(`GroupandTeamCommentsId eq ${postId} and AuthorId eq ${userId}`)().then(async (ele: any) => {
                if (ele.length > 0)
                    setLiked(ele[0].userHasLiked);
                // setLikesCount(ele.length);
            })

        await sp.web.lists.getByTitle('ARGGroupandTeamUserLikes').items.select("*,Author/Id,Author/EMail").expand("Author")
            .filter(`GroupandTeamCommentsId eq ${postId} and userHasLiked eq 1`)().then(async (ele: any) => {
                if (ele.length > 0)
                    setLikesCount(ele.length);
            })

        await sp.web.lists.getByTitle('ARGGroupandTeamUserComments').items.select("*,Author/Id,Author/EMail").expand("Author")
            .filter(`GroupandTeamCommentsId eq ${postId}`)().then(async (ele: any) => {
                if (ele.length > 0)
                    setCommentsCount(ele.length);
            })

        await sp.web.lists.getByTitle('ARGGroupandTeamUserComments').items.select("*,Author/Id,Author/EMail,Author/Title,Author/SPSPicturePlaceholderState").expand("Author")
            .filter(`GroupandTeamCommentsId eq ${postId}`)().then(async (ele: any) => {
                if (ele.length > 0)
                    setComments(ele);
            })

    };

    // const fetchPosts = async () => {
    //     try {
    //         const fetchedPosts = await sp.web.lists.getByTitle('ARGGroupandTeamComments').items.getAll();
    //         setPosts(fetchedPosts);
    //     } catch (error) {
    //         console.error("Error fetching posts:", error);
    //     }
    // };

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
            await sp.web.lists.getByTitle('ARGGroupandTeamUserLikes').items.select("*,Author/Id,Author/EMail").expand("Author").filter(`GroupandTeamCommentsId eq ${postId} and AuthorId eq ${authorId}`).top(1)().then(async (ele: any) => {
                console.log(ele, 'ele');
                debugger
                if (ele.length > 0) {
                    var likePosts = {
                        GroupandTeamCommentsId: postId,
                        userHasLiked: !ele[0].userHasLiked
                    }
                    await sp.web.lists.getByTitle('ARGGroupandTeamUserLikes').items.getById(ele[0].Id).delete().then(async (ele1: any) => {
                        console.log(ele1);

                        await addActivityLeaderboard(sp, "Unlike on Post");

                    })
                }
                else {
                    const likePostsJson = {
                        GroupandTeamCommentsId: postId,
                        userHasLiked: true
                    }
                    let likePostsJson1 = Array.isArray(likePostsJson) ? likePostsJson : [likePostsJson];
                    debugger
                    await sp.web.lists.getByTitle('ARGGroupandTeamUserLikes').items.add(likePostsJson).then(async (ele1: any) => {
                        console.log(ele1);
                        await addActivityLeaderboard(sp, "Likes on Post");
                        let notifiedArr = {
                            ContentId: ele.data.Id,
                            NotifiedUserId: ele.data.AuthorId,
                            ContentType0: "Likes on Post",
                            ContentName: ele.data.Title,
                            ActionUserId: CurrentUser,
                            DeatilPage: "Groups/Team",
                            ReadStatus: false
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
        debugger
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

                    const existingPost = await sp.web.lists.getByTitle('ARGGroupandTeamComments').items.getById(postId)();

                    if (!existingPost) {
                        throw new Error("Post not found");
                    }
                    // Parse existing comments (if any)
                    const existingComments = existingPost.GroupTeamCommentsJson ? JSON.parse(existingPost.GroupTeamCommentsJson) : [];

                    // Create a new comment object

                    const newComments = [...existingComments, { text: newComment, user: post.userName }];

                    const commentsBody = {
                        Comments: newComment,  // Assuming Comments is the correct field name
                        GroupandTeamCommentsId: postId,
                        UserProfile: `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentEmail}`
                    };

                    const commentResponse = await sp.web.lists.getByTitle('ARGGroupandTeamUserComments').items.add(commentsBody);
                    console.log('Comment added:', commentResponse);
                    await sp.web.lists.getByTitle('ARGGroupandTeamUserComments').items.select("*,Author/Id,Author/EMail,Author/Title,Author/SPSPicturePlaceholderState").expand("Author")
                        .filter(`GroupandTeamCommentsId eq ${postId}`)().then(async (ele: any) => {
                            if (ele.length > 0)
                                await addActivityLeaderboard(sp, "Comments on Post");
                            setCommentsCount(ele.length)
                            setComments(ele);
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

                    await sp.web.lists.getByTitle('ARGGroupandTeamComments').items.getById(postId).delete();
                    // Remove post from UI

                    setPosts(prevPosts => prevPosts.filter(post => post.postId !== postId));
                    post = {};
                    //await fetchPosts();
                    fetchPosts && fetchPosts();

                    // window.location.reload()

                } catch (error) {

                }
            }
        }).catch()
        {
        }
    };

    const copyToClipboard = (Id: number, e?: React.MouseEvent<HTMLButtonElement>) => {
        e?.preventDefault();
        const ids = window.location.search;
        const originalString = ids;
        const idNum = originalString.substring(1);
        const link = `${siteUrl}/SitePages/GroupandTeamDetails.aspx?${idNum}`;
        navigator.clipboard.writeText(link)
            .then(() => {
                setCopySuccess('Link copied!');
                // setTimeout(() => setCopySuccess(''), 2000); // Clear message after 2 seconds
            })
            .catch(err => {
                setCopySuccess('Failed to copy link');
            });
    };

    const sendanEmail = () => {
        const ids = window.location.search;
        const originalString = ids;
        const idNum = originalString.substring(1);
        //const subject = "Group Details";
        //const body = 'Here is the link to the event:' + `${siteUrl}/SitePages/GroupandTeamDetails.aspx?${idNum}`;

        // window.open(`https://outlook.office365.com/mail/deeplink/compose?body= Here is the group link: "${siteUrl}/SitePages/GroupandTeamDetails.aspx?${idNum}"`);
        //const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
        const subject = "Thought You’d Find This Interesting!";
        const body = 'Hi,' +
            'I came across something that might interest you: ' +
            `<a href="${siteUrl}/SitePages/GroupandTeamDetails.aspx?${idNum}"></a>`
        const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${subject}&body=${body}`;

        window.open(office365MailLink, '_blank');
    };
    const OpenImagePreview = (imageurl: any) => {
        //#region Methods

        window.open(imageurl, "_blank"); // imge prwview type
        //#endregion
    };

    // const handleToggleImages = () => {
    //     setShowMore((prevShowMore) => !prevShowMore);
    // };
    const handleToggleImages = () => {
        // setShowMore((prevShowMore) => !prevShowMore);
        setShowModal(true);
        setCurrentImageIndex(0);
    };

    const sendanEmailStop = () => {
        setIsMenuOpenshare(false);
    }
    const displayedImages = showMore
        ? SocialFeedImagesJson
        : SocialFeedImagesJson.slice(0, 3);

    console.log(post, '{liked}');

    const handleReportClick = async (e: any, commentRepliesObject: any, flag: string) => {
        console.log("Report Clicked");
        e.preventDefault()
        try {

            const currentUser = await sp.web.currentUser();
            const reportListName = flag === "replies" ? "ARGGroupandTeamUserComments" : "ARGGroupandTeamComments";
            const reportedListItemId = flag === "replies" ? commentRepliesObject.Id : post.postId;
            const eventReportData = await sp.web.lists.getByTitle("ReportedIssueList").items.select("*").filter(`ProcessName eq 'Groups And Teams' and ReportedById eq ${currentUser.Id} and ListName eq '${reportListName}' and ListItemId eq ${reportedListItemId}`)();
            console.log("eventReportData", eventReportData);

            if (eventReportData.length > 0) {
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
                    console.log(groupsArray, "groupsArray");
                    const payload = {
                        ReportReason: issueValue,
                        ProcessName: "Groups And Teams",
                        ReportedDate: new Date(),
                        Status: "Pending",
                        ListName: flag === "replies" ? "ARGGroupandTeamUserComments" : "ARGGroupandTeamComments",
                        ReportedContentAddedOn: flag === "replies" ? commentRepliesObject.Created : commentRepliesObject.Created,
                        ReportedContent: flag === "replies" ? commentRepliesObject.Comments : commentRepliesObject.Contentpost,
                        ReportedById: currentUser.Id,
                        ListItemId: flag === "replies" ? commentRepliesObject.Id : post.postId,
                        ReportedContentAddedById: flag === 'replies' ? commentRepliesObject.AuthorId : commentRepliesObject.AuthorId,
                        Title: groupsArray[0].GroupName,
                        Action: "Active",
                        MainListColumnName: flag === "replies" ? "UserCommentsJSON" : "",
                        MainListName: flag === "replies" ? "ARGGroupandTeamComments" : "",
                        MainListItemId: flag === "replies" ? post.postId : 0,
                        MainListStatus: flag === "replies" ? "Available" : "NA",
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
            console.log("Error adding the data into the list", error);
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
                        {Number(post.AuthorDetails.SPSPicturePlaceholderState) == 0 ?
                            <img

                                src={

                                    `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${post.AuthorDetails.EMail}`

                                }
                                className="me-2 mt-0 avatar-sm rounded-circle"
                                alt="User"
                            />
                            :
                            post.AuthorDetails.EMail !== null && post.AuthorDetails.EMail !== "" &&
                            <Avatar sx={{ bgcolor: 'primary.main' }} className="me-2 mt-0 avatar-sm rounded-circle font-14">
                                {`${post.AuthorDetails.EMail?.split('.')[0].charAt(0)}${post.AuthorDetails.EMail?.split('.')[1].charAt(0)}`.toUpperCase()}
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

                            <>
                                {post.Author == currentEmail &&
                                    <div className="post-actions">
                                        <div className="menu-toggle" onClick={toggleMenu}>
                                            <MoreVertical size={20} />
                                        </div>
                                        {isMenuOpen && (
                                            <div className="dropdown-menucsspost" ref={menuRef}>
                                                <button onClick={(e) => handleEditClick(e)}>Edit</button>
                                                <button onClick={(e) => handleDeletePost(e, post.postId)}>Delete</button>
                                            </div>
                                        )}
                                    </div>
                                }
                            </>
                        )}
                    </div>
                    <div className="post-content">
                        {(
                            <><>
                            </>
                                {/* {CurrentUser.Id !== post.AutherId && ( */}
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
                                {/* )} */}
                            </>
                        )}
                    </div>
                </div>
            </div>

            {/* Post Content */}
            {console.log("editedContenteditedContent", editedContent, editedContentnew)}
            <div className="post-content" onClick={() => sendanEmailStop()} style={{ whiteSpace: "pre-wrap" }}>
                {isEditing ? (
                    <textarea
                        value={editedContent}
                        onChange={(e) => setEditedContent(e.target.value)}
                        rows={4}
                        className="edit-post-textarea"
                        style={{ whiteSpace: "pre-wrap" }}
                        onKeyPress={(e) => {
                            if (e.key === 'Enter' && !e.shiftKey) {
                                e.preventDefault(); // Prevents the new line in textarea
                                handleSaveOnEnter(e); // Calls the function to add comment
                            }
                        }}
                    //onKeyDown={handleSaveOnEnter}
                    />
                ) : (
                    <p style={{ whiteSpace: "pre-wrap" }}>{post.Contentpost}   {/* Edit Button */}
                    </p>
                )}

                {/* <div className="image-preview mt-2">
                    <div className="grid-container">
                        {displayedImages.map((image: any, index) => {
                        
                            const imageUrl = mergeAndRemoveDuplicates(siteUrl, image.fileUrl);
                      
                            const className = index === 0 ? "large-image" : "small-image";
                         
                            const isThirdImage = index === 2 && SocialFeedImagesJson.length > 3 && !showMore;
                            return (image.fileType === "image/jpeg" ?
                                <div
                                    key={index}
                                    className={`grid-item ${className}`}
                                    style={{ position: "relative" }}
                                >
                                    <img
                                        src={imageUrl}
                                        alt={`Group/Team ${index}`} onClick={(e) => OpenImagePreview(imageUrl)}
                                        style={{ width: "100%", height: "auto" }}
                                    />
                                  
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
                                            onClick={handleToggleImages} 
                                        >
                                            +{SocialFeedImagesJson.length - 3}
                                        </div>
                                    )}
                                </div>
                                : <div className="item-design"
                                    key={index}                                  
                                    style={{ position: "relative" }}
                                >
                               
                                    <a target="_blank" href={imageUrl} style={{fontSize:'16px'}}>{image.fileName}</a>
                                    <br />


                                </div>
                            );
                        })}
                    </div>

             

                    {showMore && (
                        <div
                            className="show-less"
                            style={{ marginTop: "10px", cursor: "pointer" }}
                            onClick={handleToggleImages}
                        >
                            Show Less
                        </div>
                    )}
                </div> */}
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
            {isedit &&
                <div className="post-interactions hovertext mt-3 mb-3" aria-disabled={!isedit}>
                    <div className="likes hovertext" onClick={!loadingLike ? (e) => handleLike(e, liked) : undefined}  >
                        {liked ? <FontAwesomeIcon icon={faThumbsUp} fontSize={25} color="#1fb0e5" /> : <ThumbsUp size={20} color="gray" />}
                        <span>{likesCount} Likes</span>{liked}
                    </div>
                    <span className="likes"><MessageSquare size={20} /> {CommentsCount} Comments</span>
                    <div className="post-actions hovertext likes">
                        <div className="menu-toggle" onClick={toggleMenushare}>
                            <Share2 size={20} /> <span className="sahrenew"> Share</span>
                        </div>
                        {isMenuOpenshare && (
                            <div className="dropdown-menucsspost" ref={menuRef}>
                                <button onClick={(e) => sendanEmail()} type="button">Share by email</button>
                                <button onClick={(e) => copyToClipboard(post.postId)} type="button">Copy link</button>
                                {copySuccess && <span className="text-success">{copySuccess}</span>}
                            </div>
                        )}
                    </div>
                </div>

            }
            {comments.length > 0 ? comments.map((comment: any, index: React.Key) => (
                <div className="d-flex align-items-start commentss">
                    <div className="flex-shrink-0">
                        {/* <img src={comment.UserProfile} alt="user avatar" className="commentsImg" /> */}
                        {comment.Author && Number(comment.Author?.SPSPicturePlaceholderState) == 0 ?
                            <img

                                src={

                                    `${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${comment.Author.EMail}`

                                }
                                className="me-2 mt-0 avatar-sm rounded-circle"
                                alt="User"
                            />
                            :
                            comment.Author && comment.Author?.EMail !== null && comment.Author?.EMail !== "" &&
                            <Avatar sx={{ bgcolor: 'primary.main' }} className="me-2 mt-0 avatar-sm rounded-circle">
                                {`${comment.Author?.EMail?.split('.')[0].charAt(0)}${comment.Author?.EMail?.split('.')[1].charAt(0)}`.toUpperCase()}
                            </Avatar>
                        }
                    </div>
                    <div className="flex-grow-1 ms-2">
                        <p className="mb-1 fw-bold">
                            {comment?.Author?.Title}  </p>
                        <p style={{
                            fontSize: '0.9rem',
                            fontWeight: '400',
                            color: '#6c757d',
                            marginBottom: '0px'
                        }}>
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
            )) : ""}
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
                {console.log("CurrenuserProfilepicnmgrouppst",CurrenuserProfilepic,CurrentuserPicturePlaceholderState,currentEmail)}
                {currentEmail !== "" &&SPSPicturePlaceholderState !=null && SPSPicturePlaceholderState !=""&& Number(SPSPicturePlaceholderState) == 0 ?
                    <img src={`${siteUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${currentEmail}`} alt="user avatar" className="commentsImg" />
                    :
                    currentEmail !== "" &&currentEmail !== null &&currentEmail !== undefined &&
                    <Avatar sx={{ bgcolor: 'primary.main' }} className="commentsImg font-15 img-thumbnail
                                  avatar-xl">
                        {`${currentEmail?.split('.')[0]?.charAt(0)}${currentEmail?.split('.')[1]?.charAt(0)}`.toUpperCase()}
                    </Avatar>
                }


                <textarea
                    value={newComment}
                    onChange={(e) => setNewComment(e.target.value)}
                    placeholder="Write a comment..."
                    disabled={loadingReply || !isedit}
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