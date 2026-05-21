import React, { useEffect, useState, useRef } from 'react'
import "bootstrap/dist/css/bootstrap.min.css";
import "./CustomBlogWebpartTemplate.scss";
import "../../CustomCss/mainCustom.scss";
import "../../Assets/Figtree/Figtree-VariableFont_wght.ttf";
import { Share2 } from 'feather-icons-react';
import { Bookmark } from 'feather-icons-react';
import { Calendar, Link, Share } from 'feather-icons-react';
import moment from 'moment';
import { getCurrentUserProfileEmail } from "../../APISearvice/CustomService";
import ReactQuill from "react-quill";
import "react-quill/dist/quill.snow.css";
import { FontAwesomeIcon } from "@fortawesome/react-fontawesome";
import { faPaperclip, faEdit, faTrashAlt } from '@fortawesome/free-solid-svg-icons';
import Swal from 'sweetalert2';
import { uploadFile, uploadFileToLibrary } from "../../APISearvice/BlogService";
import { getCurrentUserProfileEmail, getEntity } from "../../APISearvice/CustomService";
import {
    fetchBlogdata, fetchBookmarkBlogdata, getBlogsByID, getBlogByID, GetBlogCategory, GetCategory,
    fetchPinstatus, updateItem, uploadFileBanner, DeleteBusinessAppsAPI
} from "../../APISearvice/BlogService";
import Multiselect from "multiselect-react-dropdown";
import { CONTENTTYPE_Media, LIST_TITLE_ContentMaster, LIST_TITLE_Blogs, LIST_TITLE_MyRequest,CONTENTTYPE_Blogs } from '../../Shared/Constants';
import { Modal } from 'react-bootstrap';
import { AddContentLevelMaster, AddContentMaster, getApprovalConfiguration, getLevel, UpdateContentMaster,getMyRequestBlogPending } from '../../APISearvice/ApprovalService';


const CustomBlogWebpartTemplate = ({ _sp, SiteUrl }) => {
    const siteUrl = SiteUrl;
    const [itemsToShow, setItemsToShow] = useState(5);
    const [copySuccess, setCopySuccess] = useState('');
    const [show, setShow] = useState(false)
    const [NewsData, setNews] = useState([])
    const [showDropdownId, setShowDropdownId] = useState(null);
    const [currentEmail, setEmail] = useState('');
    const [blogData, setBlogData] = useState('');
    const menuRef = useRef(null);
    const [isMenuOpenshare, setIsMenuOpenshare] = useState(false);
    const [activeTab, setActiveTab] = useState("all");
    const [filteredBlogItems, setFilteredBlogItems] = useState('');
    const [Bookmarkblogs, setBookmarkblogs] = useState('');
    const [Bookmarkstatus, setBookmarkstatus] = useState('');
    const [EnityData, setEnityData] = React.useState([]);
      const [modeValue, setmode] = React.useState(null);
   const [Loading, setLoading] = React.useState(false);
    const [ImagepostArr, setImagepostArr] = React.useState([]);
    const [ImagepostArr1, setImagepostArr1] = React.useState([]);
    const [BnnerImagepostArr, setBannerImagepostArr] = React.useState([]);
    const [ImagepostArr1Edit, setImagepostArr1Edit] = React.useState([]);
    const [BnnerImagepostArrEdit, setBannerImagepostArrEdit] = React.useState([]);
    const [ImagepostIdsArr, setImagepostIdsArr] = React.useState([]);
    const [DocumentpostIdsArr, setDocumentpostIdsArr] = React.useState([]);
    const [DocumentpostArr1, setDocumentpostArr1] = React.useState([]);
    const [DocumentpostArr, setDocumentpostArr] = React.useState([]);
    const [richTextValues, setRichTextValues] = React.useState({});
    const [CategoryData, setCategoryData] = React.useState([]);
    const [EditFormData, setEditFormData] = useState([]);
    const [editID, setEditID] = React.useState(null);
    const [editForm, setEditForm] = React.useState(false);
    const [IsBannerAdded, setBannerAdded] = React.useState(false);
    const [IsGalImageAdded, setGalImageAdded] = React.useState(false);
    const [rows, setRows] = React.useState([]);
    const [levels, setLevel] = React.useState([]);
     const [ApprovalRequestItem, setApprovalRequestItem] = React.useState(null);
      const [showModal, setShowModal] = React.useState(false);
       const [showDocTable, setShowDocTable] = React.useState(false);
       const [showImgModal, setShowImgTable] = React.useState(false);
       const [showBannerModal, setShowBannerTable] = React.useState(false);
     

    const [formData, setFormData] = React.useState({
        topic: "",
        category: "",
        entity: "",
        Type: "",
        bannerImage: null,
        // announcementGallery: null,
        // announcementThumbnail: null,
        description: "",
        overview: "",
        FeaturedAnnouncement: false,
    });
    const modules = {
        toolbar: [
            [{ size: ["small", false, "large", "huge"] }],
            ["bold", "italic", "underline", "strike", "blockquote"],
            [{ list: "ordered" }, { list: "bullet" }],
            ["link", "image"],
            [
                { list: "ordered" },
                { list: "bullet" },
                { indent: "-1" },
                { indent: "+1" },
            ],
            [
                {
                    color: [
                        "#000000",
                        "#e60000",
                        "#ff9900",
                        "#ffff00",
                        "#008a00",
                        "#0066cc",
                        "#9933ff",
                        "#ffffff",
                        "#facccc",
                        "#ffebcc",
                        "#ffffcc",
                        "#cce8cc",
                        "#cce0f5",
                        "#ebd6ff",
                        "#bbbbbb",
                        "#f06666",
                        "#ffc266",
                        "#ffff66",
                        "#66b966",
                        "#66a3e0",
                        "#c285ff",
                        "#888888",
                        "#a10000",
                        "#b26b00",
                        "#b2b200",
                        "#006100",
                        "#0047b2",
                        "#6b24b2",
                        "#444444",
                        "#5c0000",
                        "#663d00",
                        "#666600",
                        "#003700",
                        "#002966",
                        "#3d1466",
                        "custom-color",
                    ],
                },
            ],
        ],
    };
    const formats = [
        "header",
        "height",
        "bold",
        "italic",
        "underline",
        "strike",
        "blockquote",
        "list",
        "color",
        "bullet",
        "indent",
        "link",
        "image",
        "align",
        "size",
    ];
    useEffect(() => {
        //console.log("activeeee", activeTab);
        if (activeTab.toLowerCase() === "all") {
            if(blogData.length >0){
                 const filteredItems = blogData.filter(
                (item) => item.Status === "Approved"
            );
             
                 setFilteredBlogItems(filteredItems);
            }
           
           // setFilteredBlogItems(blogData);
         
        } else {
            // Find the selected category based on activeTab
            const selectedCategory = FilterOptions.find(
                (category) => category.Name.toLowerCase() === activeTab.toLowerCase()
            );
           // { //console.log("filteredMediaItemsselectedCategory", filteredBlogItems, activeTab, selectedCategory, currentEmail, blogData) }
            if (selectedCategory) {
                // Filter items based on the selected category's ID
                if (selectedCategory.Name == "Bookmarked") {
                    // const filteredItems = blogData.filter(
                    //     (item) =>item.BookmarkedBy && item.BookmarkedBy.EMail === currentEmail
                    // );
                   // //console.log("currentEmailcurrentEmail", Bookmarkblogs)
                    setFilteredBlogItems(Bookmarkblogs);
                } else {
                    if (selectedCategory.Name == "Save as Draft") {
                        const filteredItems = blogData.filter(
                            (item) => item.Status === selectedCategory.Name && (item.Author.EMail == currentEmail || item.Editor.EMail == currentEmail)
                        );
                        setFilteredBlogItems(filteredItems);
                    }
                    else if (selectedCategory.Name == "Submitted"||selectedCategory.Name == "Rejected"||selectedCategory.Name == "Approved") {
                        const filteredItems = blogData.filter(
                            (item) => item.Status === selectedCategory.Name && item.Author.EMail == currentEmail
                        );
                        setFilteredBlogItems(filteredItems);
                    }
                    else if (selectedCategory.Name == "Rework") {
                        const filteredItems = blogData.filter(
                            (item) => (item.Status === selectedCategory.Name && item.Author.EMail == currentEmail)||item.WillReworkEdit ==true
                        );
                        setFilteredBlogItems(filteredItems);
                    }
                     else {
                        const filteredItems = blogData.filter(
                            (item) => item.Status === selectedCategory.Name
                        );
                        setFilteredBlogItems(filteredItems);
                    }

                }

            } else {
                // If no category matches, show no items
                setFilteredBlogItems([]);
            }

        }
        // { //console.log("filteredMediaItemsafter", filteredBlogItems) }
        const handleClickOutside = (event) => {
            if (menuRef.current && !menuRef.current.contains(event.target)) {
                // setIsMenuOpenshare(false);
                setShowDropdownId(null)
            }
        };

        document.addEventListener('mousedown', handleClickOutside);

        return () => {
            document.removeEventListener('mousedown', handleClickOutside);
        };
    }, [_sp, SiteUrl, activeTab, blogData])

    useEffect(() => {
        setActiveTab("All");
        ApIcall();
    }, []);
    const FilterOptions = [{ id: 1, Name: "All", name: "All Published" },{ id: 2, Name: "Submitted", name: "Submitted" },
    //   { id: 3, Name: "Published", name: "Published" },
    { id: 3, Name: "Approved", name: "Approved" },
    { id: 4, Name: "Save as Draft", name: "Your Drafts" },{ id: 5, Name: "Rejected", name: "Rejected" },{ id: 6, Name: "Rework", name: "Rework" }, { id: 7, Name: "Bookmarked", name: "Bookmarked" }]

    const ApIcall = async () => {
        setLevel(await getLevel(_sp));
        setEmail(await getCurrentUserProfileEmail(_sp));
        setBlogData(await fetchBlogdata(_sp));

        setBookmarkstatus(await fetchPinstatus(_sp));
        setCategoryData(await GetCategory(_sp));
        //setBlogCategoryData(await GetBlogCategory(_sp));


   
        setBookmarkblogs(await fetchBookmarkBlogdata(_sp));
        // setActiveTab("All");
        setEnityData(await getEntity(_sp)) //Entity
        // //console.log("check data of blogs---",blogData)
        // const dataofblog = await fetchBlogdata(sp);
        // //console.log("check the data of blog",dataofblog)
        // setBlogData(dataofblog);

    }
// /**************changes**************** */
const setShowModalFunc = (bol, name) => {
    if (name == "bannerimg") {
      setShowModal(bol);
      setShowBannerTable(true);
      setShowImgTable(false);
      setShowDocTable(false);
    } else if (name == "Gallery") {
      setShowModal(bol);
      setShowImgTable(true);
      setShowBannerTable(false);
      setShowDocTable(false);
    } else {
      setShowModal(bol);
      setShowDocTable(true);
      setShowBannerTable(false);
      setShowImgTable(false);
    }
  };

  //#region deleteLocalFile
const deleteLocalFile = (index, filArray, name) => {
    debugger
    //console.log(filArray, 'filArray');
 
    // Remove the file at the specified index
    filArray.splice(index, 1);
    //console.log(filArray, 'filArray');
 
    // Update the state based on the title
    if (name === "bannerimg") {
      setBannerImagepostArr([...filArray]);
      filArray[0].files.length > 0 ? "" : setShowModal(false); clearFileInput(name);
    } else if (name === "Gallery") {
      setImagepostArr1([...filArray]);
      filArray[0].files.length > 0 ? "" : setShowModal(false); clearFileInput(name);
    } else {
      setDocumentpostArr1([...filArray]);
      filArray[0].files.length > 0 ? "" : setShowModal(false); clearFileInput(name);
    }
    // Clear the file input
 
  };

 
    const handleClick = async (contentId, contentName, EntityId) => {
     
        //  //console.log("Creating approval hierarchy with data:", rows);
     
          let boolval = false
     
          for (let i = 0; i < rows.length; i++) {
     
            const userIds = rows[i].approvedUserListupdate.map((user) => user.id);
     
            let arrPost = {
               //LevelSequence: i + 1,
        LevelSequence: rows[i].LevelId,
              ContentId: contentId,
     
              ContentName: "ARGBlogs",
     
              EntityMasterId: EntityId,
     
              ARGLevelMasterId: rows[i].LevelId,
     
              ApproverId: userIds,
     
              ApprovalType: rows[i].selectionType == "All" ? 1 : 0,
     
              SourceName: contentName
     
            }
     
            const addedData = await AddContentLevelMaster(_sp, arrPost)
     
            ////console.log("created content level master items", addedData);
     
     
          }
     
          boolval = true
     
          return boolval;
     
          // Process rows data, e.g., submit to server or save to SharePoint list
     
          // Add your submit logic here
     
        };

        const handleUserSelect = (selectedUsers, rowId) => {

            setRows((prevRows) =>
       
              prevRows.map((row) =>
       
                row.id === rowId
       
                  ? { ...row, approvedUserListupdate: selectedUsers }
       
                  : row
       
              )
       
            );
       
          };

           const [visibleItems, setVisibleItems] = React.useState(5);
         
              const handleLoadMore = () => {
                  event.preventDefault()
                  //event.stopImmediatePropagation()
                  setVisibleItems(prevVisibleItems => prevVisibleItems + 5);
              };
         
  // /**************changes**************** */

    const handleTabClick = (tab, Id) => {
        setActiveTab(tab.toLowerCase());
    };
    const truncateText = (text, maxLength) => {
        if (text != null) {
            return text.length > maxLength ? text.substring(0, maxLength) + "..." : text;

        }
    };

    const gotoBlogsDetails = (valurArr) => {
        localStorage.setItem("NewsId", valurArr.Id)
        localStorage.setItem("NewsArr", JSON.stringify(valurArr))
        setTimeout(() => {
            window.location.href = `${SiteUrl}/SitePages/BlogDetails.aspx?${valurArr.Id}`;
        }, 500);
    }

    const copyToClipboard = (Id) => {
        const link = `${SiteUrl}/SitePages/BlogDetails.aspx?${Id}`;
        navigator.clipboard.writeText(link)
            .then(() => {
                setCopySuccess('Link copied!');
                setTimeout(() => setCopySuccess(''), 2000); // Clear message after 2 seconds
            })
            .catch(err => {
                setCopySuccess('Failed to copy link');
            });
    };
    // Edit blogs
    const getDescription = (des) => {
        // You can fetch this from an API, or return dynamic content.
        return des;
    };
    const EditBlogs = async (id) => {

        let selectedData = await getBlogsByID(_sp, Number(id));
        let myrequestdata = await getMyRequestBlogPending(_sp, selectedData[0]);
        /////
        // if(name == "entity"){
       
            //ARGApprovalConfiguration
         
                 const rowData = await getApprovalConfiguration(_sp, Number(selectedData[0].Entity)) //baseUrl
         
                 const initialRows = rowData.map((item) => ({
         
                   id: item.Id,
         
                   Level: item.Level.Level,
         
                   LevelId: item.LevelId,
         
                   approvedUserListupdate: item.Users.map((user) => ({
         
                     id: user.ID,
         
                     name: user.Title,
         
                     email: user.EMail
         
                   })),
         
                   selectionType: 'All' // default selection type, if any
         
                 }));
         
                 setRows(initialRows);
   
        //  }
        //  /////
        if(myrequestdata){
            setApprovalRequestItem(myrequestdata[0]);
        }
        setEditFormData(selectedData);
        setEditID(id);
        setEditForm(true)
        // //console.log(id, "----id", selectedData);
        //setCategoryData(await getCategory(_sp, Number(selectedData[0].Category)));
        setFormData((prevData) => ({
            ...prevData,

            //Status: selectedData[0].Status,
            //BlogBannerImage: selectedData[0].BlogBannerImage && parse.JSON(selectedData[0].BlogBannerImage),
            topic: selectedData[0].Title,
            //category: selectedData[0].Category,
            entity: selectedData[0].Entity,

            description: selectedData[0].Description,
            overview: selectedData[0].Overview,

        }));
        const initialContent = getDescription(selectedData[0].description);
        let banneimagearr = []
        setRichTextValues((prevValues) => ({
            ...prevValues,
            description: initialContent,
        }));
        setImagepostIdsArr(selectedData[0]?.BlogGalleryId)
        setDocumentpostIdsArr(selectedData[0]?.BlogsDocsId)
        setImagepostArr1(selectedData[0].BlogGalleryJSON)
        setImagepostArr(selectedData[0].BlogGalleryJSON)
        setDocumentpostArr1(selectedData[0].BlogDocsJSON)
        if (selectedData[0].BannerImage.length > 0 && selectedData[0].BannerImage != null) {
            banneimagearr = selectedData[0].BannerImage
            // //console.log(banneimagearr, 'banneimagearr');
            setBannerImagepostArr(banneimagearr);
            setBannerImagepostArrEdit(banneimagearr);
            setImagepostArr1Edit(selectedData[0].BlogGalleryJSON)
            //setFormData(arr)
        }
        // //console.log("editformdata", EditFormData, selectedData)
        //window.location.href = `${SiteUrl}/SitePages/DiscussionForumDetail.aspx?${id}`;
    };

    const onFileChange = async (
        event,
        libraryName,
        docLib
    ) => {
        //debugger;
        // //console.log("libraryName-->>>>", libraryName)
        event.preventDefault();
        let uloadDocsFiles = [];
        let uloadDocsFiles1 = [];

        let uloadImageFiles = [];
        let uloadImageFiles1 = [];

        let uloadBannerImageFiles = [];

        if (event.target.files && event.target.files.length > 0) {
            const files = Array.from(event.target.files);
            event.target.value = '';
            if (libraryName === "Gallery") {
                setGalImageAdded(true);
            }
            if (libraryName === "bannerimg") {
                setBannerAdded(true);
            }
           
            if (libraryName === "Gallery" || libraryName === "bannerimg") {
                var imageVideoFiles =[];
                if(libraryName === "Gallery"){
                    imageVideoFiles = files.filter(
                        (file) =>
                            file.type.startsWith("image/") || file.type.startsWith("video/")
                    );
                }
                else if(libraryName === "bannerimg"){
                    imageVideoFiles = files.filter(
                        (file) =>
                            file.type.startsWith("image/")
                    );
                }
               

                if (imageVideoFiles.length > 0) {
                    const arr = {
                        files: imageVideoFiles,
                        libraryName: libraryName,
                        docLib: docLib,
                        name:imageVideoFiles[0].name,
                        fileUrl:  URL.createObjectURL(imageVideoFiles[0])
                    };
                    // //console.log("arr-->>>", arr)
                    if (libraryName === "Gallery") {
                        uloadImageFiles.push(arr);
                        setImagepostArr(uloadImageFiles);
                        if (ImagepostArr1.length > 0) {
                            imageVideoFiles.forEach((ele) => {
                                // //console.log("ele in if-->>>>", ele)
                                let arr1 = {
                                    ID: 0,
                                    Createdby: "",
                                    Modified: "",
                                    fileUrl: URL.createObjectURL(ele),
                                    fileSize: ele.size,
                                    fileType: ele.type,
                                    fileName: ele.name,
                                };
                                ImagepostArr1.push(arr1);
                            });
                            setImagepostArr1(ImagepostArr1);
                        } else {
                            imageVideoFiles.forEach((ele) => {
                                // //console.log("ele in else-->>>>", ele)
                                let arr1 = {
                                    ID: 0,
                                    Createdby: "",
                                    Modified: "",
                                    fileUrl: URL.createObjectURL(ele),
                                    fileSize: ele.size,
                                    fileType: ele.type,
                                    fileName: ele.name,
                                };
                                uloadImageFiles1.push(arr1);
                            });
                            setImagepostArr1(uloadImageFiles1);
                        }
                    } else if (libraryName === "bannerimg") {
                        uloadBannerImageFiles.push(arr);
                        // //console.log("uloadBannerImageFiles-->>", uloadBannerImageFiles)
                        setBannerImagepostArr(uloadBannerImageFiles);
                    }
                } else {
                    if(libraryName === "bannerimg"){
                                Swal.fire("only image can be upload");
                              }else{
                                Swal.fire("only image & video can be upload");
                              }
                }
            }
        }
    };

    const validateForm = () => {
        const { topic, Type, category, entity, overview, FeaturedAnnouncement } =
            formData;
        const { description } = richTextValues;
        let valid = true;

        if (!topic) {
            Swal.fire("Error", "Topic is required!", "error");
            valid = false;
        } else if (!entity) {
            Swal.fire("Error", "Entity is required!", "error");
            valid = false;
        }
        else if (BnnerImagepostArr.length == 0 && BnnerImagepostArrEdit.length == 0) {
            Swal.fire('Error', 'Banner image is required!', 'error');
            valid = false;
        }
        else if (ImagepostArr.length == 0 && ImagepostArr1Edit.length == 0) {
            Swal.fire('Error', 'Gallery image is required!', 'error');
            valid = false;
        }
       

        return valid;
    };
    const validateFormDraft = () => {
        const { topic, Type, category, entity, overview, FeaturedAnnouncement } =
            formData;
        const { description } = richTextValues;
        let valid = true;

        if (!topic) {
            Swal.fire("Error", "Topic is required!", "error");
            valid = false;
        } else if (!entity) {
            Swal.fire("Error", "Entity is required!", "error");
            valid = false;
        }
        else if (BnnerImagepostArr.length == 0 && BnnerImagepostArrEdit.length == 0) {
            Swal.fire('Error', 'Banner image is required!', 'error');
            valid = false;
        }
        // else if (ImagepostArr.length == 0 && ImagepostArr1Edit.length == 0) {
        //     Swal.fire('Error', 'Gallery image is required!', 'error');
        //     valid = false;
        // }
        //else if (!overview) {
        //   Swal.fire('Error', 'Overview is required!', 'error');
        //   valid = false;
        // } else if (!description) {
        //   Swal.fire('Error', 'Description is required!', 'error');
        //   valid = false;
        // } else if (!FeaturedAnnouncement) {
        //   Swal.fire('Error', 'Featured Announcement is required!', 'error');
        //   valid = false;
        // }

        return valid;
    };
    const handleFormSubmit = async () => {
        if (validateForm()) {
            if (editForm) {
                Swal.fire({
                    // title: "Are you sure you want to publish this blog?",
                    title: "Are you sure you want to submit this blog?",
                    showConfirmButton: true,
                    showCancelButton: true,
                    confirmButtonText: "Yes",
                    cancelButtonText: "No",
                    icon: "warning",
                }).then(async (result) => {
                    // //console.log(result);
                    if (result.isConfirmed) {
                        ////console.log("Form Submitted:", formValues, bannerImages, galleryImages, documents);
                        // //debugger;
                        setLoading(true);
                        const modalBackdrop = document.querySelector('.modal-backdrop');
                        if (modalBackdrop) {
                            modalBackdrop.classList.remove('modal-backdrop');
                            modalBackdrop.classList.remove('fade');
                            modalBackdrop.classList.remove('show');
                            // modalBackdrop.remove();
                        }
                        let bannerImageArray = {};
                        let galleryIds = [];
                        let documentIds = [];
                        let galleryArray = [];
                        let documentArray = [];

                        if (
                            BnnerImagepostArr.length > 0 &&
                            BnnerImagepostArr[0]?.files?.length > 0 && IsBannerAdded
                        ) {
                            //debugger
                            for (const file of BnnerImagepostArr[0].files) {
                                //debugger
                                //  const uploadedBanner = await uploadFile(file, _sp, "Documents", Url);
                                bannerImageArray = await uploadFileBanner(
                                    file,
                                    _sp,
                                    "Documents",
                                    "https://officeindia.sharepoint.com"
                                );
                            }
                        } else {
                            bannerImageArray = null;
                        }
                        // //debugger;
                        // //console.log("bannerImageArraybannerImageArray on submit", bannerImageArray,
                        //     BnnerImagepostArrEdit,
                        //     BnnerImagepostArrEdit.length
                        // )

                        if (bannerImageArray != null || BnnerImagepostArrEdit != null) {
                            // Create Post
                            const postPayload = IsBannerAdded ?
                                {
                                    Title: formData.topic,
                                    Overview: formData.overview,
                                    Description: richTextValues.description,
                                    EntityId: formData.entity && Number(formData.entity),
                                    // Status: "Published",
                                    Status: "Submitted",
                                    BlogBannerImage: bannerImageArray && JSON.stringify(bannerImageArray)
                                    // DiscussionForumCategoryId: Number(formData.category),
                                }
                                :
                                {
                                    Title: formData.topic,
                                    Overview: formData.overview,
                                    Description: richTextValues.description,
                                    EntityId: formData.entity && Number(formData.entity),
                                    // Status: "Published",
                                    Status: "Submitted",
                                    //BlogBannerImage: bannerImageArray && JSON.stringify(bannerImageArray)
                                    // DiscussionForumCategoryId: Number(formData.category),
                                };
                            // //console.log("postPayload-->>>>>", postPayload);

                            const postResult = await updateItem(postPayload, _sp, editID);
                            const postId = postResult?.data?.ID;
                            // //debugger;
                         
                            if (ImagepostArr.length > 0 && IsGalImageAdded) {
                                for (const file of ImagepostArr[0]?.files) {
                                    const uploadedGalleryImage = await uploadFileToLibrary(
                                        file,
                                        _sp,
                                        "BlogGallery"
                                    );

                                    galleryIds = galleryIds.concat(
                                        uploadedGalleryImage.map((item) => item.ID)
                                    );
                                    galleryArray.push(uploadedGalleryImage);
                                    // //console.log("uploadedGalleryImage draft", uploadedGalleryImage, galleryIds)
                                }
                            }

                            // //console.log("IsGalImageAdded draft", IsGalImageAdded)
                            if (IsGalImageAdded) {
                                const updatePayload = {
                                    ...(galleryIds.length > 0 && {
                                        BlogsGalleryId: galleryIds,
                                        BlogGalleryJSON: JSON.stringify(
                                            flatArray(galleryArray)
                                        ),
                                    }),

                                };

                                if (Object.keys(updatePayload).length > 0) {
                                    const updateResult = await updateItem(updatePayload, _sp, editID);
                                    // //console.log("Update Result:", updateResult);
                                }
                            }
                           
                        }

                       
                                 // //////######### changes ############  ////////////
                       
                                          let arr = {
                                              Title:formData.topic,
                                               ContentID: editID,
                                 
                                               ContentName: "ARGBlogs",
                                 
                                               Status: "Pending",
                                 
                                               EntityId: Number(formData.entity),
                                 
                                               SourceName: "Blogs",
                                               ReworkRequestedBy: "Initiator"
                                 
                                 
                                             }
                                 
                                            let boolval = false;      

                        if (ApprovalRequestItem && ApprovalRequestItem.IsRework && ApprovalRequestItem.IsRework == 'Yes') {
                            const ctmasteritm = await _sp.web.lists.getByTitle(LIST_TITLE_ContentMaster).items.filter('ContentID eq ' + ApprovalRequestItem.ContentId + " and SourceName eq '" + CONTENTTYPE_Blogs + "'")();
                            if (ctmasteritm && ctmasteritm.length > 0) {
                                let updaterec = { 'Status': 'Pending', 'ReworkRequestedBy': 'Initiator' }
                                if (ApprovalRequestItem.LevelSequence == 1) updaterec.ReworkRequestedBy = "Level 1";
                                await UpdateContentMaster(_sp, ctmasteritm[0].Id, updaterec);
                                await _sp.web.lists.getByTitle(LIST_TITLE_MyRequest).items.getById(ApprovalRequestItem.Id).update({ 'Status': 'Submitted' });
                                await _sp.web.lists.getByTitle(LIST_TITLE_Blogs).items.getById(editID).update({ 'Status': 'Submitted' });
                                boolval = true;
                            }
                        }
                        else {
                             await AddContentMaster(_sp, arr)
                           boolval = await handleClick(editID, "Blogs", Number(formData.entity))
                           
                            // boolval = true;
                        }
                       
                                    // //////######### changes ############  ////////////
                       
                        // Swal.fire("Item updated successfully", "", "success");
                        dismissModal();
                        if (boolval == true) {
                            Swal.fire("Item submitted successfully", "", "success");
                        setTimeout(async () => {
                           
                            setBlogData(await fetchBlogdata(_sp));
                           
                            // dismissModal();
                            ApIcall();
                            setLoading(false);
                            const modalBackdrop = document.querySelector('.modal-backdrop');
                        if (modalBackdrop) {
                            modalBackdrop.classList.remove('modal-backdrop');
                            modalBackdrop.classList.remove('fade');
                            modalBackdrop.classList.remove('show');
                            // modalBackdrop.remove();
                        }
                            // window.location.href = `${SiteUrl}/SitePages/Blogs.aspx`;
                        }, 100);
                    }
                       
                    }
                });
            }
        }
    };
    const flatArray = (arr) => {
        return arr.reduce((acc, val) => acc.concat(val), []);
    };
    const handleFormSaveasDraft = async () => {
        if (validateFormDraft()) {
            Swal.fire({
                title: "Do you want to save this blog?",
                showConfirmButton: true,
                showCancelButton: true,
                confirmButtonText: "Yes",
                cancelButtonText: "No",
                icon: "warning",
            }).then(async (result) => {
                ////console.log("Form Submitted:", formValues, bannerImages, galleryImages, documents);
                if (result.isConfirmed) {
                    // //debugger;
                    setLoading(true);
                    const modalBackdrop = document.querySelector('.modal-backdrop');
                    if (modalBackdrop) {
                        modalBackdrop.classList.remove('modal-backdrop');
                        modalBackdrop.classList.remove('fade');
                        modalBackdrop.classList.remove('show');
                        // modalBackdrop.remove();
                    }
                    let bannerImageArray = {};
                    let galleryIds = [];
                    let documentIds = [];
                    let galleryArray = [];
                    let documentArray = [];

                    // formData.FeaturedAnnouncement === "on"?  true :false;

                    // Upload Banner Images
                    // //console.log("BnnerImagepostArrBnnerImagepostArr draft", BnnerImagepostArr)
                    if (
                        BnnerImagepostArr.length > 0 &&
                        BnnerImagepostArr[0]?.files?.length > 0 && IsBannerAdded
                    ) {
                        for (const file of BnnerImagepostArr[0].files) {
                            //  const uploadedBanner = await uploadFile(file, _sp, "Documents", Url);
                            bannerImageArray = await uploadFileBanner(
                                file,
                                _sp,
                                "Documents",
                                "https://officeindia.sharepoint.com"
                            );
                        }
                    }
                    // //debugger;
                    // Create Post
                    // //console.log("BnnerImagepostArrBnnerImagepostArr draft", bannerImageArray)
                    const postPayload = IsBannerAdded ?
                        {
                            Title: formData.topic,
                            Overview: formData.overview,
                            Description: richTextValues.description,
                            EntityId: formData.entity && Number(formData.entity),
                            Status: "Save as Draft",
                            BlogBannerImage: bannerImageArray && JSON.stringify(bannerImageArray)
                        }
                        :
                        {
                            Title: formData.topic,
                            Overview: formData.overview,
                            Description: richTextValues.description,
                            EntityId: formData.entity && Number(formData.entity),
                            Status: "Save as Draft",
                        };
                    //console.log("postPayload 3-->>>>>", postPayload);

                    const postResult = await updateItem(postPayload, _sp, editID);
                    const postId = postResult?.data?.ID;
                    // //debugger;
                    //console.log("postID", postId, postResult)
                    //console.log("ImagepostArrImagepostArr draft", ImagepostArr, IsGalImageAdded)
                    if (ImagepostArr.length > 0 && IsGalImageAdded) {
                        for (const file of ImagepostArr[0]?.files) {
                            const uploadedGalleryImage = await uploadFileToLibrary(
                                file,
                                _sp,
                                "BlogGallery"
                            );

                            galleryIds = galleryIds.concat(
                                uploadedGalleryImage.map((item) => item.ID)
                            );
                            galleryArray.push(uploadedGalleryImage);
                            //console.log("uploadedGalleryImage draft", uploadedGalleryImage, galleryIds)
                        }
                    }

                    //console.log("IsGalImageAdded draft", IsGalImageAdded)
                    if (IsGalImageAdded) {
                        const updatePayload = {
                            ...(galleryIds.length > 0 && {
                                BlogsGalleryId: galleryIds,
                                BlogGalleryJSON: JSON.stringify(
                                    flatArray(galleryArray)
                                ),
                            }),

                        };

                        if (Object.keys(updatePayload).length > 0) {
                            const updateResult = await updateItem(updatePayload, _sp, editID);
                            //console.log("Update Result:", updateResult);
                        }
                    }
                    dismissModal();
                    Swal.fire("Item saved successfully", "", "success");
                    setTimeout(async () => {
                       
                        setBlogData(await fetchBlogdata(_sp));
                       
                       
                        ApIcall();
                        setLoading(false);
                       
                        // window.location.href = `${SiteUrl}/SitePages/Blogs.aspx`;

                    }, 100);
                }
            });
        }

    };

    const dismissModal = () => {
        const modalElement = document.getElementById('discussionModalEdit');

        // Remove Bootstrap classes and attributes manually
        modalElement.classList.remove('show');
        // modalElement.style.display = 'none';
        modalElement.setAttribute('aria-hidden', 'true');
        modalElement.removeAttribute('aria-modal');
        modalElement.removeAttribute('role');

        // Optionally, remove the backdrop if it was added manually
        const modalBackdrop = document.querySelector('.modal-backdrop');
        if (modalBackdrop) {
            modalBackdrop.classList.remove('modal-backdrop');
            modalBackdrop.classList.remove('fade');
            modalBackdrop.classList.remove('show');
            // modalBackdrop.remove();
        }
    };
    const onChange = async (name, value) => {
        setFormData((prevData) => ({
            ...prevData,
            [name]: value,
        }));

        // if (name == "Type") {
        //   setCategoryData(await getCategory(_sp, Number(value))); // Category
        // }
         if(name == "entity"){
       
               //ARGApprovalConfiguration
             
                    const rowData = await getApprovalConfiguration(_sp, Number(value)) //baseUrl
             
                    const initialRows = rowData.map((item) => ({
             
                      id: item.Id,
             
                      Level: item.Level.Level,
             
                      LevelId: item.LevelId,
             
                      approvedUserListupdate: item.Users.map((user) => ({
             
                        id: user.ID,
             
                        name: user.Title,
             
                        email: user.EMail
             
                      })),
             
                      selectionType: 'All' // default selection type, if any
             
                    }));
             
                    setRows(initialRows);
       
            }
    };


    // Edit Blogs End
    const DeleteBlog = (id) => {
        Swal.fire({
            title: "Are you sure?",
            text: "You won't be able to revert this!",
            icon: "warning",
            showCancelButton: true,
            confirmButtonColor: "#3085d6",
            cancelButtonColor: "#d33",
            confirmButtonText: "Yes, delete it!"
        }).then((result) => {
            if (result.isConfirmed) {
                const DeleteRes = DeleteBusinessAppsAPI(_sp, id)
                setLoading(true);
                ApIcall();
                Swal.fire({
                    title: "Deleted!",
                    text: "Item has been deleted.",
                    icon: "success"
                }).then(async res => {
                    setBlogData(await fetchBlogdata(_sp));
                    setLoading(false);
                }
                ).catch(async err => {
                    setBlogData(await fetchBlogdata(_sp));
                    setLoading(false);
                }
                )

            }
        })
    }

    //new code end

    const toggleDropdown = (itemId) => {
        if (showDropdownId === itemId) {
            setShowDropdownId(null); // Close the dropdown if already open
        } else {
            setShowDropdownId(itemId); // Open the dropdown for the clicked item
        }
    };
    const togglePin = async (e, item) => {

        //debugger

        e.preventDefault();

        //setLoadingUsers((prev) => ({ ...prev, [item.ID]: true })); // Set loading state for the specific user


        try {

            setLoading(true);

            const currentUser = await _sp.web.currentUser();

            const saveRecords = await _sp.web.lists.getByTitle("ARGSavedBlogs").items
                .select("*,BlogId/Id")
                .expand("BlogId")
                .filter(`BlogSavedById eq ${currentUser.Id} and BlogId/Id eq ${item}`)();


            if (saveRecords.length > 0) {

                // Unpin logic

                await _sp.web.lists.getByTitle("ARGSavedBlogs").items.getById(saveRecords[0].Id).delete();

                setBookmarkstatus((prev) => ({ ...prev, [item.ID]: false })); // Update [pin] status
                // setLoading(false);


            } else {

                // pin logic

                await _sp.web.lists.getByTitle("ARGSavedBlogs").items.add({

                    BlogSavedById: currentUser.Id,

                    BlogIdId: item

                }).then(async (ress) => {

                    //console.log(ress);
                    // setLoading(false);
                });

                setBookmarkstatus((prev) => ({ ...prev, [item.ID]: true })); // Update pin status
                Swal.fire({
                    title: 'Successful',
                    text: 'Bookmarks added successfully!',
                    icon: 'success',
                    confirmButtonText: 'OK'
                });
               

            }

        } catch (error) {
            setLoading(false);
            //setLoadingUsers((prev) => ({ ...prev, [item.ID]: false })); // End loading state for the specific user


            console.error("Error toggling pin status:", error);

            alert("Failed to toggle pin status. Please try again.");

        } finally {
            setLoading(false);
            ApIcall();
            //  fetchUserInformationList()

            //   setLoadingUsers((prev) => ({ ...prev, [item.ID]: false })); // End loading state for the specific user



        }

    };
    const sendanEmail = (item) => {
        // window.open("https://outlook.office.com/mail/inbox");

        // const subject = "Blog Title -" + item.Title;
        // const body = 'Here is the link to the Blog:' + `${SiteUrl}/SitePages/BlogDetails.aspx?${item.Id}`;
        const subject = "Thought You’d Find This Interesting!";
        const body = 'Hi,' +
            'I came across something that might interest you: ' +
            `<a href="${siteUrl}/SitePages/BlogDetails.aspx?${item.Id}"></a>`
        'Let me know what you think!';

        // window.open(`https://outlook.office365.com/mail/deeplink/compose?body= Here is the group link: "${siteUrl}/SitePages/GroupandTeamDetails.aspx?${idNum}"`);
        const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${subject}&body=${body}`;
        //const mailtoLink = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

        // Open the link to launch the default mail client (like Outlook)
        //window.location.href = mailtoLink;
        //const office365MailLink = `https://outlook.office.com/mail/deeplink/compose?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

        window.open(office365MailLink, '_blank');
    };
    const loadMore = () => {
        event.preventDefault()
        event.stopImmediatePropagation()
        setItemsToShow(itemsToShow + 5); // Increase the number by 8
    };
   
   
    return (
        <><div className="row mt-2">
            {blogData.length > 0 ?
                // blogData.filter(x => x.Status == "Published").slice(0, 1).map(item => {
                    blogData.filter(x => x.Status == "Approved").slice(0, 1).map(item => {

                    const AnnouncementandNewsBannerImage = item.BlogBannerImage == undefined || item.BlogBannerImage == null ? ""
                        : JSON.parse(item.BlogBannerImage);
                    return (
                        <>
                            <div className="col-lg-5" onClick={() => gotoBlogsDetails(item)}>
                                <div className="imagemani mt- mb-3 me-2">
                                    <img src={AnnouncementandNewsBannerImage?.serverUrl + AnnouncementandNewsBannerImage?.serverRelativeUrl}
                                        className="d-flex align-self-center me-3 w-100" lt="Generic placeholder image" style={{ objectFit: 'cover' }} />
                                </div>
                            </div>
                            <div className="col-lg-7">
                                <div className="row" style={{ paddingLeft: '0.5rem' }}>
                                    <div className="col-sm-4 text-left">
                                        <span style={{ padding: '5px', borderRadius: '4px', fontWeight: '500', color: '#009157', background: 'rgba(0, 135, 81, 0.20)' }} className=" font-14 float-start mt-2">
                                            Latest Blog</span>

                                    </div>
                                    <div className="col-lg-12">
                                        <h4 style={{ lineHeight: '34px' }} className="page-title fw-700 mb-1  pe-5 font-28 titleHeading">
                                            {truncateText(item.Title, 120)}</h4>
                                    </div>
                                    <div className="row">
                                        <div className="col-sm-12">
                                            <p className="mb-2 mt-1 text-dark d-block customhead">
                                                <span style={{ fontWeight: '400' }} className="pe-2 text-nowrap color-new font-12 mb-0 d-inline-block">
                                                    <Calendar size={12}  strokeWidth={1} className="pl-2" style={{ fontWeight: '400' }} />
                                                    {moment(item.Created).format("DD-MMM-YYYY")} &nbsp;  &nbsp;  &nbsp;|
                                                </span>
                                                <span style={{ fontWeight: '400' }} className="text-nowrap mb-0 color-new font-12 d-inline-block">
                                                    Author: <span style={{ color: '#009157', fontWeight: '600' }}>{item.Author.Title} &nbsp;  &nbsp;  &nbsp;
                                                    </span>

                                                </span></p>

                                            <div style={{ clear: 'both' }}>
                                                <p style={{fontSize:'15px', lineHeight:'22px'}} className="d-block text-dark cursor-none customdescription">{truncateText(item.Overview, 300)}</p>
                                            </div>
                                        </div>
                                        <a onClick={() => gotoBlogsDetails(item)} style={{ textDecoration: 'none' }}>
                                            <div style={{ height: '40px', lineHeight: '24px' }} className="btnCustomcss btn-primary">Read more..</div> </a>
                                    </div>
                                </div>

                            </div></>)
                }) : null}
        </div>
            <div className="row mt-0">
                <div className="col-12">
                    <div className="card mb-0 cardcsss">
                        <div className="card-body">
                            <div className="d-flex flex-wrap align-items-center justify-content-center">
                                <ul
                                    className="navs nav-pillss navtab-bgs justify-content-center"
                                    role="tablist"
                                    style={{
                                        
                                        display: "flex",
                                        listStyle: "none",
                                        marginBottom: "unset",
                                    }}
                                >
                                    {/* {//console.log("FilterOptions", FilterOptions)} */}
                                    {FilterOptions.length > 0 && FilterOptions.map((res) => (
                                        <li className="nav-itemcss">
                                            <a
                                                className={`nav-linkss ${activeTab.toLowerCase() ===
                                                    (res.Name.toLowerCase()
                                                        ? res.Name.toLowerCase()
                                                        : "")
                                                    ? "active"
                                                    : ""
                                                    }`}

                                                //aria-selected={activeTab === "cardView"}
                                                role="tab"
                                                onClick={() => handleTabClick(res.Name.toLowerCase(), res.ID)}
                                            >
                                                {res.name}
                                            </a>
                                        </li>
                                    ))}
                                </ul>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            {Loading ?
                   
                   <div style={{minHeight:'100vh',marginTop:'100px'}} className="loadernewadd mt-10">
                   <div>
                       <img
                           src={require("../../CustomAsset/birdloader.gif")}
                           className="alignrightl"
                           alt="Loading..."
                         />
                       </div>
                     <span>Loading </span>{" "}
                     <span>
                       <img
                         src={require("../../CustomAsset/argloader.gif")}
                         className="alignrightl"
                         alt="Loading..."
                       />
                     </span>
                   </div>
                 :
            <div className="tab-content mt-2">
                <div className="tab-pane show active" id="home1" role="tabpanel">
                    {filteredBlogItems.length > 0 ?  
                    filteredBlogItems.slice(0, visibleItems).map(item => {
                       
                        // filteredBlogItems.map(item => {
                            const AnnouncementandNewsBannerImage = item.BlogBannerImage == undefined || item.BlogBannerImage == null ? ""
                                : JSON.parse(item.BlogBannerImage);
                            // console.log("AnnouncementandNewsBannerImage", AnnouncementandNewsBannerImage, item);
                            const imageData = item.BlogBannerImage ? JSON.parse(item.BlogBannerImage) : null;
                            let siteIdAl = '338f2337-8cbb-4cd1-bed1-593e9336cd0e,e2837b3f-b207-41eb-940b-71c74da3d214';
                            let listIDAl = '33ff70ea-7112-4dfb-a0fe-477786aee9e3';

                            let siteId = SiteUrl.toLowerCase().includes('alrostmani') ? siteIdAl : '02993535-33e8-44d1-9edf-0d484e642ea1,d9374a3d-ae79-4d2a-8d36-d48f86e3201e';
                            let listID = SiteUrl.toLowerCase().includes('alrostmani') ? listIDAl : '1e0eead0-ed78-4b4d-92dc-4cd058deac82';

                            let img1 = imageData && imageData.fileName ? `${SiteUrl}/_api/v2.1/sites('${siteId}')/lists('${listID}')/items('${item.ID}')/attachments('${imageData.fileName}')/thumbnails/0/c400x400/content?prefer=noredirect%2Cclosestavailablesize` : ""
                            let img = imageData && imageData.serverRelativeUrl ? `https://officeindia.sharepoint.com${imageData.serverRelativeUrl}` : img1
                            const imageUrl = imageData
                                ? img
                                : require("../../webparts/businessApps/assets/userimg.png");
                            // { console.log("imageData", imageData, imageUrl, item, SiteUrl, img) }

                            return (
                                <div className="card mb-2 annuncementcard">
                                    <div className="card-body">
                                        <div className="row align-items-start">
                                            <div className="col-sm-2">
                                                <a onClick={() => gotoBlogsDetails(item)}>   <div className="imagehright">
                                                    {/* <img className="d-flex align-self-center me-3 w-100" src={g1} alt="Generic placeholder image" /> */}
                                                    <img src={imageUrl}
                                                        className="d-flex align-self-center me-3 w-100" lt="Generic placeholder image" style={{ height: '100%' }} />
                                                </div>
                                                </a>
                                            </div>
                                            <div className="col-sm-9 padd-12">
                                                <div className="row">
                                                    <div className="col-sm-4 date-color">
                                                        <span className="font-12 date-color float-start mt-0 mb-2 ng-binding" style={{ color: '#6b6b6b', fontSize: '12px', paddingRight: '0.2rem' }}>
                                                            <Calendar size={12} color="#6b6b6b" strokeWidth={2} style={{ fontWeight: '400' }} /></span>

                                                        <span className="font-12 date-color float-start mt-0 mb-2 ng-binding" style={{ color: '#6b6b6b', fontSize: '12px' }}>{moment(item.Created).format("DD-MMM-YYYY")}
                                                            {/* 12-Mar-2024 18:37 */}
                                                        </span><br></br>
                                                    </div>
                                                    <h6  className="font-12 date-color float-start mt-0 mb-2 ng-binding" style={{ color: '#6b6b6b', fontSize: '12px' }}> {item.Author.Title}  </h6>

                                                </div>
                                                <a> <div className="w-100">
                                                    <h4 className="mt-0 mb-1 font-16 fw-bold ng-binding" style={{ color: '#343a40', fontSize: '16px' }}> {truncateText(item.Title, 90)}
                                                    </h4>
                                                    <p style={{ color: '#6b6b6b', fontSize: '15px', lineHeight:'20px' }} className="mb-2 ng-binding">
                                                        {truncateText(item.Overview, 200)}</p>
                                                    <div className="readmore" onClick={() => gotoBlogsDetails(item)} style={{ cursor: 'pointer' }}>Read more..                                                    {item.Status == "Submitted" &&<span style={{marginLeft:'82%'}}>Waiting for approval</span>}
                                                    </div>
                                                </div>
                                                </a>
                                            </div>
                                            <div className="col-sm-1 posx">
                                                <div className="d-flex" style={{ justifyContent: 'end', gap: '10px', marginRight: '3px', cursor: 'pointer' }}>
                                                    <div className="col-lg-12">
                                                        {(item.Status == "Save as Draft" || item.WillReworkEdit == true) &&
                                                            <><div style={{gap:'10px'}} className="d-flex flex-wrap align-items-center justify-content-end mt-0">
                                                                {/* Button to trigger modal */}
                                                                <FontAwesomeIcon
                                                                    icon={faEdit}
                                                                    data-bs-toggle="modal"
                                                                    data-bs-target="#discussionModalEdit"
                                                                    className="text-dark1 waves-effect waves-light"
                                                                    onClick={() => EditBlogs(item.ID)} />


                                                           

                                                                  {item.Status == "Save as Draft" && <a
                                                                       
                                                                        onClick={() => DeleteBlog(item.ID)}
                                                                    >
                                                                        <FontAwesomeIcon  className="text-dark1 waves-effect waves-light" icon={faTrashAlt} />
                                                                    </a>}

                                                                </div></>
                                                        }
                                                        {/* Bootstrap Modal */}
                                                        {/* {//console.log("formdata", formData)} */}
                                                        <div
                                                            className="modal fade bd-example-modal-lg"
                                                            id="discussionModalEdit"
                                                            tabIndex={-1}
                                                            aria-labelledby="exampleModalLabel"
                                                            aria-hidden="true"
                                                            data-target=".bd-example-modal-lg"
                                                        >
                                                            <div style={{ minWidth: '80%' }} className="modal-dialog modal-lg ">
                                                                <div className="modal-content">
                                                                    <div className="modal-header d-block">
                                                                        <h5 className="modal-title" id="exampleModalLabel">
                                                                            Edit Blog
                                                                        </h5>
                                                                        <button style={{ right: '20px', top: '20px' }}
                                                                            type="button"
                                                                            className="btn-close"
                                                                            data-bs-dismiss="modal"
                                                                            aria-label="Close"
                                                                        ></button>
                                                                    </div>
                                                                   
                                                                    <div className="modal-body">
                                                                        <form className="row">
                                                                            <div className="col-lg-3">
                                                                                <div className="mb-3">
                                                                                    <label htmlFor="topic" className="form-label">
                                                                                        Topic <span className="text-danger">*</span>
                                                                                    </label>
                                                                                    <input
                                                                                        type="text"
                                                                                        id="topic"
                                                                                        name="topic"
                                                                                        placeholder="Enter Topic"
                                                                                        className="form-control"
                                                                                        value={formData.topic}
                                                                                        onChange={(e) =>
                                                                                            onChange(e.target.name, e.target.value)
                                                                                        }
                                                                                    />
                                                                                </div>
                                                                            </div>

                                                                            <div className="col-lg-3">
                                                                                <div className="mb-3">
                                                                                    <label
                                                                                        htmlFor="entity"
                                                                                        className="form-label"
                                                                                    >
                                                                                        Entity{" "}
                                                                                        <span className="text-danger">*</span>
                                                                                    </label>
                                                                                    <select
                                                                                        className="form-select"
                                                                                        id="entity"
                                                                                        name="entity"
                                                                                        value={formData.entity}
                                                                                        onChange={(e) =>
                                                                                            onChange(e.target.name, e.target.value)
                                                                                        }
                                                                                    >
                                                                                        <option value="">Select</option>
                                                                                        {EnityData.map((item, index) => (
                                                                                            <option key={index} value={item.id}>
                                                                                                {item.name}
                                                                                            </option>
                                                                                        ))}
                                                                                    </select>
                                                                                </div>
                                                                            </div>

                                                                            <div className="col-lg-3">
                                                                                <div className="mb-3">
                                                                                    <div className="d-flex justify-content-between">
                                                                                        <div>
                                                                                            <label
                                                                                                htmlFor="bannerImage"
                                                                                                className="form-label"
                                                                                            >
                                                                                                Banner Image{" "}
                                                                                                <span className="text-danger">*</span>
                                                                                            </label>
                                                                                        </div>
                                                                                        <div>
                                                                                            <div>
                                                                                                {
                                                                                                    BnnerImagepostArr[0] != false &&
                                                                                                        BnnerImagepostArr.length > 0 &&
                                                                                                        BnnerImagepostArr != undefined
                                                                                                        ? BnnerImagepostArr.length == 1 && (
                                                                                                            <a style={{ fontSize: "0.875rem" }} onClick={() =>
                                                                                                                setShowModalFunc(
                                                                                                                    true,
                                                                                                                    "bannerimg"
                                                                                                                )
                                                                                                            }>
                                                                                                                <FontAwesomeIcon
                                                                                                                    icon={faPaperclip}
                                                                                                                />
                                                                                                                1 file Attached
                                                                                                            </a>
                                                                                                        )
                                                                                                        : ""
                                                                                                    // || BnnerImagepostArr.length > 0 && BnnerImagepostArr[0].files.length > 1 &&
                                                                                                    // (<a onClick={() => setShowModalFunc(true, "bannerimg")} style={{ fontSize: '0.875rem' }}>
                                                                                                    //   <FontAwesomeIcon icon={faPaperclip} /> {BnnerImagepostArr[0].files.length} files Attached
                                                                                                    // </a>)
                                                                                                }
                                                                                            </div>
                                                                                        </div>
                                                                                    </div>
                                                                                    <input
                                                                                        type="file"
                                                                                        id="bannerImage"
                                                                                        name="bannerImage"
                                                                                        className="form-control"
                                                                                         accept="image/*"
                                                                                        onChange={(e) =>
                                                                                            onFileChange(e, "bannerimg", "Document")
                                                                                        }
                                                                                    />
                                                                                </div>
                                                                            </div>

                                                                            <div className="col-lg-3">
                                                                                <div className="mb-3">
                                                                                    <div className="d-flex justify-content-between">
                                                                                        <div>
                                                                                            <label
                                                                                                htmlFor="discussionGallery"
                                                                                                className="form-label"
                                                                                            >
                                                                                                Blog Gallery{" "}
                                                                                                <span className="text-danger">*</span>
                                                                                            </label>
                                                                                        </div>
                                                                                        <div>
                                                                                            {(ImagepostArr1.length > 0 &&
                                                                                                ImagepostArr1.length == 1 && (
                                                                                                    <a
                                                                                                        onClick={() =>
                                                                                                            setShowModalFunc(true, "Gallery")
                                                                                                        }
                                                                                                        style={{ fontSize: "0.875rem" }}
                                                                                                    >
                                                                                                        <FontAwesomeIcon
                                                                                                            icon={faPaperclip}
                                                                                                        />{" "}
                                                                                                        {ImagepostArr1.length} file Attached
                                                                                                    </a>
                                                                                                )) ||
                                                                                                (ImagepostArr1.length > 0 &&
                                                                                                    ImagepostArr1.length > 1 && (
                                                                                                        <a
                                                                                                            onClick={() =>
                                                                                                                setShowModalFunc(
                                                                                                                    true,
                                                                                                                    "Gallery"
                                                                                                                )
                                                                                                            }
                                                                                                            style={{ fontSize: "0.875rem" }}
                                                                                                        >
                                                                                                            <FontAwesomeIcon
                                                                                                                icon={faPaperclip}
                                                                                                            />{" "}
                                                                                                            {ImagepostArr1.length} files
                                                                                                            Attached
                                                                                                        </a>
                                                                                                    ))}
                                                                                        </div>
                                                                                    </div>
                                                                                    <input
                                                                                        type="file"
                                                                                        id="discussionforumGallery"
                                                                                        name="discussionforumGallery"
                                                                                        className="form-control"
                                                                                        multiple
                                                                                        onChange={(e) =>
                                                                                            onFileChange(
                                                                                                e,
                                                                                                "Gallery",
                                                                                                "DiscussionForumGallery"
                                                                                            )
                                                                                        }
                                                                                    />
                                                                                </div>
                                                                            </div>

                                                                            <div className="col-lg-12">
                                                                                <div className="mb-3">
                                                                                    <label
                                                                                        htmlFor="overview"
                                                                                        className="form-label"
                                                                                    >
                                                                                        Overview{" "}
                                                                                        <span className="text-danger">*</span>
                                                                                    </label>
                                                                                    <textarea
                                                                                        className="form-control"
                                                                                        id="overview"
                                                                                        placeholder="Enter Overview"
                                                                                        name="overview"
                                                                                        //style={{ height: "auto" }}
                                                                                        style={{ minHeight: "80px", maxHeight: "80px" }}
                                                                                        value={formData.overview}
                                                                                        onChange={(e) =>
                                                                                            onChange(e.target.name, e.target.value)
                                                                                        }
                                                                                    ></textarea>
                                                                                </div>
                                                                            </div>

                                                                            <div className="col-lg-12">
                                                                                <div className="mb-3">
                                                                                    <label
                                                                                        htmlFor="description"
                                                                                        className="form-label"
                                                                                    >
                                                                                        Description{" "}
                                                                                        <span className="text-danger">*</span>
                                                                                    </label>
                                                                                    <div
                                                                                        style={{
                                                                                            display: "contents",
                                                                                            justifyContent: "start",
                                                                                            width: "100px",
                                                                                        }}
                                                                                    >
                                                                                        <ReactQuill
                                                                                            theme="snow"
                                                                                            modules={modules}
                                                                                            formats={formats}
                                                                                            placeholder={""}
                                                                                            value={richTextValues.description}
                                                                                            onChange={(content) => {
                                                                                                setRichTextValues((prevValues) => ({
                                                                                                    ...prevValues,
                                                                                                    ["description"]: content,
                                                                                                }));
                                                                                            }}
                                                                                            style={{
                                                                                                width: "100%",
                                                                                                fontSize: "6px",
                                                                                                // height: "100px",
                                                                                            }}
                                                                                        />
                                                                                    </div>
                                                                                </div>
                                                                            </div>
                                                                         {/* /////////changes/////////// */}
                                                                            {

                                                                                rows != null && rows.length > 0 && (

                                                                                    <div className="">

                                                                                        
                                                <div style={{border:"1px solid #ccc", borderRadius:'7px'}} className="p-3">
                             
                             <div >
        
                               <strong className="font-16 mb-1">Approval Hierarchy</strong>
                               <p className="font-14 text-muted mb-3">Define Approval Hierarchy for the document submitted</p>

                                                                                            </div>

                                                                                     


                                                                                            <div className="d-flex flex-column">
                                                                                            <table className="mtbalenew ">
                                                    <thead>
                                                      <tr>
                                                      <th style={{minWidth:'60px', maxWidth:'60px'}} className="newpad">  Select Level</th>
                                                      <th className="newpad"> Select Approver</th>
                                                      </tr>
                                                    </thead>
                                                    </table>




                                                                                                {rows.map((row) => (

                                                                                                    <div className="row mb-0" key={row.id}>

<table className="mtbalenew">
                                                    <tbody>
                                                      <tr>
                                                        <td style={{minWidth:'60px', maxWidth:'60px'}}>
                                                        <select style={{border:"0px solid #ccc", background:'#fff'}}
                             
                             className="form-control removeb"
                                      id={`Level-${row.id}`}

                                                                                                                name="Level"

                                                                                                                value={row.LevelId}
                                                                                                                disabled={true}


                                                                                                                onChange={(e) => {

                                                                                                                    const selectedLevel = e.target.value;

                                                                                                                    setRows((prevRows) =>

                                                                                                                        prevRows.map((r) =>

                                                                                                                            r.id === row.id

                                                                                                                                ? { ...r, LevelId: selectedLevel }

                                                                                                                                : r

                                                                                                                        )

                                                                                                                    );

                                                                                                                }}

                                                                                                            >

                                                                                                                <option value="">Select</option>

                                                                                                                {levels.map((item) => (

                                                                                                                    <option key={item.Id} value={item.Id}>

                                                                                                                        {item.Level}

                                                                                                                    </option>

                                                                                                                ))}

                                                                                                            </select>
                                                        </td>
                                                        <td>
                                                        <Multiselect className="removeb" style={{border:"0px solid #ccc", background:'#fff'}}
                                              options={row.approvedUserList}

                                                                                                                selectedValues={row.approvedUserListupdate}

                                                                                                                onSelect={(selected) => handleUserSelect(selected, row.id)}

                                                                                                                onRemove={(selected) => handleUserSelect(selected, row.id)}

                                                                                                                displayValue="name"
                                                                                                                disable={true}
                                                                                                                placeholder=''
                                                                                                                hidePlaceholder={true}

                                                                                                            />


                                                        </td>
                                                      </tr>
                                                    </tbody>

                                                  </table>

                                                                                                        <div className="col-12 col-md-5">

                                                                                                            {/* <label htmlFor={`Level-${row.id}`} className="form-label">

                                                                                                                Select Level

                                                                                                            </label> */}

                                                                                                            

                                                                                                        </div>


                                                                                                        <div className="col-12 col-md-5">

                                                                                                            {/* <label htmlFor={`approver-${row.id}`} className="form-label">

                                                                                                                Select Approver

                                                                                                            </label> */}

                                                                                                            
                                                                                                        </div>




                                                                                                    </div>

                                                                                                ))}

                                                                                            </div>




                                                                                        </div>

                                                                                    </div>

                                                                                )



                                                                            }



                                                                                                                  {/* /////////changes/////////// */}


                                                                            <div className="text-center butncss mt-2">
                                                                                <div style={{width:'140px', justifyContent:'center', textAlign:'center'}}
                                                                                    className="btn btn-success waves-effect waves-light m-1"
                                                                                  
                                                                                    onClick={handleFormSaveasDraft}
                                                                                >
                                                                                    <div
                                                                                        className="d-flex"
                                                                                        style={{
                                                                                            justifyContent: "center",
                                                                                            
                                                                                        }}
                                                                                    >
                                                                                        <img
                                                                                            src={require("../../Assets/ExtraImage/checkcircle.svg")}
                                                                                            style={{ width: "1rem", marginRight:'3px' }}
                                                                                            alt="Check"
                                                                                        />{" "}
                                                                                        Save as Draft
                                                                                    </div>
                                                                                </div>
                                                                                <div style={{width:'140px', justifyContent:'center', textAlign:'center'}}
                                                                                    className="btn btn-success waves-effect waves-light m-1"
                                                                                  
                                                                                    onClick={handleFormSubmit}
                                                                                >
                                                                                    <div
                                                                                        className="d-flex"
                                                                                        style={{
                                                                                            justifyContent: "center",
                                                                                           
                                                                                        }}
                                                                                    >
                                                                                        <img
                                                                                            src={require("../../Assets/ExtraImage/checkcircle.svg")}
                                                                                            style={{ width: "1rem", marginRight:'3px' }}
                                                                                            alt="Check"
                                                                                        />{" "}
                                                                                        Submit
                                                                                    </div>
                                                                                </div>
                                                                                <button style={{width:"140px", justifyContent:'center'}}
                                                                                    type="button"
                                                                                    className="btn cancel-btn waves-effect waves-light m-1"
                                                                                    data-bs-dismiss="modal"
                                                                                    aria-label="Close"
                                                                                >
                                                                                    <img
                                                                                        src={require("../../Assets/ExtraImage/xIcon.svg")}
                                                                                        style={{ width: "1rem", marginRight:'3px' }}
                                                                                        className="me-1"
                                                                                        alt="x"
                                                                                    />
                                                                                    Cancel
                                                                                </button>
                                                                            </div>
                                                                        </form>
                                                                    </div>
                 
                                                                </div>
                                                            </div>

                                                           
                                                        </div>
                                                    </div>
                                                    <div className="" style={{ position: 'relative' }}>
                                                        <div className="" onClick={() => toggleDropdown(item.Id)} key={item.Id}>
                                                            <Share2 size={22} color="#6c757d" strokeWidth={2} style={{ fontWeight: '400' }} />
                                                        </div>
                                                        {showDropdownId === item.Id && (
                                                            <div className="dropdown-menu dropcss" ref={menuRef}>
                                                                <a className="dropdown-item dropcssItem" onClick={() => sendanEmail(item)}> <Share size={16} /> Share by email</a>
                                                                <a className="dropdown-item dropcssItem" onClick={() => copyToClipboard(item.Id)}>
                                                                <Link size={14} />  Copy Link
                                                                </a>
                                                                <a>{copySuccess && <span className="text-success">{copySuccess}</span>}</a>
                                                            </div>
                                                        )}
                                                    </div>
                                                    <img color="#6c757d" width={"20px"} style={{ cursor: "pointer", height: '22px', fontWeight: '400' }}

                                                        //src={require("../../CustomAsset/unbookmark.png")}
                                                        src={Bookmarkstatus[item.ID] ? require("../../CustomAsset/bookmark.png") : require("../../CustomAsset/unbookmark.png")}

                                                        className="alignrightpin"
                                                        alt="pin"
                                                        onClick={(e) => togglePin(e, item.Id)}
                                                    />
                                                    {/* <Bookmark size={20} color="#6c757d" strokeWidth={2} style={{ fontWeight: '400' }} /> */}
                                                </div>
                                            </div>

                                        </div>
                                    </div>
                                </div>
                            )
                        }
                        ) : <div className='mt-2'>
                        <div style={{ height: '300px' }} className="card card-body align-items-center  annusvg text-center"
                        >

                            <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 8A6 6 0 0 0 6 8c0 7-3 9-3 9h18s-3-2-3-9"></path><path d="M13.73 21a2 2 0 0 1-3.46 0"></path></svg>

                            <p className="font-14 text-muted text-center">No Blog Available </p>

                        </div>
                    </div>

                    }

                     {visibleItems < filteredBlogItems.length && (
                        <div className="text-center">
                            <button onClick={() => handleLoadMore()} className="btn btn-primary mt-3">Load More</button>
                        </div>
                    )}

                     {/* Modal to display uploaded files */}
            {/* <Modal show={showModal} onHide={() => setShowModal(false)} size='lg' >
              <Modal.Header closeButton>
                {showDocTable && <Modal.Title>Documents</Modal.Title>}
                {showImgModal && <Modal.Title>Gallery Images/Videos</Modal.Title>}
                {showBannerModal && <Modal.Title>Banner Images</Modal.Title>}
              </Modal.Header>
              <Modal.Body className="scrollbar" id="style-5">

                {showDocTable &&
                  (
                    <>
                      <table className="table table-bordered" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th>Serial No.</th>
                            <th>File Name</th>
                            <th>File Size</th>
                            <th className='text-center'>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {DocumentpostArr1.map((file, index) => (
                            <tr key={index}>
                              <td className='text-center'>{index + 1}</td>
                              <td>{file.fileName.replace("/sites/Intranet", "")}</td>
                              <td className='text-right'>{file.fileSize}</td>
                              <td className='text-center'> <img src={require("../../CustomAsset/trashed.svg")} style={{ width: '15px' }} onClick={() => deleteLocalFile(index, DocumentpostArr1, "docs")} /> </td>
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )
                }
                {showImgModal &&
                  (
                    <>
                      <table className="table table-bordered" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th>Serial No.</th>
                            <th> Image </th>
                            <th>File Name</th>
                            <th>File Size</th> */}
                            {/* {modeValue == 'null' && <th className='text-center'>Action</th>
                            } */}
                            {/* <th className='text-center'>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {ImagepostArr1.map((file, index) => (
                            <tr key={index}>
                              <td className='text-center'>{index + 1}</td>
                              <td>  <img className='imagefe' src={file.fileType.startsWith('video/') ?
                                        require("../../Assets/ExtraImage/video.jpg") :file.fileUrl ?file.fileUrl: `${siteUrl}/AnnouncementAndNewsGallary/${file.fileName}`}
                              /></td>

                              <td>{file.fileName}</td>
                              <td className='text-right'>{file.fileSize}</td> */}
                              {/* {modeValue != 'view' && <td className='text-center'>
                                <img src={require("../../CustomAsset/trashed.svg")} style={{ width: '15px' }} onClick={() => deleteLocalFile(index, ImagepostArr1, "Gallery")} />

                              </td>
                              } */}
                              {/* <td className='text-center'>
                                <img src={require("../../CustomAsset/trashed.svg")} style={{ width: '15px' }} onClick={() => deleteLocalFile(index, ImagepostArr1, "Gallery")} />

                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )}
                {showBannerModal &&
                  (
                    <>
                      <table className="table table-bordered" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th>Serial No.</th>
                            <th>Image</th>
                            <th>File Name</th>
                            <th>File Size</th>
                            <th className='text-center'>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {BnnerImagepostArr.map((file, index) => (
                            <tr key={index}>
                              <td className='text-center'>{index + 1}</td>
                              <img src={BnnerImagepostArr[0].fileUrl?BnnerImagepostArr[0].fileUrl:`${file.serverUrl}${file.serverRelativeUrl}`} />
                              <td>{file.name?file.name:file.fileName?file.fileName:""}</td>
                              <td className='text-right'>{file.size}</td>
                              <td className='text-center'> <img src={require("../../CustomAsset/trashed.svg")} style={{ width: '15px' }} onClick={() => deleteLocalFile(index, BnnerImagepostArr, "bannerimg")} /> </td>
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )}

              </Modal.Body>

            </Modal> */}
            {/*  */}

            {/* modal css */}

            <Modal show={showModal} onHide={() => setShowModal(false)} size='lg' className="newm" >
              <Modal.Header closeButton>
                { showDocTable && <Modal.Title>Documents</Modal.Title>}
                { showImgModal && <Modal.Title>Gallery Images/Videos</Modal.Title>}
                {BnnerImagepostArr.length > 0 && showBannerModal && <Modal.Title>Banner Images</Modal.Title>}
              </Modal.Header>
              <Modal.Body className="" id="style-5">
 
                { showDocTable &&
                  (
                    <>
                      <table className="mtbalenew" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>Serial No.</th>
                            <th style={{minWidth:'100px',maxWidth:'100px'}}>File Name</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>File Size</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {DocumentpostArr1.map((file, index) => (
                            <tr key={index}>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>{index + 1}</td>
                              <td style={{minWidth:'100px',maxWidth:'100px'}}>{file.fileName.replace("/sites/Intranetdemos", "")}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-right'>{file.fileSize}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'> <img style={{cursor:'pointer'}} src={require("../../CustomAsset/del.png")}  onClick={() => deleteLocalFile(index, DocumentpostArr1, "docs")} /> </td>
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )
                }
                { showImgModal &&
                  (
                    <>
                      <table className="mtbalenew" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th  style={{minWidth:'40px',maxWidth:'40px'}}>Serial No.</th>
                            <th style={{minWidth:'50px',maxWidth:'50px'}}> Image </th>
                            <th>File Name</th>
                            <th  style={{minWidth:'40px',maxWidth:'40px'}}>File Size</th>
                            <th className='text-center'>Action</th>
                           
                          </tr>
                        </thead>
                        <tbody>
                          {ImagepostArr1.map((file, index) => (
                            <tr key={index}>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>{index + 1}</td>
                              <td style={{minWidth:'50px',maxWidth:'50px', textAlign:'center'}}>  <img style={{width:'40px',height:'40px', borderRadius:'1000px'}} className='imagefe' src={file.fileType.startsWith('video/') ?
                                        require("../../Assets/ExtraImage/video.jpg") :file.fileUrl ?file.fileUrl: `${siteUrl}/AnnouncementAndNewsGallary/${file.fileName}`}
                              /></td>
 
                              <td  >{file.fileName}</td>
                              <td  style={{minWidth:'40px',maxWidth:'40px'}} className='text-right'>{file.fileSize}</td>
                              <td className='text-center'>
                                <img style={{cursor:'pointer'}} src={require("../../CustomAsset/del.png")}  onClick={() => deleteLocalFile(index, ImagepostArr1, "Gallery")} />
 
                              </td>
                             
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )}
                {showBannerModal &&
                  (
                    <>
                      <table className="mtbalenew" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>Serial No.</th>
                            <th style={{minWidth:'50px',maxWidth:'50px'}}>Image</th>
                            <th>File Name</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>File Size</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                        {BnnerImagepostArr[0] != null && BnnerImagepostArr.map((file, index) => (
                            <tr key={index}>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>{index + 1}</td>
                             <td style={{minWidth:'50px',maxWidth:'50px',textAlign:'center'}} >  <img style={{width:'40px',height:'40px', borderRadius:'1000px'}} src={BnnerImagepostArr[0].fileUrl?BnnerImagepostArr[0].fileUrl:`${file.serverUrl}${file.serverRelativeUrl}`} /></td>
                              {/* <td>{file.name}</td> */}
                              <td>{file.name?file.name:file.fileName?file.fileName:""}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-right'>{file.name?file.name:file.fileName?file.fileName:""}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'> <img style={{cursor:'pointer'}} src={require("../../CustomAsset/del.png")}  onClick={() => deleteLocalFile(index, BnnerImagepostArr, "bannerimg")} /> </td>
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )}
 
              </Modal.Body>
 
            </Modal>
 

 
 
  {/* <Modal show={showModal} onHide={() => setShowModal(false)} size='lg' className="newm" >
              <Modal.Header closeButton>
                {DocumentpostArr1.length > 0 && showDocTable && <Modal.Title>Documents</Modal.Title>}
                {ImagepostArr1.length > 0 && showImgModal && <Modal.Title>Gallery Images/Videos</Modal.Title>}
                {BnnerImagepostArr.length > 0 && showBannerModal && <Modal.Title>Banner Images</Modal.Title>}
              </Modal.Header>
              <Modal.Body className="" id="style-5">
 
                {DocumentpostArr1.length > 0 && showDocTable &&
                  (
                    <>
                      <table className="mtbalenew" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>Serial No.</th>
                            <th >File Name</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>File Size</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {DocumentpostArr1.map((file, index) => (
                            <tr key={index}>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>{index + 1}</td>
                              <td>{file.fileName.replace("/sites/Intranet", "")}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-right'>{file.fileSize}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}}className='text-center'> <img src={require("../../CustomAsset/del.png")} style={{ width: '15px' }} onClick={() => deleteLocalFile(index, DocumentpostArr1, "docs")} /> </td>
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )
                }
                {ImagepostArr1.length > 0 && showImgModal &&
                  (
                    <>
                      <table className="mtbalenew" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>Serial No.</th>
                            <th style={{minWidth:'50px',maxWidth:'50px'}}> Image </th>
                            <th>File Name</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>File Size</th>
                            {modeValue == 'null' && <th className='text-center'>Action</th>
                            }
                          </tr>
                        </thead>
                        <tbody>
                          {ImagepostArr1.map((file, index) => (
                            <tr key={index}>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>{index + 1}</td>
                              <td style={{minWidth:'50px',maxWidth:'50px', textAlign:'center'}}>  <img style={{Width:'40px',height:'40px', borderRadius:'1000px'}}  className='imagefe' src={file.fileUrl ?file.fileUrl: `${siteUrl}/AnnouncementAndNewsGallary/${file.fileName}`}
                              /></td>
 
                              <td>{file.fileName}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-right'>{file.fileSize}</td>
                              {modeValue != 'view' && <td className='text-center'>
                                <img src={require("../../CustomAsset/del.png")} style={{ cursor:'pointer' }} onClick={() => deleteLocalFile(index, ImagepostArr1, "Gallery")} />
 
                              </td>
                              }
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )}
                {BnnerImagepostArr.length > 0 && showBannerModal &&
                  (
                    <>
                      <table className="mtbalenew" style={{ fontSize: '0.75rem' }}>
                        <thead style={{ background: '#eef6f7' }}>
                          <tr>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>Serial No.</th>
                            <th style={{minWidth:'50px',maxWidth:'50px'}}>Image</th>
                            <th>File Name</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}}>File Size</th>
                            <th style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>Action</th>
                          </tr>
                        </thead>
                        <tbody>
                          {BnnerImagepostArr.map((file, index) => (
                            <tr key={index}>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'>{index + 1}</td>
                             <td style={{minWidth:'50px',maxWidth:'50px', textAlign:'center'}}><img style={{Width:'40px',height:'40px',borderRadius:'1000px'}} src={BnnerImagepostArr[0].fileUrl?BnnerImagepostArr[0].fileUrl:`${file.serverUrl}${file.serverRelativeUrl}`} /> </td>
                              <td>{file.name}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-right'>{file.size}</td>
                              <td style={{minWidth:'40px',maxWidth:'40px'}} className='text-center'> <img src={require("../../CustomAsset/del.png")} style={{ cursor:'pointer' }} onClick={() => deleteLocalFile(index, BnnerImagepostArr, "bannerimg")} /> </td>
                            </tr>
                          ))}
                        </tbody>
                      </table></>
                  )}
 
              </Modal.Body>
 
            </Modal> */}
 

            {/* modal css ends */}

                </div>
                {/* {itemsToShow < blogData.length && (
                    <div className="col-12 text-center mt-3 mb-3">
                        <button type='button' onClick={loadMore} className="btn btn-primary">
                            Load More
                        </button>
                    </div>
                )} */}
            </div>
}
            </>

    )
}

export default CustomBlogWebpartTemplate