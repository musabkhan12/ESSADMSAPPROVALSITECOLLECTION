import Swal from "sweetalert2";
// export const fetchBlogdata = async (_sp) => {
//   let arr = []
//   let currentUser;
//   await _sp.web.currentUser()
//     .then(user => {
//       console.log("user", user);
//       currentUser = user.Id; // Get the current user's Email
//     })
//     .catch(error => {
//       console.error("Error fetching current user: ", error);
//       return [];
//     });
//   await _sp.web.lists.getByTitle("ARGBlogs")
//     .items.select("*,Author/ID,Author/Title,Author/EMail").expand("Author").orderBy("Created", false).getAll().then((res) => {
//       console.log(res);
//       let filtereddata =
//         res.filter(x => x.Status !== "Save as Draft" || (x.Status === "Save as Draft" && x.AuthorId == currentUser))
//       arr = filtereddata;
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }

export const fetchBookmarkBlogdata = async (_sp) => {
  let arr = []
  let bookmarkarr = [];
  let currentUser;
  await _sp.web.currentUser()
    .then(user => {
      console.log("user", user);
      currentUser = user.Email; // Get the current user's Email
    })
    .catch(error => {
      console.error("Error fetching current user: ", error);
      return [];
    });

  await _sp.web.lists.getByTitle("ARGSavedBlogs")
    .items.select("*,BlogId/ID,BlogId/Title,BlogSavedBy/ID,BlogSavedBy/Title,BlogSavedBy/EMail").expand("BlogId,BlogSavedBy")
    .filter(`BlogSavedBy/EMail eq '${currentUser}'`)
    .orderBy("Created", false).getAll().then(async (resnew) => {
      console.log("resnew", resnew);
      if (resnew.length > 0) {

        for (let i = 0; i < resnew.length; i++) {
          await _sp.web.lists.getByTitle("ARGBlogs")
            .items.select("*,Author/ID,Author/Title").expand("Author")
            .filter(`ID eq ${resnew[i].BlogIdId}`)
            .orderBy("Created", false).getAll().then((res) => {
              console.log("bookmarkarr", res, resnew[i].BlogIdId);
              bookmarkarr.push(res[0])
              //res.filter(x=>x.Category?.Category==str)

            })
            .catch((error) => {
              console.log("Error fetching data: ", error);
            });
        }
      }
    })
  console.log("bookmarkarrbookmarkarr", bookmarkarr);
  arr = bookmarkarr.slice(0, 3);
  return arr;
}
export const fetchPinstatus = async (_sp) => {
  let bookmarkarrnew = [];
  let arr = []
  const initialPinStatus = {};
  const currentUser = await _sp.web.currentUser();
  await _sp.web.lists.getByTitle("ARGBlogs")
    .items.select("*,Author/ID,Author/Title,BookmarkedBy/ID,BookmarkedBy/Title,BookmarkedBy/EMail").expand("Author,BookmarkedBy").orderBy("Created", false).getAll().then(async (res) => {
      console.log(res);
      for (let i = 0; i < res.length; i++) {
        //const initialPinStatus = {};

        const pinRecords = await _sp.web.lists.getByTitle("ARGSavedBlogs").items
          .select("*,BlogId/ID,BlogId/Title,BlogSavedBy/ID,BlogSavedBy/Title,BlogSavedBy/EMail")
          .expand("BlogId,BlogSavedBy")
          .filter(`BlogSavedById eq ${currentUser.Id} and BlogIdId eq ${res[i].ID}`)
          .getAll();

        initialPinStatus[res[i].ID] = pinRecords.length > 0;
        //bookmarkarrnew.push(initialPinStatus)
        //setPinStatus((prev) => ({ ...prev, initialPinStatus: false }));
        console.log("initialPinStatus", initialPinStatus,);

        //          // });
      }

      console.log("bookmarkarrnew", initialPinStatus);
      arr = initialPinStatus;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const fetchBookmarkID = async (_sp) => {
  let arr = []
  let bookmarkarr = [];
  let currentUser;
  await _sp.web.currentUser()
    .then(user => {
      console.log("user", user);
      currentUser = user.Email; // Get the current user's Email
    })
    .catch(error => {
      console.error("Error fetching current user: ", error);
      return [];
    });

  await _sp.web.lists.getByTitle("ARGSavedBlogs")
    .items.select("*,BlogId/ID,BlogId/Title,BlogSavedBy/ID,BlogSavedBy/Title,BlogSavedBy/EMail").expand("BlogId,BlogSavedBy")
    .filter(`BlogSavedBy/EMail eq '${currentUser}'`)
    .orderBy("Created", false).getAll().then(async (resnew) => {
      console.log("resnew", resnew);
      arr = resnew;
    })
  console.log("resnewwwwww", arr);
  return arr;
}

export const getBlog = async (_sp) => {
  let arr = []
  let str = "Announcements"
  await _sp.web.lists.getByTitle("ARGBlogs").items.select("*,Author/ID,Author/Title").expand("Author").orderBy("Created", false).getAll()
    .then((res) => {
      console.log("--discussion", res);


      //res.filter(x=>x.Category?.Category==str)
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}

export const GetCategory = async (_sp) => {
  let arr = []
  await _sp.web.lists.getByTitle("ARGDiscussionForumCategory").items.orderBy("Created", false).getAll().then((res) => {
    console.log("---category", res);
    arr = res;
  }).catch((error) => {
    console.log("Error fetching data: ", error);
  })
  return arr;
}
export const GetBlogCategory = async (_sp) => {
  let arr = []
  await _sp.web.lists.getByTitle("ARGBlogCategory").items.orderBy("Created", false).getAll().then((res) => {
    console.log("---category", res);
    arr = res;
  }).catch((error) => {
    console.log("Error fetching data: ", error);
  })
  return arr;
}
export const updateItem = async (itemData, sp, id) => {
  let resultArr = []
  try {
    const newItem = await sp.web.lists.getByTitle('ARGBlogs').items.getById(id).update(itemData);
    //Swal.fire('Item update successfully', '', 'success');
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    Swal.fire(' Cancelled', '', 'error')
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};

export const addItem = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGBlogs').items.add(itemData);

    console.log('Item added successfully:', newItem);

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    Swal.fire(' Cancelled', '', 'error')
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};


export const uploadFile = async (file, _sp, docLib, siteUrl) => {
  let arr = {};

  const uploadResult = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file, data => {
    console.log(`progress`, data);
  }, true);

  const fileUrl = uploadResult.data.ServerRelativeUrl;

  const imgMetadata = {
    "__metadata": { "type": "SP.FieldUrlValue" },
    "Description": file.name,
    "Url": `${siteUrl}${fileUrl}`
  };

  // await sp.web.lists.getByTitle(docLib).items.getById(uploadResult.data.UniqueId).update({
  //   "AnnouncementandNewsBannerImage": imgMetadata
  // });
  arr = {
    "type": "thumbnail",
    "fileName": file.name,
    "serverUrl": siteUrl,
    "fieldName": "Image",
    "serverRelativeUrl": fileUrl
  };
  return arr;
};
export const uploadFileToLibrary = async (file, _sp, docLib) => {

  let arrFIleData = [];
  let fileSize = 0
  try {
    const result = await _sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file,

      // const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(
      // file.name,
      // file,
      (progress, data) => {
        console.log(progress, data);
        fileSize = progress.fileSize
      },
      true
    );
    //const resultnew = await sp.web.get
    if (result.data != undefined) {
      console.log(result.data, 'itemitemhg', docLib);
      const item = await _sp.web.lists.getByTitle(`${docLib}`).items.orderBy("Created", false)
        //.filter(`Name eq '${result.data.Name}'`)

        //.top(1)
        .getAll();
      //console.log(item1, 'itemitem111');
      //const item = await _sp.web.getFileByServerRelativePath(result.data.ServerRelativeUrl).getItem("*", "ID", "AuthorId", "Modified")
      console.log(item, 'itemitem');
      let arr = {
        ID: item[0].Id,
        Createdby: item[0].AuthorId,
        Modified: item[0].Modified,
        fileUrl: result.data.ServerRelativeUrl,
        fileSize: fileSize,
        fileType: file.type,
        fileName: file.name,
      }
      arrFIleData.push(arr)
      console.log(arrFIleData);

      return arrFIleData;
    }

  } catch (error) {
    console.log("Error uploading file:", error);
    return null; // Or handle error differently
  }
};

export const getBlogByID = async (_sp, id) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("ARGBlogs").items.getById(id).select("*,BlogCategory/ID,BlogCategory/CategoryName,Entity/ID,Entity/Entity").expand("BlogCategory,Entity")()
    .then((res) => {
      console.log(res, 'blog let arrs=[]');
      const bannerimgobject = res.BlogBannerImage != "{}" && JSON.parse(res.BlogBannerImage)
      // const bannerimgobject = res.AnnouncementandNewsBannerImage != "{}" && JSON.parse(res.AnnouncementandNewsBannerImage)
      // console.log(bannerimgobject[0], 'bannerimgobject');

      bannerimg.push(bannerimgobject);
      const parsedValues = {
        Title: res.Title != undefined ? res.Title : "",
        description: res.Description != undefined ? res.Description : "",
        //   overview: res.Overview != undefined ? res.Overview : "",
        //   IsActive: res.IsActive,
        ID: res.ID,
        // BannerImage: bannerimg,
        //   TypeMaster: res?.AnnouncementandNewsTypeMaster?.ID != undefined ? res.AnnouncementandNewsTypeMaster?.ID : "",
        Category: res.BlogCategory?.ID != undefined ? res.BlogCategory?.ID : "",
        Entity: res.Entity?.ID != undefined ? res.Entity?.ID : "",
        ///FeaturedAnnouncement: res.FeaturedAnnouncement,
        BlogGalleryJSON: res.BlogGalleryJSON != null ? JSON.parse(res.BlogGalleryJSON) : "",
        BlogDocsJSON: res.BlogDocsJSON != null ? JSON.parse(res.BlogDocsJSON) : "",
        BlogGalleryId: res.BlogGalleryId,
        BlogDocsId: res.BlogDocsId
        // other fields as needed
      };

      arr.push(parsedValues)
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arrblllll');
  return arr;
}
export const getBlogsByID = async (_sp, id) => {

  let arr = []
  let arrs = []
  let bannerimg = []
  await _sp.web.lists.getByTitle("ARGBlogs").items.getById(id).select("*,Author/ID,Author/Title,Entity/ID,Entity/Entity").expand("Author,Entity")()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      const bannerimgobject = res.BlogBannerImage != "{}" && JSON.parse(res.BlogBannerImage)
      console.log(bannerimgobject, 'bannerimgobject');

      bannerimg.push(bannerimgobject);
      const parsedValues = {
        Title: res.Title != undefined ? res.Title : "",
        description: res.Description != undefined ? res.Description : "",
        //   overview: res.Overview != undefined ? res.Overview : "",
        //   IsActive: res.IsActive,
        ID: res.ID,
        Topic: res.Title,
        Overview: res.Overview,
        BannerImage: bannerimg,
        //GroupType: res.GroupType,
        //   TypeMaster: res?.AnnouncementandNewsTypeMaster?.ID != undefined ? res.AnnouncementandNewsTypeMaster?.ID : "",
        //Category: res.DiscussionForumCategory?.ID != undefined ? res.DiscussionForumCategory?.ID : "",
        Entity: res.Entity?.ID != undefined ? res.Entity?.ID : "",
        //FeaturedAnnouncement: res.FeaturedAnnouncement,
        BlogGalleryJSON: res.BlogGalleryJSON != null ? JSON.parse(res.BlogGalleryJSON) : "",
        BlogDocsJSON: res.BlogDocsJSON != null ? JSON.parse(res.BlogDocsJSON) : "",
        BlogGalleryId: res.BlogGalleryId,
        BlogsDocsId: res.BlogsDocsId,

        // other fields as needed

      };

      arr.push(parsedValues);
      console.log("arrarrarrarr", arr, parsedValues);
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arr, 'arr');
  return arr;
}
//new code 
export const DeleteBusinessAppsAPI = async (_sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGBlogs').items.getById(id).delete();
    console.log('Item deleted successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
}
// new code end
export const getAllBlogsnonselected = async (_sp, Idnum, categoryId) => {
  let arr = []
  let str = "Announcements"
  await _sp.web.lists.getByTitle("ARGBlogs").items
    .select("*,BlogCategory/ID,BlogCategory/CategoryName").expand("BlogCategory")
    //.filter(`ID ne ${Idnum} `)
    .filter(`ID ne ${Idnum} and Status eq'Approved'`)
    //and BlogCategoryId eq '${categoryId}'
    .top(3)
    .orderBy("Created", false)
    .getAll()
    .then((res) => {
      let resnew = res.slice(0, 3);
      console.log("getallBlogs excluding", res, resnew);

      //res.filter(x=>x.Category?.Category==str)
      arr = resnew;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
// export const getBlogDetailsById = async (_sp, idNum) => {
//   let arr = []
//   let arr1 = []

//   await _sp.web.lists
//     .getByTitle("ARGBlogs")
//     .items.select("*,BlogCategory/ID,BlogCategory/CategoryName")
//     .expand("BlogCategory")
//     .filter(`ID eq ${Number(idNum)}`)()
//     .then((res) => {
//       // arr=res;
//       arr1.push(res)
//       arr = res
//     }).catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }

// export const uploadFile = async (file, sp, docLib, siteUrl) => {
//   var arr ={};
//   sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file, data => {
//     console.log(`progress`, data);
//   }, true).then(async result => {
//     console.log(result, 'result')
//   })
//   //  const siteUrl=  await getUrl(sp)
//   const img = {
//     "type": "thumbnail",
//     "fileName": file.name,
//     "serverUrl": siteUrl,
//     "fieldName": "BlogBannerImage",
//     "serverRelativeUrl": '/Shared%20Documents/' + file.name
//   };
//   arr=img
//   return arr;
// };
export const uploadFileBlog = async (file, sp, docLib, siteUrl) => {
  let arr = {};

  const uploadResult = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file, data => {
    console.log(`progress`, data);
  }, true);

  const fileUrl = uploadResult.data.ServerRelativeUrl;

  const imgMetadata = {
    "__metadata": { "type": "SP.FieldUrlValue" },
    "Description": file.name,
    "Url": `${siteUrl}${fileUrl}`
  };

  // await sp.web.lists.getByTitle(docLib).items.getById(uploadResult.data.UniqueId).update({
  //   "AnnouncementandNewsBannerImage": imgMetadata
  // });
  arr = {
    "type": "thumbnail",
    "fileName": file.name,
    "serverUrl": siteUrl,
    "fieldName": "BlogBannerImage",
    "serverRelativeUrl": fileUrl
  };
  return arr;
};
export const fetchBlogdatatop = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("ARGBlogs").items.select("*,Author/ID,Author/Title,BlogCategory/Id,BlogCategory/CategoryName").expand("Author,BlogCategory").top(4)().then((res) => {
    console.log(res);

    //res.filter(x=>x.Category?.Category==str)
    arr = res;
  })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const uploadFileBanner = async (file, sp, docLib, siteUrl) => {
  let arr = {};

  const uploadResult = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file, data => {
    console.log(`progress`, data);
  }, true);

  const fileUrl = uploadResult.data.ServerRelativeUrl;

  const imgMetadata = {
    "__metadata": { "type": "SP.FieldUrlValue" },
    "Description": file.name,
    "Url": `${siteUrl}${fileUrl}`
  };

  // await sp.web.lists.getByTitle(docLib).items.getById(uploadResult.data.UniqueId).update({
  //   "AnnouncementandNewsBannerImage": imgMetadata
  // });
  arr = {
    "type": "thumbnail",
    "fileName": file.name,
    "serverUrl": siteUrl,
    "fieldName": "BlogBannerImage",
    "serverRelativeUrl": fileUrl
  };
  return arr;
};

export const getProjectDetailsById = async (_sp, idNum) => {
  let arr = []
  let arr1 = []
  //,TeamMembers/EMail
  await _sp.web.lists.getByTitle("ARGProject").items.getById(idNum).select("*, ProjectStatus , TeamMembers/Id,TeamMembers/Title, Author/ID,Author/Title,Author/EMail,Author/SPSPicturePlaceholderState").expand("TeamMembers ,Author")()

    .then((res) => {
      console.log("check the data for project id--->>", res)
      // arr=res;
      arr1.push(res)
      arr = arr1
    }).catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
// export const fetchBlogdata = async (_sp) => {
//   let arr = []
//   let currentUser;
//   await _sp.web.currentUser()
//   .then(user => {
//       console.log("user", user);
//       currentUser = user.Id; // Get the current user's Email
//     })
//     .catch(error => {
//       console.error("Error fetching current user: ", error);
//       return [];
//     });
//   await _sp.web.lists.getByTitle("ARGBlogs")
//     .items.select("*,Author/ID,Author/Title,Author/EMail").expand("Author").orderBy("Created", false).getAll().then(async (res) => {
//       console.log(res);
//       let filtereddata =
//         res.filter(x => x.Status !== "Save as Draft" || (x.Status === "Save as Draft" && x.AuthorId == currentUser))
//      // Add a key based on existence in another list
 
 
//      for (const item of filtereddata) {
 
//       await _sp.web.lists.getByTitle("ARGMyRequest").items.select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title").expand("Approver,Requester")
   
//       .filter(`Status eq 'Pending' and ContentId eq ${item.ID} and EntityId eq ${item.EntityId} and ProcessName eq 'Blog' and ApproverId eq ${currentUser} and ApprovalTask eq 'Assignment'`)
 
//       .orderBy("Created",false)
 
//       .getAll().then((resp) => {
 
//         item.WillReworkEdit = resp.length > 0;
 
//       })
//       // try {
//       //   // Check existence in another list
//       //   await _sp.web.lists.getByTitle("ARGMyRequest").items.select("*").filter(`Status eq 'Pending' and ContentId eq ${item.ID} and EntityId eq ${item.EntityId} and ProcessName eq 'Blog' and ApproverId eq ${currentUser} and ApprovalTask eq 'Assignment' and IsRework eq 'Yes'`).top(1).getAll().then(resp => {
//       //     item.WillReworkEdit = resp.length > 0;
//       //   });
         
//       // } catch (error) {
//       //  // console.error(`Error checking item ${item.Id} in AnotherList: `, error);
//       //   item.WillReworkEdit = false; // Default to false on error
//       // }
//     }    
     
     
//         arr = filtereddata;
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }
export const getBlogDetailsById = async (_sp, idNum) => {
  let arr = []
  let arr1 = []
 
  await _sp.web.lists
    .getByTitle("ARGBlogs")
    .items.select("*,BlogCategory/ID,BlogCategory/CategoryName,Entity/Entity")
    .expand("BlogCategory,Entity")
    .filter(`ID eq ${Number(idNum)}`)()
    .then((res) => {
      // arr=res;
      arr1.push(res)
      arr = res
    }).catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const fetchBlogdata = async (_sp) => {
  let arr = []
  let currentUser;
  await _sp.web.currentUser()
  .then(user => {
      console.log("user", user);
      currentUser = user.Id; // Get the current user's Email
    })
    .catch(error => {
      console.error("Error fetching current user: ", error);
      return [];
    });
  await _sp.web.lists.getByTitle("ARGBlogs")
    .items.select("*,Author/ID,Author/Title,Author/EMail,Editor/ID,Editor/Title,Editor/EMail").expand("Author,Editor").orderBy("Created", false).getAll().then(async (res) => {
      console.log(res);
      let filtereddata =
        res.filter(x => x.Status !== "Save as Draft" || (x.Status === "Save as Draft" && (x.AuthorId == currentUser ||x.EditorId == currentUser)))
     // Add a key based on existence in another list
 
 
     for (const item of filtereddata) {
 
      await _sp.web.lists.getByTitle("ARGMyRequest").items.select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title").expand("Approver,Requester")
   
      .filter(`Status eq 'Pending' and ContentId eq ${item.ID} and EntityId eq ${item.EntityId} and ProcessName eq 'Blog' and ApproverId eq ${currentUser} and ApprovalTask eq 'Assignment'`)
 
      .orderBy("Created",false)
 
      .getAll().then((resp) => {
 
        item.WillReworkEdit = resp.length > 0;
 
      })
      // try {
      //   // Check existence in another list
      //   await _sp.web.lists.getByTitle("ARGMyRequest").items.select("*").filter(`Status eq 'Pending' and ContentId eq ${item.ID} and EntityId eq ${item.EntityId} and ProcessName eq 'Blog' and ApproverId eq ${currentUser} and ApprovalTask eq 'Assignment' and IsRework eq 'Yes'`).top(1).getAll().then(resp => {
      //     item.WillReworkEdit = resp.length > 0;
      //   });
         
      // } catch (error) {
      //  // console.error(`Error checking item ${item.Id} in AnotherList: `, error);
      //   item.WillReworkEdit = false; // Default to false on error
      // }
    }    
     
     
        arr = filtereddata;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
 