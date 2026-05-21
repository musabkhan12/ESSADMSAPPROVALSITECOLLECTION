import Swal from 'sweetalert2';
import { getSP } from '../webparts/essaApproval/loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";
import { AssignFrom } from "@pnp/core";
import "@pnp/sp/items";
import "@pnp/sp/site-users";
const sp = getSP();
export const getLevel = async (sp) => {
  let arr = []
  await sp.web.lists.getByTitle("ARGLevelMaster").items.select("Level,Id")().then((res) => {
    arr = res
    console.log(arr, 'arr');

  })
  return arr
}
export const getLevelId = async (levelName) => {
  try {
    const items = await sp.web.lists.getByTitle("ARGLevelMaster").items
      .select("Id", "Level")
      .filter(`Level eq '${levelName}'`)(); // Fetch the ID based on the level name

    return items.length > 0 ? items[0].Id : null; // Return the ID if found
  } catch (error) {
    console.error("Error fetching level ID:", error);
    return null;
  }
};

export const AddDataonConfuguration = async (sp, itemData) => {
  let arr = []
  await sp.web.lists.getByTitle("ARGApprovalConfiguration").items.add(itemData).then((res) => {
    arr = res
    console.log(arr, 'arr');
  })
  return arr
}

export const GetARGApprovalConfiguration = async (sp) => {
  let arr = []
  let sampleDataArray = []
  await sp.web.lists.getByTitle("ARGApprovalConfiguration").items.select("*,Users/Id,Users/Title,Users/EMail").expand("Users")().then((res) => {
    arr = res
    console.log(arr, 'arr');
    for (let i = 0; i < arr.length; i++) {
      let ars = {
        entity: arr[i].EntityId,
        levels: arr[i].Users,
        rework: arr[i].Rework0
      }
      sampleDataArray.push(ars)
    }

  })
  return sampleDataArray
}
export const getApprovalConfiguration = async (sp, EntityId) => {
  debugger
  let arr = []
  let sampleDataArray = []
  arr = await sp.web.lists.getByTitle("ARGApprovalConfiguration").items.select("*,Users/ID,Users/Title,Users/EMail,Level/Id,Level/Level").expand("Users,Level").filter(`EntityId eq ${EntityId}`)();
  // .then((res) => {
  //   arr = res
  //   console.log(arr, 'arr');
  // })
  return arr
}
export const AddContentLevelMaster = async (sp, itemData) => {
  let arr = []
  await sp.web.lists.getByTitle("ARGContentLevelMaster").items.add(itemData).then((res) => {
    arr = res
    console.log(arr, 'arr');
  })
  return arr
}
export const AddContentMaster = async (sp, itemData) => {
  let arr = []
  await sp.web.lists.getByTitle("ARGContentMaster").items.add(itemData).then((res) => {
    arr = res
    console.log(arr, 'arr');
  })
  return arr
}


export const UpdateContentMaster = async (sp, contentmasteritemid, itemData) => {
  let arr;
  await sp.web.lists.getByTitle("ARGContentMaster").items.getById(contentmasteritemid).update(itemData).then((res) => {
    arr = res
    console.log(arr, 'arr');
  })
  return arr
}

//My request
export const getListDataFromSiteCollection = async (_sp, listName, status, portal, SiteBaseURL) => {
  // Setup PnPJs for a specific site collection URL
  let arr = [];
  console.log("sdsssss", sp, _sp)
  
  let apiUrl;
  let FinalStatus = status;
  // if (status == "Pending") {
  //   FinalStatus = "Not Started"
  // } else if (status == "Approved") {
  //   FinalStatus = "Completed"
  // }
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
  if (!currentUser) return arr; // Return empty array if user fetch failed
  apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail&$expand=Requestor_x0020_Name&$filter=${`Requestor_x0020_Name/EMail eq '${currentUser}' and RequestStatus eq '${FinalStatus}'`}`;


  try {
    console.log("apiUrlapiUrl", apiUrl);
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
      },
    });
    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }
    // Wait for the response to be converted to JSON
    const data = await response.json();
    // Immediately handle data
    console.log('List Items ellllllllnew:', data.value);
    // You can now manipulate or return the data as needed
    if (data.value.length > 0) {
      arr = data.value;
    } else {
      arr = []
    }
    return arr;  // Directly returning the array if needed
  } catch (error) {
    console.error('Error fetching data:', error);
  }
}

export const getDataFromMultipleSites = async (_sp, listName, status, portal, SiteBaseURL) => {
  const allData = [];
  const siteUrls = [
    SiteBaseURL
  ];
  if(portal == "Others"){
  // Loop through each site collection URL and fetch data
  for (const siteUrl of siteUrls) {
    const data = await getListDataFromSiteCollection(_sp, listName, status, portal, siteUrl);
    allData.push(...data);
    console.log("dadadadadad", data);
  }
} else{
    for (const siteUrl of siteUrls) {
      const data = await getMyRequestsdata(_sp, listName, status, portal, siteUrl);
      allData.push(...data);
      console.log("dadadadadad", data);
    }
}
  console.log('Combined data from all site collections:', allData);
  return allData;
}

export const getRequestListsData = async (_sp, status) => {
  
  let arr = []
  let Status = status == "Pending" ? "Submitted" : status;
  await _sp.web.lists.getByTitle("AllRequestLists").items.filter(`IsActive eq 'Yes'`).orderBy("Created", false)()
    .then(async (res) => {
      console.log("AllRequestLists", res);
      let AllRequestArr = [];
      for (let i = 0; i < res.length; i++) {
        if (res[i].RedirectionLinkSource == "Others" && res[i].Portal == "Others") {
          //alert("othererrr")
          await getDataFromMultipleSites(_sp, res[i].Title, status, res[i].RedirectionLinkSource, res[i].SiteBaseURL)
            .then((resData) => {
              if (resData && resData.length > 0) {
                for (let j = 0; j < resData.length; j++) {
                  AllRequestArr.push({
                    ID: resData[j].ID,
                    RequestID: resData[j].Title,
                    // ApprovalTitle: res[i].RequestTitle,RequestTitle
                    ApprovalTitle: resData[j]?.RequestTitle!= ""? resData[j]?.RequestTitle:"",
                    Author: resData[j].Requestor_x0020_Name,
                    ProcessName: res[i].ProcessName,
                    SiteBaseURL:res[i].SiteBaseURL,
                    listName: res[i].Title,
                    Created: new Date(resData[j].Created),
                    Status: resData[j].RequestStatus,
                    AppID: res[i].AppId,
                    RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].ID}`
                  })
                }
              }
              // Handle the combined data here
              //console.log("Final combined data:", data);
            })
            .catch((error) => {
              console.error("Error fetching data:", error);
            });
        } else {
         await getDataFromMultipleSites(_sp, res[i].Title, Status, res[i].RedirectionLinkSource, res[i].SiteBaseURL).then((resData) => {
            if (resData && resData.length > 0) {
              for (let j = 0; j < resData.length; j++) {
                AllRequestArr.push({
                  ID: resData[j].ID,
                  RequestID: resData[j].RequestID,
                  ApprovalTitle: resData[j].ApprovalTitle,
                  Author: resData[j].Author,
                  ProcessName: resData[j].ProcessName,
                  SiteBaseURL:res[i].SiteBaseURL,
                  listName: res[i].Title,
                  Created: new Date(resData[j].Created),
                  Status: resData[j].Status,
                  TaskID: "",
                  AppID: "",
                  //RedirectionLink: resData[j].RedirectionLink
                  RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}${resData[j].RedirectionLink}`
                })
                //AllRequestArr.push(resData[j])
              }
            }
          })
        }
      }
     
      console.log("AllRequestArr", AllRequestArr);
      arr = AllRequestArr;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}

export const getMyRequestsdata = async (_sp, listName, status, portal, SiteBaseURL) => {
  // Setup PnPJs for a specific site collection URL
  let arr = [];
  console.log("sdsssss", sp, _sp)
  
  let apiUrl;
  let FinalStatus = status;
  // if (status == "Pending") {
  //   FinalStatus = "Pending"
  // } else if (status == "Approved") {
  //   FinalStatus = "Approved"
  // }
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
  if (!currentUser) return arr; // Return empty array if user fetch failed
  apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Author/ID,Author/Title,Author/EMail,EmployeeName/Title,EmployeeName/ID,EmployeeName/EMail&$expand=Author,EmployeeName&$filter=${`EmployeeName/EMail eq '${currentUser}' and Status eq '${FinalStatus}'`}`;


  try {
    console.log("apiUrlapiUrl", apiUrl);
    const response = await fetch(apiUrl, {
      method: 'GET',
      headers: {
        'Accept': 'application/json',
      },
    });
    if (!response.ok) {
      throw new Error(`HTTP error! Status: ${response.status}`);
    }
    // Wait for the response to be converted to JSON
    const data = await response.json();
    // Immediately handle data
    console.log('List Items ellllllllnew:', data.value);
    // You can now manipulate or return the data as needed
    if (data.value.length > 0) {
      arr = data.value;
    } else {
      arr = []
    }
    return arr;  // Directly returning the array if needed
  } catch (error) {
    console.error('Error fetching data:', error);
  }
}

// export const getMyRequestsdata = async (_sp, listName, status, portal, SiteBaseURL) => {

//   let arr = []

//   let currentUser;

//   await _sp.web.currentUser()

//     .then(user => {

//       console.log("usertesttt", user, listName, status);

//       currentUser = user.Id; // Get the current user's Email

//     })

//     .catch(error => {

//       console.error("Error fetching current user: ", error);

//       return [];

//     });


//   if (!currentUser) return arr; // Return empty array if user fetch failed


//   await _sp.web.lists.getByTitle(listName).items

//     .select("*,Author/ID,Author/Title,Author/EMail").expand("Author")

//     .filter(`AuthorId eq ${currentUser} and Status eq '${status}'`)

//     .orderBy("Created", false)()

//     .then((res) => {

//       console.log(`--MyRequest${listName}`, res);

//       arr = res

//       // arr = res.filter(item => 

//       //     // Include public groups or private groups where the current user is in the InviteMembers array

//       //     item.GroupType === "Public" || 

//       //     (item.GroupType === "Private" && item.InviteMemebers && item.InviteMemebers.some(member => member.Id === currentUser))

//       //   );

//     })

//     .catch((error) => {

//       console.log("Error fetching data: ", error);

//     });

//   return arr;

// }

// Intranet My Request
export const getRequestListsDataIntranet = async (_sp, status) => {
  let arr = []
  await _sp.web.lists.getByTitle("AllRequestLists").items.orderBy("Created", false)()
    .then((res) => {
      console.log("AllRequestLists i ntranett", res);
      let AllRequestArr = [];
      for (let i = 0; i < res.length; i++) {
        getMyRequestsdataIntranet(_sp, res[i].Title, status).then((resData) => {
          for (let j = 0; j < resData.length; j++) {
            AllRequestArr.push(resData[j])
          }
        })
      }
      console.log("AllRequestArr intra", AllRequestArr);
      arr = AllRequestArr;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const getMyRequestsdataIntranet = async (_sp, listName, status) => {
  let arr = []
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
  if (!currentUser) return arr; // Return empty array if user fetch failed
  await _sp.web.lists.getByTitle(listName).items
    .select("*,Author/ID,Author/Title,Author/EMail").expand("Author")
    .filter(`Author/EMail eq '${currentUser}' and Status eq '${status}'`)
    .orderBy("Created", false)()
    .then((res) => {
      console.log(`--MyRequestintra${listName}`, res);
      arr = res
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
//Intranet My Request end
// export const getMyRequest = async (sp,status)=>
// {
//   const currentUser = await sp.web.currentUser();
//   let arr = []
//   await sp.web.lists.getByTitle("ARGMyRequest").items.select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title")
//     .expand("Approver,Requester").filter(`RequesterId eq ${currentUser.Id} and Status eq '${status}'`)
//   .orderBy("Created",false)
//   ().then((res) => {

//     arr = res
//     console.log(arr, 'arr');
//   })
//   return arr
// }
// export const getMyRequest = async (sp, status) => {
//   const currentUser = await sp.web.currentUser();
//   let arr = []
//   await sp.web.lists.getByTitle("ARGContentMaster").items.select("*,Author/Id,Author/Title")
//     .expand("Author").filter(`AuthorId eq ${currentUser.Id} and Status eq '${status}'`)
//     .orderBy("Created", false)
//     ().then((res) => {

//       arr = res
//       console.log(arr, 'arr');
//     })
//   return arr
// }
export const getMyRequest = async (sp, status) => {
  try {
    const currentUser = await sp.web.currentUser();

    // 1. Get the list of Site URLs from the Master list on the current site
    const masterSites = await sp.web.lists
      .getByTitle("MasterSiteCollection")
      .items.select("SiteURL")();

    // 2. Prepare promises to fetch data from each site
    const fetchPromises = masterSites.map(async (site) => {
      const siteUrl = site.SiteURL;
      
      try {
        // Create a web object for the specific site using the current context's auth
        const remoteWeb = Web(siteUrl).using(AssignFrom(sp.web));

        // Fetch requests where Author is the current user
        const res = await remoteWeb.lists
          .getByTitle("ARGContentMaster")
          .items.select("*,Author/Id,Author/Title")
          .expand("Author")
          .filter(`AuthorId eq ${currentUser.Id} and Status eq '${status}'`)
          .orderBy("Created", false)();

        return res;
      } catch (siteError) {
        // Log error for specific site but don't break the whole process
        console.warn(`ARGContentMaster list not found or access denied at: ${siteUrl}`);
        return []; 
      }
    });

    // 3. Wait for all requests to finish and flatten the results into one array
    const results = await Promise.all(fetchPromises);
    allRequests = results.flat();

    console.log("Aggregated Requests:", allRequests);
    return allRequests;

  } catch (error) {
    console.error("Error in getMyRequest cross-site fetch:", error);
    return [];
  }
};
// export const getMyApproval = async (sp, status) => {

//   const currentUser = await sp.web.currentUser();

//   let arr = []

//   await sp.web.lists.getByTitle("ARGMyRequest").items.select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title").expand("Approver,Requester")

//     .filter(`ApproverId eq ${currentUser.Id} and Status eq '${status}'`)

//     .orderBy("Created", false)

//     ().then((res) => {

//       arr = res

//       console.log(arr, 'arr');

//     })

//   return arr

// }
// export const gteDMSApproval = async(sp)=>{
//   alert("DMS")
//   const currentUser = await sp.web.currentUser();
//   console.log(currentUser , "currentUser")
//   let arr = []
//   const FilesItems = await sp.web.lists
//   .getByTitle("MasterSiteURL")
//   .items.select("Title", "SiteID", "FileMasterList", "Active")
//   .filter(`Active eq 'Yes'`)();

//   console.log(FilesItems , "FilesItems")
//   FilesItems.forEach(async (fileItem, index) => {
//     if (fileItem.FileMasterList !== null) {
//       // if (siteIdToUpdate && fileItem.SiteID !== siteIdToUpdate) {
//       //   return;
//       // }

//       console.log("fileItem.FileMasterList",fileItem.FileMasterList);

//       const filesData = await sp.web.lists
//             .getByTitle(`${fileItem.FileMasterList}`)
//             .items.select("ID" , "FileName", "FileUID", "FileSize", "FileVersion" ,"Status" , "SiteID","CurrentFolderPath","DocumentLibraryName","SiteName","FilePreviewURL","IsDeleted","MyRequest").filter(
//               `CurrentUser eq '${currentUser.Email}' and MyRequest eq 1 and Status eq 'Pending'`
//             ).orderBy("Modified", false)().then((res) => {
//               arr = res
//             });
//       console.log("My reaquest Called");

//       // console.log("enter in the myRequest------")
//       console.log(fileItem.FileMasterList,"- FilesData",filesData)
//     // route to different-2 sideBar
//      console.log(arr , "DMS My request Data")
//      return arr

//     }
//   });

// }
// export const gteDMSApproval = async (sp) => {
//   // alert("DMS");
//   const currentUser = await sp.web.currentUser();
//   console.log(currentUser, "currentUser");
//   let arr = [];

//   const FilesItems = await sp.web.lists
//     .getByTitle("MasterSiteURL")
//     .items.select("Title", "SiteID", "FileMasterList", "Active")
//     .filter(`Active eq 'Yes'`)();

//   console.log(FilesItems, "FilesItems");

//   // Use for...of loop for proper async/await handling
//   for (const fileItem of FilesItems) {
//     if (fileItem.FileMasterList !== null) {
//       console.log("fileItem.FileMasterList", fileItem.FileMasterList);
//       //  alert(currentUser.Email)
//       const filesData = await sp.web.lists
//         .getByTitle(fileItem.FileMasterList)
//         .items.select("ID", "FileName", "FileUID", "FileSize", "FileVersion", "Status", "SiteID", "CurrentFolderPath", "DocumentLibraryName", "SiteName", "FilePreviewURL", "IsDeleted", "MyRequest" , "*")
//         .filter(`CurrentUser eq '${currentUser.Email}' and MyRequest eq 1 and Status eq 'Pending'`)
//         .orderBy("Modified", false)
//         ();

//       arr = [...arr, ...filesData]; // Collect data in the array
//       console.log(arr, "DMS My request Data");
//     }
//   }

//   return arr; // Return the collected data
// }
export const getMyApproval = async (sp, status, actingfor) => {
  try {
    // alert(`Actingfor is ${actingfor}`);
    let arr = [];

    if (!actingfor) {
      // alert(`Actingfor is null ${actingfor}`);
      const currentUser = await sp.web.currentUser();

      arr = await sp.web.lists.getByTitle("ARGMyRequest").items.select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title", "Approver/EMail")
        .expand("Approver,Requester")
        .filter(`ApproverId eq ${currentUser.Id} and Status eq '${status}'`)
        .orderBy("Created", false)
        ();

      console.log(arr, 'arr of intranet if actingfor is null');
    } else {
      // alert(`Actingfor is not null ${actingfor}`);
      const user = await sp.web.siteUsers.getByEmail(actingfor)();
      // alert(user.Id);

      if (user.Id) {
        arr = await sp.web.lists.getByTitle("ARGMyRequest").items.select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title", "Approver/EMail")
          .expand("Approver,Requester")
          .filter(`ApproverId eq ${user.Id} and Status eq '${status}'`)
          .orderBy("Created", false)
          ();

        console.log(arr, 'arr of intranet if actingfor is not null');
      } else {
        console.log("User not found in Approval");
      }
    }

    return arr;
  } catch (error) {
    console.error("Error fetching list items:", error);
    return [];
  }
};
export const gteDMSApproval = async (sp, value) => {
  try {
    const currentUser = await sp.web.currentUser();
    console.log(currentUser, "currentUser");
    let arr = [];

    const FilesItems = await sp.web.lists
      .getByTitle("MasterSiteURL")
      .items.select("Title", "SiteID", "FileMasterList", "Active")
      .filter(`Active eq 'Yes'`)();

    console.log(FilesItems, "FilesItems");

    for (const fileItem of FilesItems) {
      if (fileItem.FileMasterList) {
        console.log("fileItem.FileMasterList", fileItem.FileMasterList);
        const filesData = await sp.web.lists
          .getByTitle(fileItem.FileMasterList)
          .items.select("ID", "FileName", "FileUID", "FileSize", "FileVersion", "Status", "SiteID", "CurrentFolderPath", "DocumentLibraryName", "SiteName", "FilePreviewURL", "IsDeleted", "MyRequest", "*")
          .filter(`CurrentUser eq '${currentUser.Email}' and MyRequest eq 1 and Status eq '${value}'`)
          .orderBy("Created", false)
          ();

        arr = [...arr, ...filesData];
        console.log(arr, "DMS My request Data");
      }
    }

    return arr; // Return the collected data
  } catch (error) {
    console.error("Error in gteDMSApproval function", error);
    return []; // Return an empty array on error
  }
};


export const getDataByID = async (_sp, id, ContentName) => {
  debugger
  let arr = []
  let arrs = []
  let bannerimg = []
  if (ContentName != null && ContentName != undefined)
  //  alert(ContentName )
  {
    await _sp.web.lists.getByTitle(ContentName).items.getById(id)
      ()
      .then((res) => {
        console.log(res, ' let arrs=[]');
        arrs.push(res)
        arr = arrs
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    console.log(arr, 'arr');
  }

  return arr;
}
export const updateItemApproval = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGMyRequest').items.getById(id).update(itemData);
    Swal.fire('Item update successfully', '', 'success');
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

export const getMyRequestBlog = async (sp, item) => {

  const currentUser = await sp.web.currentUser();

  let arr = []

  await sp.web.lists.getByTitle("ARGMyRequest").items.select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title").expand("Approver,Requester")

    .filter(`Status eq 'Pending' and ContentId eq ${item.ID} and EntityId eq ${item.EntityId} and ProcessName eq 'Blog' and ApproverId eq ${currentUser.Id} and ApprovalTask eq 'Approval'`)

    .orderBy("Created", false)

    ().then((res) => {

      arr = res

      console.log(arr, 'arr');

    })

  return arr

}

export const getMyRequestBlogPending = async (sp, item) => {

  const currentUser = await sp.web.currentUser();

  let arr = []

  await sp.web.lists.getByTitle("ARGMyRequest").items.select("*,Requester/Id,Requester/Title,Approver/Id,Approver/Title").expand("Approver,Requester")

    .filter(`Status eq 'Pending' and ContentId eq ${item.ID} and EntityId eq ${item.Entity} and ProcessName eq 'Blog' and ApproverId eq ${currentUser.Id} and ApprovalTask eq 'Assignment'`)

    .orderBy("Created", false)

    ().then((res) => {

      arr = res

      console.log(arr, 'arr');

    })

  return arr

}