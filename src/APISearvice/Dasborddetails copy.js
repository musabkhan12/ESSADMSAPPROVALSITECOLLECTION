// export const getAnncouncementone = async (_sp) => {
//   let arr = []
//   let str = "Announcements"
//   let lengtharr = []
//   await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.select("*,AnnouncementandNewsTypeMaster/Id,AnnouncementandNewsTypeMaster/TypeMaster,Category/Id,Category/Category").expand("AnnouncementandNewsTypeMaster,Category").filter(`AnnouncementandNewsTypeMaster/TypeMaster eq '${str}'`).top(2)().then(async (res) => {
//     console.log(res);


//     for (var i = 0; i < res.length; i++) {

//       await _sp.web.lists.getByTitle("ARGAnnouncementandNewsComments").items.select("*,AnnouncementAndNews/Id").expand("AnnouncementAndNews")
//         .filter(`AnnouncementAndNewsId eq ${res.Id}`)().then(async (ress) => {
//           debugger
//           res[i].CommentCount = ress.length
//           if (ress.length > 0) {
//             ress.forEach(async item1 => {
//               res[i].LikeCount = item1.LikesCount
//             });
//           }
//           else {
//             res[i].CommentCount = 0
//             res[i].LikeCount = 0
//           }

//         });

//     }
//     arr = res

//   })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }
export const getAnncouncementone = async (_sp) => {
  let arr = [];
  let str = "Announcement";

  try {
    let res = await _sp.web.lists
      .getByTitle("ARGAnnouncementAndNews")
      .items.select("*,AnnouncementandNewsTypeMaster/Id,AnnouncementandNewsTypeMaster/TypeMaster,Category/Id,Category/Category")
      .expand("AnnouncementandNewsTypeMaster,Category")
      .filter(`AnnouncementandNewsTypeMaster/TypeMaster eq '${str}' and Status eq 'Approved'`)
      .orderBy("Modified", false)
      .top(2)();

    // Iterate over the results
    for (let i = 0; i < res.length; i++) {
      let announcementId = res[i].Id;

      // Fetch comments and like counts
      let comments = await _sp.web.lists
        .getByTitle("ARGAnnouncementandNewsComments")
        .items.select("*,AnnouncementAndNews/Id")
        .expand("AnnouncementAndNews")
        .filter(`AnnouncementAndNewsId eq ${announcementId}`)();

      // Set comment and like counts
      res[i].CommentCount = comments.length;
      res[i].LikeCount = comments.length > 0 ? comments.reduce((acc, comment) => acc + (comment.LikesCount || 0), 0) : 0;
    }

    arr = res;
  } catch (error) {
    console.log("Error fetching data: ", error);
  }

  return arr;
};
export const getNewsone = async (_sp) => {
  let arr = []
  let str = "News"
  await _sp.web.lists.getByTitle("ARGAnnouncementAndNews").items.select("*,AnnouncementandNewsTypeMaster/Id,AnnouncementandNewsTypeMaster/TypeMaster,Category/Id,Category/Category").expand("AnnouncementandNewsTypeMaster,Category")
    .filter(`AnnouncementandNewsTypeMaster/TypeMaster eq '${str}' and Status eq 'Approved'`)
    .orderBy("Modified", false).top(2)().then((res) => {
      console.log(res);

      //res.filter(x=>x.Category?.Category==str)
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}

export const fetchEventdataone = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("ARGEventMaster").items
    .select("*")
    .filter(`Status eq 'Approved'`)
    .orderBy("EventDate", true)
    ()

    .then((res) => {
      console.log(res);
      const today = new Date();
      today.setHours(0, 0, 0, 0); 
 
      
      const filteredEvents = res.filter(event => {
          const eventDate = new Date(event.EventDate);
          return eventDate >= today;
      });
      //const limitedEvents = filteredEvents.slice(0, 4);
      //res.filter(x=>x.Category?.Category==str)
      arr = filteredEvents.slice(0, 4);

    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const fetchUserInformationList = async (sp) => {
  let arr = []
  try {
    const userList = await sp.web.lists.getByTitle("User Information List").items.select("ID", "Title", "EMail", "Department", "JobTitle", "Picture").filter("EMail ne null").top(4)();
    console.log(userList, 'userList');
    arr = userList
    // setUsersArr(userList);
  } catch (error) {
    console.log("Error fetching users:", error);
  }
  return arr;
};
export const fetchPinnedUser = async (sp) => {
  let currentUser;
  await sp.web.currentUser()
    .then(user => {
      currentUser = user.Id; // Get the current user's ID
    })
    .catch(error => {
      console.error("Error fetching current user: ", error);
      return [];
    });
    //select=Pinned/Title,Pinned/SPSPicturePlaceholderState&$expand=Pinned
  let arr = []
  try {
    const userList = await sp.web.lists.getByTitle("ARGPinned").items
      .select("*,Pinned/ID,Pinned/Title,Pinned/EMail,Pinned/Department,Pinned/WorkPhone,Pinned/WorkPhone,Pinned/SPSPicturePlaceholderState")
      .expand("Pinned")
      //.select("ID", "Title", "EMail", "Department", "JobTitle", "Picture")
      .filter(`PinnedById eq ${currentUser}`)
      .top(4)();
    console.log(userList, 'userList');
    arr = userList
    // setUsersArr(userList);
  } catch (error) {
    console.log("Error fetching users:", error);
  }
  return arr;
};
export const fetchComments = async (sp, Id) => {


  let arr = []
  try {
    const userList = await sp.web.lists.getByTitle("ARGAnnouncementandNewsComments").items.select("*,AnnouncementAndNews/Id").expand("AnnouncementAndNews").filter(`AnnouncementAndNewsId eq ${Id}`)();
    console.log(userList, 'userList');
    // arr = userList.length
    userList.forEach(async element => {
      const Likelength = await fetchLikes(sp, element.Id)
      let arrJson = {
        CommentLength: userList.length,
        Likelength: Likelength
      }
      arr.push(arrJson)
    });


    // setUsersArr(userList);
  } catch (error) {
    console.log("Error fetching users:", error);
  }
  return arr;
};
export const fetchLikes = async (sp, Id) => {

  let arr = 0
  try {
    const userList = await sp.web.lists.getByTitle("ARGAnnouncementAndNewsUserLikes").items.select("*,AnnouncementAndNewsComments/Id").expand("AnnouncementAndNewsComments").filter(`AnnouncementAndNewsComments eq ${Id}`)();
    console.log(userList, 'userList');
    arr = userList.length
    // setUsersArr(userList);
  } catch (error) {
    console.log("Error fetching users:", error);
  }
  return arr;
};