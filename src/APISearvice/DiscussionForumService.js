import { forEach } from "lodash";
import Swal from "sweetalert2";
import { DateTime } from "luxon";
export const getDiscussionForum = async (_sp) => {
    let arr = []
    let arr1 = []

    let str = "Announcements";
    let currentUser;
    await _sp.web.currentUser()
        .then(user => {
            currentUser = user.Id; // Get the current user's ID
        })
        .catch(error => {
            console.error("Error fetching current user: ", error);
            return [];
        });

    //   if (!currentUser) return arr; // Return empty array if user fetch failed

    await _sp.web.lists.getByTitle("ARGDiscussionForum")
        .items.select("*,DiscussionForumCategory/Id,DiscussionForumCategory/CategoryName,Author/ID,Author/Title,Author/EMail,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType,Author/SPSPicturePlaceholderState")
        .expand("DiscussionForumCategory,Author,InviteMemebers").orderBy("Created", false).getAll()
        .then((res) => {
            console.log("--discussion", res);
            arr1 = res;
        })
        .catch((error) => {
            console.log("Error fetching data: ", error);
        });
    //item?.Author?.ID == currentUser && 
    // arr = arr1.filter(item =>
    //     item?.Author?.ID != currentUser && item.InviteMemebersId != null ? ((item.InviteMemebersId && item.InviteMemebersId.includes(currentUser))) : (item)
    // );
    arr = arr1.filter(item =>
        item.GroupType == "Private" ? (item?.Author?.ID != currentUser && item.InviteMemebersId != null ? ((item.InviteMemebersId && item.InviteMemebersId.includes(currentUser))) : item?.Author?.ID == currentUser) : (item)
    );
    return arr;
}
export const get7DaysDiscussionForum = async (_sp) => {
    let arr = [];
    let arr1 = [];

    let currentUser;

    // Fetch current user ID
    await _sp.web.currentUser()
        .then(user => {
            currentUser = user.Id; // Get the current user's ID
        })
        .catch(error => {
            console.error("Error fetching current user: ", error);
            return [];
        });

    if (!currentUser) return arr; // Return empty array if user fetch failed

    // Get current date and date 7 days ago
    const today = new Date();
    const lastWeek = new Date();
    lastWeek.setDate(today.getDate() - 7);

    // Format dates in 'yyyy-MM-dd' format for the filter
    const todayStr = today.toISOString()?.split('T')[0]; // 'yyyy-MM-dd'
    const lastWeekStr = lastWeek.toISOString()?.split('T')[0]; // 'yyyy-MM-dd'

    // Fetch all discussion forum data created within the last 7 days (without filtering by current user in InviteMemebersId)
    await _sp.web.lists.getByTitle("ARGDiscussionForum")
        .items.select("*,DiscussionForumCategory/Id,DiscussionForumCategory/CategoryName,Author/ID,Author/Title,Author/EMail,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType")
        .expand("DiscussionForumCategory,Author,InviteMemebers")
        .filter(`Created ge '${lastWeekStr}T00:00:00Z' and Created le '${todayStr}T23:59:59Z'`) // Filter by date range only
        .orderBy("Created", false) // Order by 'Created' field in descending order
        .getAll()
        .then((res) => {
            console.log("--discussion", res);
            arr1 = res;
        })
        .catch((error) => {
            console.log("Error fetching data: ", error);
        });
    //item?.Author?.ID == currentUser &&    
    arr = arr1.filter(item =>
        item.GroupType == "Private" ? (item?.Author?.ID != currentUser && item.InviteMemebersId != null ? ((item.InviteMemebersId && item.InviteMemebersId.includes(currentUser))) : item?.Author?.ID == currentUser) : (item)
    );

    return arr;
}
export const getOldDiscussionForum = async (_sp) => {
    let arr = [];
    let arr1 = [];

    let currentUser;

    // Fetch current user ID
    await _sp.web.currentUser()
        .then(user => {
            currentUser = user.Id; // Get the current user's ID
        })
        .catch(error => {
            console.error("Error fetching current user: ", error);
            return [];
        });

    if (!currentUser) return arr; // Return empty array if user fetch failed

    // Get current date and date 7 days ago
    const today = new Date();
    const lastWeek = new Date();
    lastWeek.setDate(today.getDate() - 7);

    // Format dates in 'yyyy-MM-dd' format for the filter
    const todayStr = today.toISOString()?.split('T')[0]; // 'yyyy-MM-dd'
    const lastWeekStr = lastWeek.toISOString()?.split('T')[0]; // 'yyyy-MM-dd'
    let variableBool = false;
    // Fetch all discussion forum data created within the last 7 days (without filtering by current user in InviteMemebersId)
    await _sp.web.lists.getByTitle("ARGDiscussionForum")
        .items.select("*,DiscussionForumCategory/Id,DiscussionForumCategory/CategoryName,Author/ID,Author/Title,Author/EMail,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType")
        .expand("DiscussionForumCategory,Author,InviteMemebers")
        .filter(`GroupType eq 'Public'`) // Filter by date range only
        .orderBy("Created", false) // Order by 'Created' field in descending order
        .getAll()
        .then((res) => {
            console.log("--discussion old", res);
            arr1 = res;
        })
        .catch((error) => {
            console.log("Error fetching data: ", error);
        });

    arr = arr1
    // .filter(item =>
    //     item?.Author?.ID != currentUser && item.InviteMemebersId != null ? ((item.InviteMemebersId || item.InviteMemebersId.includes(currentUser))) : (item?.Author?.ID == currentUser && item)
    // );

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
export const updateItem = async (itemData, _sp, id) => {
    let resultArr = []
    try {
        const newItem = await _sp.web.lists.getByTitle('ARGDiscussionForum').items.getById(id).update(itemData);
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
export const addItem = async (itemData, _sp) => {
    debugger
    let resultArr = []
    try {
        const newItem = await _sp.web.lists.getByTitle('ARGDiscussionForum').items.add(itemData);

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
export const fetchTrendingDiscussionBasedOn = async (_sp) => {
    let arr = []
    await _sp.web.lists.getByTitle("TrendingDiscussionBasis").items.select("Title").top(1)()
        .then((res) => {
            // arr=res;

            arr = res[0]
        }).catch((error) => {
            console.log("Error fetching data: ", error);
        });
    return arr;
}

export const getDiscussionForumByID = async (_sp, id) => {

    let arr = []
    let arrs = []
    let bannerimg = []
    await _sp.web.lists.getByTitle("ARGDiscussionForum").items.getById(id).select("*", "DiscussionForumCategory/ID", "DiscussionForumCategory/CategoryName", "Entity/ID", "Entity/Entity", "InviteMemebers/Id", "InviteMemebers/Title", "InviteMemebers/EMail", "GroupType").expand("Entity", "DiscussionForumCategory", "InviteMemebers")()
        .then((res) => {
            console.log(res, ' let arrs=[]');
            // const bannerimgobject = res.AnnouncementandNewsBannerImage != "{}" && JSON.parse(res.AnnouncementandNewsBannerImage)
            // console.log(bannerimgobject[0], 'bannerimgobject');

            //bannerimg.push(bannerimgobject);
            const parsedValues = {
                Title: res.Title != undefined ? res.Title : "",
                description: res.Description != undefined ? res.Description : "",
                //   overview: res.Overview != undefined ? res.Overview : "",
                //   IsActive: res.IsActive,
                ID: res.ID,
                Topic: res.Topic,
                Overview: res.Overview,
                // BannerImage: bannerimg,
                GroupType: res.GroupType,
                //   TypeMaster: res?.AnnouncementandNewsTypeMaster?.ID != undefined ? res.AnnouncementandNewsTypeMaster?.ID : "",
                Category: res.DiscussionForumCategory?.ID != undefined ? res.DiscussionForumCategory?.ID : "",
                Entity: res.Entity?.ID != undefined ? res.Entity?.ID : "",
                //FeaturedAnnouncement: res.FeaturedAnnouncement,
                DiscussionForumGalleryJSON: res.DiscussionForumGalleryJSON != null ? JSON.parse(res.DiscussionForumGalleryJSON) : "",
                DiscussionForumDocsJSON: res.DiscussionForumDocsJSON != null ? JSON.parse(res.DiscussionForumDocsJSON) : "",
                DiscussionForumGalleryId: res.DiscussionForumGalleryId,
                DiscussionForumDocsId: res.DiscussionForumDocsId,
                inviteMembers: res?.InviteMemebers.length > 0 ? res?.InviteMemebers.map((ele) => {
                    return {
                        id: ele.Id,
                        name: ele.Title
                    }
                }
                ) : null,
                inviteMembersIds: res?.InviteMemebers.length > 0 ? res?.InviteMemebers.map((ele) => ele.Id)
                    : null
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
export const getDiscussionForumDetailsById = async (_sp, idNum) => {
    let arr = []
    let arr1 = []
    await _sp.web.lists.getByTitle("ARGDiscussionForum").items.getById(idNum).select("*,DiscussionForumCategory/ID,DiscussionForumCategory/CategoryName,Entity/ID,Entity/Entity,InviteMemebers/Id,InviteMemebers/Title,GroupType,Author/ID,Author/Title,Author/EMail,Author/SPSPicturePlaceholderState,ARGDiscussionStatus").expand("Entity,DiscussionForumCategory,InviteMemebers,Author")()
        .then(async (res) => {
            // await _sp.web.lists.getByTitle("ARGDiscussionForum").items.getById(idNum)
            //     .select("*,DiscussionForumCategory/ID,DiscussionForumCategory/CategoryName,Entity/ID,Entity/Entity,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType ,Author/ID,Author/Title,Author/EMail ").expand("Entity,DiscussionForumCategory,InviteMemebers , Author")()
            //     .then((res) => {
            // arr=res;
            if(res.InviteMemebers){
                for (var j = 0; j < res.InviteMemebers.length; j++) {
                  var user = await _sp.web.getUserById(res.InviteMemebers[j].Id)();
                  var profile = await _sp.profiles.getPropertiesFor(`i:0#.f|membership|${user.Email}`);
                  res.InviteMemebers[j].EMail = user.Email;
              
                  res.InviteMemebers[j].SPSPicturePlaceholderState = profile.UserProfileProperties?profile.UserProfileProperties[profile.UserProfileProperties.findIndex(obj=>obj.Key === "SPS-PicturePlaceholderState")].Value:"1";
              
                }
              
               }
            console.log(res, 'hhhjhjh');

            arr1.push(res)
            arr = arr1
        }).catch((error) => {
            console.log("Error fetching data: ", error);
        });
    return arr;
}
export const getDiscussion = async (_sp) => {
    let arr = []
    await _sp.web.lists.getByTitle("ARGDiscussionForum").items.select("*,DiscussionForumCategory/Id,DiscussionForumCategory/CategoryName,Entity/ID,Entity/Entity,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType")
        .expand("DiscussionForumCategory,Entity,InviteMemebers").top(3)()
        .then((res) => {
            // arr=res;

            arr = res
        }).catch((error) => {
            console.log("Error fetching data: ", error);
        });
    return arr;
}
export const getDiscussionMe = async (_sp) => {
    let arr = []
    let currentUser;
    await _sp.web.currentUser()
        .then(user => {
            currentUser = user.Id; // Get the current user's ID
        })
        .catch(error => {
            console.error("Error fetching current user: ", error);
            return [];
        });
    await _sp.web.lists.getByTitle("ARGDiscussionForum").items.select("*,DiscussionForumCategory/Id,DiscussionForumCategory/CategoryName,Entity/ID,Entity/Entity,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType")
        .expand("DiscussionForumCategory,Entity,InviteMemebers").filter(`Author eq ${currentUser}`).top(3)()
        .then((res) => {
            // arr=res;

            arr = res
        }).catch((error) => {
            console.log("Error fetching data: ", error);
        });
    return arr;
}
export const getDiscussionMeAll = async (_sp) => {
    let arr = []
    let currentUser;
    await _sp.web.currentUser()
        .then(user => {
            currentUser = user.Id; // Get the current user's ID
        })
        .catch(error => {
            console.error("Error fetching current user: ", error);
            return [];
        });
    await _sp.web.lists.getByTitle("ARGDiscussionForum").items.select("*,DiscussionForumCategory/Id,DiscussionForumCategory/CategoryName,Entity/ID,Entity/Entity,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType")
        .expand("DiscussionForumCategory,Entity,InviteMemebers").filter(`Author eq ${currentUser}`)()
        .then((res) => {
            // arr=res;

            arr = res
        }).catch((error) => {
            console.log("Error fetching data: ", error);
        });
    return arr;
}
export const getDiscussionFilterAll = async (_sp, filterOption) => {
    let arr = [];
    let filterQuery = "";
    let currentUser;
    let currentUserName;

    await _sp.web.currentUser()
        .then(user => {
            currentUser = user.Id; // Get the current user's ID
            currentUserName = user.Title
        })
    await _sp.web.lists.getByTitle("ARGDiscussionForum")
        .items
        .select("*,DiscussionForumCategory/Id,DiscussionForumCategory/CategoryName,Entity/ID,Entity/Entity,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType")
        .expand("DiscussionForumCategory,Entity,InviteMemebers")
        .filter(`AuthorId eq ${currentUser} `)
        ()
        .then((res) => {
            arr = res;
        }).catch((error) => {
            console.log("Error fetching data: ", error);
        });

    return arr;
};
export const getDiscussionFilter = async (_sp) => {
    let arr = [];
    let filterQuery = "";

    // Get current date and times for filtering
    const today = DateTime.now().startOf('day');
    const yesterday = today.minus({ days: 1 });
    const lastWeek = today.minus({ weeks: 1 });
    const lastMonth = today.minus({ months: 1 });
    const filtervalue = await fetchTrendingDiscussionBasedOn(_sp);

    // Build the filter query based on the selected option
    switch (filtervalue.Title) {
        case "Today":
            filterQuery = `Created ge datetime'${today.toISO()}'`;
            break;
        case "Yesterday":
            filterQuery = `Created ge datetime'${yesterday.toISO()}' and Created lt datetime'${today.toISO()}'`;
            break;
        case "Last Week":
            filterQuery = `Created ge datetime'${lastWeek.toISO()}' and Created lt datetime'${today.toISO()}'`;
            break;
        case "Month":
            filterQuery = `Created ge datetime'${lastMonth.toISO()}' and Created lt datetime'${today.toISO()}'`;
            break;
        default:
            filterQuery = ""; // No filter
    }

    await _sp.web.lists.getByTitle("ARGDiscussionForum")
        .items
        .select("*,DiscussionForumCategory/Id,DiscussionForumCategory/CategoryName,Entity/ID,Entity/Entity,InviteMemebers/Id,InviteMemebers/Title,InviteMemebers/EMail,GroupType")
        .expand("DiscussionForumCategory,Entity,InviteMemebers")
        .filter(filterQuery)
        .top(3)()
        .then(async (res) => {
            //arr = res;
            for (let i = 0; i < res.length > 0; i++) {
                const comment = await getDiscussionCommentsByID(_sp, res[i].ID)
                let arrs = {
                    Topic: res[i].Topic,
                    Overview: res[i].Overview,
                    DiscussionForumCategory: res[i].DiscussionForumCategory,
                    CommentLength: comment,
                    ID: res[i].ID
                }
                arr.push(arrs);
            }
            arr.sort((a, b) => b.CommentLength - a.CommentLength);

        }).catch((error) => {
            console.log("Error fetching data: ", error);
        });

    return arr;
};
export const getDiscussionCommentsByID = async (sp, Id) => {
    let arr = [];
    let arrLength = 0
    let arrUsers = []
    let CreatedDate = ""
    await sp.web.lists.getByTitle("ARGDiscussionComments").items.select("*,Author/Id,Author/Title,Author/EMail").expand("Author").filter(`DiscussionForumId eq ${Id}`)().then(res => {

        let arrs;
        console.log(res, 'resres');

        arrLength = res.length

    }
    )
    return arrLength;
}
export const getDiscussionComments = async (sp, Id) => {
    let arr = [];
    let arrLength = 0
    let arrUsers = []
    let CreatedDate = ""
    await sp.web.lists.getByTitle("ARGDiscussionComments").items.select("*,Author/Id,Author/Title,Author/EMail").expand("Author").filter(`DiscussionForumId eq ${Id}`)().then(res => {

        let arrs;
        console.log(res, 'resres');

        arrLength = res.length
        for (let i = 0; i < res.length > 0; i++) {
            let arrs = {
                arrUser: res[i].Author.EMail,

            }
            arrUsers.push(arrs)
        }

        const uniqueArray = arrUsers.filter((item, index) => arrUsers.indexOf(item) === index);

        console.log(uniqueArray);

        var lengthcount = res.length - 1
        CreatedDate = res.length > 0 ? res[res.length - 1].Created : "NA"
        const totalLikes = res.reduce((total, res1) => total + ((res1.UserLikesJSON != "" && res1.UserLikesJSON != null && res1.UserLikesJSON != undefined) && JSON.parse(res1.UserLikesJSON).length || 0), 0);

        const totalRepliesCount = res.reduce(
            (total, res1) => total + ((res1.UserCommentsJSON != "" && res1.UserCommentsJSON != null && res1.UserCommentsJSON != undefined) && JSON.parse(res1.UserCommentsJSON)?.length || 0),
            0
        );
        let mainArray =
        {
            arrLength: res.length,
            arrUser: uniqueArray,
            CreatedDate: CreatedDate,
            totalLikes: totalLikes,
            totalRepliesCount: totalRepliesCount
        }
        arr = mainArray
    }
    )
    return arr;
}
export const getChoiceFieldOption = async (_sp, listName, fieldName) => {
    let arr = []
    const field2 = await _sp.web.lists.getByTitle(listName).fields.getByInternalNameOrTitle(fieldName)()
    console.log(field2, 'field2');

    arr = field2["Choices"]

    return arr;
}   