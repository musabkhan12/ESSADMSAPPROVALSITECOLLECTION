import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { getSP } from '../webparts/essaApproval/loc/pnpjsConfig';
import { SPFI } from '@pnp/sp';
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
const sp = getSP();
let web = "";
// import { sp } from "@pnp/sp/presets/all";  // Import PnPjs
// import { SPFetchClient } from "@pnp/sp/authenticators";
// sp.configure({
//   sp: {
//     baseUrl: siteUrl,
//   },
// });
export const fetchAutomationDepartment = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("ARGAutomationDepartment").items.getAll().then((res) => {
    console.log(res);

    //res.filter(x=>x.Category?.Category==str)
    arr = res;
  })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const addItem = async (itemData, _sp) => {

  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGBusinessApps').items.add(itemData);

    console.log('Item added successfully:', newItem);
    // Swal.fire('Item added successfully', '', 'success');

    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
    Swal.fire(' Cancelled', '', 'error')
  }
  return resultArr;
};
export const getCategory = async (_sp, id) => {
  let arr = []
  await _sp.web.lists.getByTitle("BusinessAppsCategory").items.select("ID,CategoryName").expand("").filter(`(Active eq 1) and(AnnouncementandNewsTypeMaster/ID eq ${id})`)()
    .then((res) => {
      console.log(res);
      const newArray = res.map(({ ID, CategoryName }) => ({ id: ID, name: CategoryName }));
      console.log(newArray, 'newArray');

      arr = newArray;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const fetchAutomationCategory = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("BusinessAppsCategory").items.getAll().then((res) => {
    console.log(res);

    //res.filter(x=>x.Category?.Category==str)
    arr = res;
  })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}

// const getUserGroups = async (context) => {
//   try {
//     if (!context) {
//       console.error("SPFx context is undefined!");
//       return [];
//     }
//     console.log("Context initialized successfully.");

//     // Initialize MSGraphClientV3
//     const graphClient = await context.msGraphClientFactory.getClient('3');
//     console.log("Graph client initialized successfully.");

//     // Fetch groups using memberOf (Azure AD Groups)
//     const response = await graphClient.api("/me/memberOf").get();
//     console.log("User Groups:", response.value);

//     // Extract Azure AD Group IDs
//     const userGroupIds = response.value.map(group => group.id);
//     console.log("User Group Azure AD IDs:", userGroupIds);

//     return userGroupIds;
//   } catch (error) {
//     console.error("Error fetching user groups:", error);
//     return [];
//   }
// };
const getUserGroups = async (_sp , context) => {
  try {
    if (!context) {
      console.error("SPFx context is undefined!");
      return { adGroupIds: [], spGroupTitles: [] };
    }
    console.log("Context initialized successfully.");

    // Initialize MSGraphClientV3
    const graphClient = await context.msGraphClientFactory.getClient('3');
    console.log("Graph client initialized successfully.");

    // Fetch Azure AD Groups using Microsoft Graph API
    const response = await graphClient.api("/me/memberOf").get();
    console.log("User Azure AD Groups:", response.value);

    // Extract Azure AD Group IDs
    const userAdGroupIds = response.value.map(group => group.id);
    console.log("User Azure AD Group IDs:", userAdGroupIds);

    // const userAdGroupDetails = await Promise.all(
    //   userAdGroupIds.map(async group => {
    //     try {
    //       const ownersResponse = await graphClient.api(`/groups/${group}/owners`).get();
    //       const owners = ownersResponse.value.map(owner => owner.id);
    //       return { ...group, owners };
    //     } catch (error) {
    //       console.error(`Error fetching owners for group ${group}:`, error);
    //       return { ...group, owners: [] }; // Return empty owners if the call fails
    //     }
    //   })
    // );

    // console.log("User Azure AD Group Details with Owners:", userAdGroupDetails);


    // Fetch SharePoint Groups using REST API
      // Fetch SharePoint Groups using PnP JS
      const spGroupsResponse = await _sp.web.currentUser.groups();
      // console.log("User SharePoint Groups:", spGroupsResponse);
  
      // Extract SharePoint Group Titles
      const userSpGroupTitles = spGroupsResponse.map(group => group.Title);
      console.log("User SharePoint Group Titles:", userSpGroupTitles);
  
      return { adGroupIds: userAdGroupIds, spGroupTitles: userSpGroupTitles };
  } catch (error) {
    console.error("Error fetching user groups:", error);
    return { adGroupIds: [], spGroupTitles: [] };
  }
};

const normalizeGroupName = (groupName) => {
  return groupName?.replace(/\s*Members\s*$/i, "").trim();
};

const getAzureADGroupIds = async (groupNames, graphClient) => {
  try {
    if (!groupNames.length) return {};

    console.log("Group Names to Fetch from Azure AD:", groupNames);

    // Normalize group names
    const normalizedGroupNames = groupNames
    // .map(name => normalizeGroupName(name));
    // console.log("Normalized Group Names:", normalizedGroupNames);

    // Construct Graph API filter query for multiple group names
    // const filterQuery = normalizedGroupNames.map(name => `displayName eq '${name}'`).join(" or ");
    // console.log("Filter Query for Azure AD:", filterQuery);

    // const response = await graphClient.api(`/groups?$filter=${filterQuery}`).get();
    const filterQuery = groupNames
    .map(name => `displayName eq '${name.replace(/'/g, "''")}'`) // Escape single quotes
    .join(" or ");
    const encodedFilterQuery = encodeURIComponent(filterQuery);
    console.log("Filter Query for Azure AD:", filterQuery);
    console.log("Filter Query for Azure AD:", encodedFilterQuery);
    // const response = await graphClient.api(`/groups?$filter=${filterQuery}`).get();
    const response = await graphClient.api(`/groups?$filter=${encodedFilterQuery}`).get();

    console.log("Azure AD Group Response:", response.value);
    

    // Map Display Names to Azure AD Group IDs
    const groupIdMapping = response.value.reduce((acc, group) => {
      acc[group.displayName] = group.id;
      return acc;
    }, {});

    console.log("Mapped Azure AD Group IDs:", groupIdMapping);
    return groupIdMapping;
  } catch (error) {
    console.error("Error fetching Azure AD Group IDs:", error);
    return {};
  }
};

export const fetchARGAutomationdata = async (_sp, context) => {
  let arr = [];

  if (!context) {
    console.error("SPFx context is undefined!");
    return [];
  }

  // Get user groups from Azure AD
  // const userGroupIds = await getUserGroups(context);
  // Get user groups from Azure AD & SharePoint
const { adGroupIds: userGroupIds, spGroupTitles: userSpGroupTitles } = await getUserGroups(_sp, context);


  // Get Graph Client
  const graphClient = await context.msGraphClientFactory.getClient('3');

  // Fetch ARGBusinessApps list items from SharePoint
  const res = await _sp.web.lists.getByTitle("ARGBusinessApps").items
    .select("*,Category/Id,Category/CategoryName,EnableAudienceTargeting,Audience/Id,Audience/Title")
    .expand("Category,Audience")
    .filter("IsActive eq 1")
    .orderBy("Order0", true)
    .getAll()
    .catch(error => {
      console.log("Error fetching data: ", error);
      return [];
    });

  if (res.length > 0) {
    // Extract valid Audience Entries
  // Extract valid Audience Entries (since Audience allows only ONE selection)
// const audienceGroupNames = [...new Set(
//   res.map(item => {
//     console.log("Audience Data:", item.Audience);
//     console.log("Audience Data Type:", typeof item.Audience);

//     if (!item.Audience) return null; // Skip if null/undefined

//     return item.Audience.Title; // Directly extract the group title (single selection)
//   }).filter(title => title) // Remove null values
// )];
const audienceGroupNames = [...new Set(
  res.flatMap(item => {
    console.log("Audience Data:", item.Audience);
    console.log("Audience Data Type:", typeof item.Audience);

    if (!item.Audience || !Array.isArray(item.Audience)) return []; // Skip if null/undefined or not an array

    return item.Audience.map(audience => audience.Title); // Extract titles from the Audience array
  })
)];


    console.log("Audience Office 365 & SP Group Names:", audienceGroupNames);

    // Get corresponding Azure AD Group IDs
    const azureADGroupMapping = await getAzureADGroupIds(audienceGroupNames, graphClient);

    console.log("Mapped Azure AD Group Name to IDs:", azureADGroupMapping);
 // Print Type of Each Group
 audienceGroupNames.forEach(groupName => {
  if (azureADGroupMapping[groupName]) {
    console.log(`Group "${groupName}" is an Office 365 / Security Group (Azure AD)`);
  } else {
    console.log(`Group "${groupName}" is a SharePoint Group or a User`);
  }
});
    // Filter based on user group membership
    arr = res.filter(item => {
      if (!item.EnableAudienceTargeting) return true; // Include if targeting is disabled

      // Extract Audience Group Titles
      // if (!item.Audience) return false;
      if (!Array.isArray(item.Audience) || item.Audience.length === 0) return false;
      // const audienceGroupNames = item.Audience ? [item.Audience.Title] : [];
      const audienceGroupNames = [...new Set(
        item.Audience?.map(audience => audience.Title) || [] // Map over the Audience array to extract Titles
      )];


      console.log("Checking Item:", item.SubTitle);
      console.log("User Group IDs:", userGroupIds);
      console.log("Audience Group Names:", audienceGroupNames);

      // Normalize Audience Group Names
      const normalizedAudienceGroupNames = audienceGroupNames
      // .map(name => normalizeGroupName(name));

      // Convert Audience Titles to Azure AD Group IDs (if they exist)
      const audienceGroupIds = normalizedAudienceGroupNames
        .map(name => azureADGroupMapping[name]) // Map to AD IDs
        .filter(id => id); // Remove undefined values

      console.log("Mapped Audience Group IDs:", audienceGroupIds);

      // Check if the user is a member of any mapped Azure AD group
      const isMemberOfAzureADGroup = audienceGroupIds.some(groupId => userGroupIds.includes(groupId));

      // If not an Azure AD Group, assume it's a SharePoint Group (fallback mechanism)
      // const isMemberOfSharePointGroup = normalizedAudienceGroupNames.some(name =>
      //   userGroupIds.includes(name) // Assuming SharePoint groups are stored in userGroupIds
      // );
      const isMemberOfSharePointGroup = normalizedAudienceGroupNames.some(name =>
        userSpGroupTitles.includes(name) // Assuming SharePoint groups are stored in userGroupIds
      );

      // return isMemberOfAzureADGroup || isMemberOfSharePointGroup;
      return isMemberOfAzureADGroup 
      || isMemberOfSharePointGroup 
      
      || 
      (item.Audience && item.Audience.some(audience => audience.Id === context.pageContext.legacyPageContext.userId));

      // || 
      // item.Audience.some(audience => 
      //   SiteUserPeople.some(user => user.Id === audience.Id && user.Title === audience.Title)
      // )
      
      // item.Audience.some(audience => userSpGroupTitles.includes(audience.Title)); // Check if user is in SharePoint Group

  // userSpGroupTitles.includes(item.Audience.Title)

    });
  }

  console.log("Filtered response-->>>", arr);
  return arr;
};

//////////// latest updated code for office 365 ggroups and security group
// const getUserGroups = async (context) => {
//   try {
//     if (!context) {
//       console.error("SPFx context is undefined!");
//       return [];
//     }
//     console.log("Context initialized successfully.");

//     // Initialize MSGraphClientV3
//     const graphClient = await context.msGraphClientFactory.getClient('3');
//     console.log("Graph client initialized successfully.");

//     // Fetch groups using memberOf
//     const response = await graphClient.api("/me/memberOf").get();
//     console.log("User Groups:", response.value);

//     // Extract Azure AD Group IDs (not SharePoint group names)
//     const userGroupIds = response.value.map(group => group.id);
//     console.log("User Group Azure AD IDs:", userGroupIds);

//     return userGroupIds; // Returns an array of Azure AD Group IDs
//   } catch (error) {
//     console.error("Error fetching user groups:", error);
//     return [];
//   }
// };

// const normalizeGroupName = (groupName) => {
//   return groupName.replace(/\s*Members\s*$/i, "").trim();
// };


// const getAzureADGroupIds = async (groupNames, graphClient) => {
//   try {
//     if (!groupNames.length) return [];
    
//     console.log("Group Names to Fetch from Azure AD:", groupNames);

//     // Normalize group names
//     const normalizedGroupNames = groupNames.map(name => normalizeGroupName(name));
//     console.log("Normalized Group Names:", normalizedGroupNames);

//     // Construct Graph API filter query for multiple group names
//     const filterQuery = normalizedGroupNames.map(name => `displayName eq '${name}'`).join(" or ");
//     console.log("Filter Query for Azure AD:", filterQuery);

//     const response = await graphClient.api(`/groups?$filter=${filterQuery}`).get();

//     console.log("Azure AD Group Response:", response.value);

//     // Map Display Names to IDs
//     const groupIdMapping = response.value.reduce((acc, group) => {
//       acc[group.displayName] = group.id;
//       return acc;
//     }, {});

//     console.log("Mapped Azure AD Group IDs:", groupIdMapping);
//     return groupIdMapping; // Returns an object { "Group Name": "Azure AD Group ID" }
//   } catch (error) {
//     console.error("Error fetching Azure AD Group IDs:", error);
//     return {};
//   }
// };

// export const fetchARGAutomationdata = async (_sp, context) => {
//   let arr = [];

//   if (!context) {
//     console.error("SPFx context is undefined!");
//     return [];
//   }

//   // Get user groups from Azure AD
//   const userGroupIds = await getUserGroups(context);
  
//   // Get Graph Client
//   const graphClient = await context.msGraphClientFactory.getClient('3');

//   // Fetch ARGBusinessApps list items from SharePoint
//   const res = await _sp.web.lists.getByTitle("ARGBusinessApps").items
//     .select("*,Category/Id,Category/CategoryName,EnableAudienceTargeting,Audience/Id,Audience/Title")
//     .expand("Category,Audience")
//     .filter("IsActive eq 1")
//     .orderBy("Order0", true)
//     .getAll()
//     .catch(error => {
//       console.log("Error fetching data: ", error);
//       return [];
//     });

//   if (res.length > 0) {
//     // Extract Office 365 Group Names from Audience column

//     // const audienceGroupNames = [...new Set(res.flatMap(item =>
//     //   item.Audience?.map(group => group.Title) || [])
    
//     // )];
//     const audienceGroupNames = [...new Set(res.flatMap(item => {
//       console.log("Audience Data:", item.Audience);
//       console.log("Audience Data type:", typeof item.Audience);
      
//       return Array.isArray(item.Audience) ? item.Audience.map(group => group.Title) : [];
//   }))];
  

//     console.log("Audience Office 365 Group Names:", audienceGroupNames);

//     // Get corresponding Azure AD Group IDs
//     const azureADGroupMapping = await getAzureADGroupIds(audienceGroupNames, graphClient);

//     console.log("Mapped Azure AD Group Name to IDs:", azureADGroupMapping);

//     // Filter based on user group membership
//     arr = res.filter(item => {
//       if (!item.EnableAudienceTargeting) return true; // Include if targeting is disabled

//       // Extract Audience Group Titles
//       const audienceGroupNames = item.Audience?.map(group => group.Title) || [];

//       console.log("Checking Item:", item.SubTitle);
//       console.log("User Group IDs:", userGroupIds);
//       console.log("Audience Group Names:", audienceGroupNames);

//       // Normalize Audience Group Names
//       const normalizedAudienceGroupNames = audienceGroupNames.map(name => normalizeGroupName(name));

//       // Convert Audience Titles to Azure AD Group IDs
//       const audienceGroupIds = normalizedAudienceGroupNames.map(name => azureADGroupMapping[name]).filter(id => id);

//       console.log("Mapped Audience Group IDs:", audienceGroupIds);

//       // Check if the user is a member of any mapped Azure AD group
//       return audienceGroupIds.some(groupId => userGroupIds.includes(groupId));
//     });
//   }

//   console.log("Filtered response-->>>", arr);
//   return arr;
// };


/////////




/////////// working code for office 365 ggroups 
// const getUserGroups = async (context) => {
//   try {
//     if (!context) {
//       console.error("SPFx context is undefined!");
//       return [];
//     }
//     console.log("Context initialized successfully.");

//     // Initialize MSGraphClientV3
//     const graphClient = await context.msGraphClientFactory.getClient('3');
//     console.log("Graph client initialized successfully.");

//     // Fetch groups using memberOf
//     const response = await graphClient.api("/me/memberOf").get();
//     console.log("User Groups:", response.value);

//     // Extract Azure AD Group IDs (not SharePoint group names)
//     const userGroupIds = response.value.map(group => group.id);
//     console.log("User Group Azure AD IDs:", userGroupIds);

//     return userGroupIds; // Returns an array of Azure AD Group IDs
//   } catch (error) {
//     console.error("Error fetching user groups:", error);
//     return [];
//   }
// };

// const getAzureADGroupIds = async (groupNames, graphClient) => {
//   try {
//     if (!groupNames.length) return [];
    
//     console.log("Group Names to Fetch from Azure AD:", groupNames);

//     // Construct Graph API filter query for multiple group names
//     const filterQuery = groupNames.map(name => `displayName eq '${name}'`).join(" or ");
//     console.log("Filter Query for Azure AD:", filterQuery);

//     const response = await graphClient.api(`/groups?$filter=${filterQuery}`).get();

//     console.log("Azure AD Group Response:", response.value);

//     // Map Display Names to IDs
//     const groupIdMapping = response.value.reduce((acc, group) => {
//       acc[group.displayName] = group.id;
//       return acc;
//     }, {});

//     console.log("Mapped Azure AD Group IDs:", groupIdMapping);
//     return groupIdMapping; // Returns an object { "Group Name": "Azure AD Group ID" }
//   } catch (error) {
//     console.error("Error fetching Azure AD Group IDs:", error);
//     return {};
//   }
// };

// export const fetchARGAutomationdata = async (_sp, context) => {
//   let arr = [];

//   if (!context) {
//     console.error("SPFx context is undefined!");
//     return [];
//   }

//   // Get user groups from Azure AD
//   const userGroupIds = await getUserGroups(context);
  
//   // Get Graph Client
//   const graphClient = await context.msGraphClientFactory.getClient('3');

//   // Fetch ARGBusinessApps list items from SharePoint
//   const res = await _sp.web.lists.getByTitle("ARGBusinessApps").items
//     .select("*,Category/Id,Category/CategoryName,EnableAudienceTargeting,Audience/Id,Audience/Title")
//     .expand("Category,Audience")
//     .filter("IsActive eq 1")
//     .orderBy("Order0", true)
//     .getAll()
//     .catch(error => {
//       console.log("Error fetching data: ", error);
//       return [];
//     });

//   if (res.length > 0) {
//     // Extract Office 365 Group Names from Audience column
//     const audienceGroupNames = [...new Set(res.flatMap(item => item.Audience?.map(group => group.Title) || []))];

//     console.log("Audience Office 365 Group Names:", audienceGroupNames);

//     // Get corresponding Azure AD Group IDs
//     const azureADGroupMapping = await getAzureADGroupIds(audienceGroupNames, graphClient);

//     console.log("Mapped Azure AD Group Name to IDs:", azureADGroupMapping);

//     // Filter based on user group membership
//     arr = res.filter(item => {
//       if (!item.EnableAudienceTargeting) return true; // Include if targeting is disabled

//       // Extract Audience Group Titles
//       const audienceGroupNames = item.Audience?.map(group => group.Title) || [];

//       console.log("Checking Item:", item.SubTitle);
//       console.log("User Group IDs:", userGroupIds);
//       console.log("Audience Group Names:", audienceGroupNames);

//       // Convert Audience Titles to Azure AD Group IDs
//       const audienceGroupIds = audienceGroupNames.map(name => azureADGroupMapping[name]).filter(id => id);

//       console.log("Mapped Audience Group IDs:", audienceGroupIds);

//       // Check if the user is a member of any mapped Azure AD group
//       return audienceGroupIds.some(groupId => userGroupIds.includes(groupId));
//     });
//   }

//   console.log("Filtered response-->>>", arr);
//   return arr;
// };
///////////
///////
// const getUserGroups = async (context) => {
//   try {
//     if (!context) {
//       console.error("SPFx context is undefined!");
//       return [];
//     }
//     console.log("Context initialized successfully.");

//     // Initialize MSGraphClientV3
//     // const graphClient: MSGraphClientV3 = await context.msGraphClientFactory.getClient('3');
//     const graphClient = await context.msGraphClientFactory.getClient('3');
//     console.log("Graph client initialized successfully.");

//     // Fetch groups using memberOf
//     const response = await graphClient.api("/me/memberOf").get();
//     console.log("User Groups:", response.value);

//     // Extract Group IDs
//     const userGroupIds = response.value.map(group => group.id);
//     console.log("User Group IDs:", userGroupIds);

//     return userGroupIds; // Returns an array of group IDs
//   } catch (error) {
//     console.error("Error fetching user groups:", error);
//     return [];
//   }
// };
// const getAzureADGroupIds = async (groupNames, graphClient) => {
//   try {
//     if (!groupNames.length) return [];
//       console.log("Group Names:", groupNames);
//     // Construct Graph API filter query for multiple group names
//     const filterQuery = groupNames.map(name => `displayName eq '${name}'`).join(" or ");
//     console.log("Filter Query:", filterQuery);
//     const response = await graphClient.api(`/groups?$filter=${filterQuery}`).get();

//     console.log("Azure AD Group Response:", response.value);

//     // Extract Azure AD Group IDs
//     return response.value.map(group => group.id);
//   } catch (error) {
//     console.error("Error fetching Azure AD Group IDs:", error);
//     return [];
//   }
// };

// export const fetchARGAutomationdata = async (_sp, context) => {
//   let arr = [];

//   if (!context) {
//     console.error("SPFx context is undefined!");
//     return [];
//   }

//   // Get user groups from Azure AD
//   const userGroupIds = await getUserGroups(context);
  
//   // Get Graph Client
//   const graphClient = await context.msGraphClientFactory.getClient('3');

//   // Fetch ARGBusinessApps list items from SharePoint
//   const res = await _sp.web.lists.getByTitle("ARGBusinessApps").items
//     .select("*,Category/Id,Category/CategoryName,EnableAudienceTargeting,Audience/Id,Audience/Title")
//     .expand("Category,Audience")
//     .filter("IsActive eq 1")
//     .orderBy("Order0", true)
//     .getAll()
//     .catch(error => {
//       console.log("Error fetching data: ", error);
//       return [];
//     });

//   if (res.length > 0) {
//     // Extract Office 365 Group Names from Audience column
//     const audienceGroupNames = [...new Set(res.flatMap(item => item.Audience?.map(group => group.Title) || []))];

//     console.log("Audience Office 365 Group Names:", audienceGroupNames);

//     // Get corresponding Azure AD Group IDs
//     const azureADGroupIds = await getAzureADGroupIds(audienceGroupNames, graphClient);

//     console.log("Mapped Azure AD Group IDs:", azureADGroupIds);

//     // Filter based on user group membership
//     arr = res.filter(item => {
//       if (!item.EnableAudienceTargeting) return true; // Include if targeting is disabled

//       // Extract Audience Group IDs
//       const audienceGroupIds = item.Audience?.map(group => group.Title) || [];

//       console.log("Checking Item:", item.SubTitle);
//       console.log("User Group IDs:", userGroupIds);
//       console.log("Audience Group Names:", audienceGroupIds);

//       // Check if the user is a member of any mapped Azure AD group
//       return audienceGroupIds.some(groupName => {
//         const groupId = azureADGroupIds.find(id => groupName === id);
//         return userGroupIds.includes(groupId);
//       });
//     });
//   }

//   console.log("Filtered response-->>>", arr);
//   return arr;
// };

/////

// export const fetchARGAutomationdata = async (_sp, context) => {
//   let arr = [];

//   // Fetch user group memberships (Office 365 Groups the user is part of)
//   const userGroupIds = await getUserGroups(context);

//   // Fetch ARGBusinessApps list items
//   const res = await _sp.web.lists.getByTitle("ARGBusinessApps").items
//     .select("*,Category/Id,Category/CategoryName,EnableAudienceTargeting,Audience/Id,Audience/Title")
//     .expand("Category,Audience")
//     .filter("IsActive eq 1")
//     .orderBy("Order0", true)
//     .getAll()
//     .catch((error) => {
//       console.error("Error fetching data: ", error);
//       return [];
//     });
//      console.log("res-->>>", res);
//   if (res.length > 0) {
//     arr = res.filter(item => {
//       // If EnableAudienceTargeting is disabled, include the item
//       if (!item.EnableAudienceTargeting) return true;

//       // Extract Audience Group IDs from the Audience column
//       // const audienceGroupIds = item.Audience?.map(group => String(group.Id ) , console.log(group.Id , "group.Id of Audiance") )|| [];
//       const audienceGroupIds = item.Audience?.map(group => {
//         console.log(group.Id, "group.Id of Audience");
//         return String(group.Id);
//       }) || [];
//       // Debugging logs
//       console.log("Checking Item:", item.SubTitle);
//       console.log("User Group IDs:", userGroupIds);
//       console.log("Audience Group IDs:", audienceGroupIds);

//       // Include the item only if the user belongs to at least one of the Audience groups
//       return audienceGroupIds.some(groupId => userGroupIds.includes(groupId));
//     });
//   }

//   console.log("Filtered response-->>>", arr);
//   return arr;
// };

// export const fetchARGAutomationdata = async (_sp, context) => {
//   let arr = [];
//   const userGroupIds = await getUserGroups(context); // Fetch user groups

//   // Fetch ARGBusinessApps list items
//   const res = await _sp.web.lists.getByTitle("ARGBusinessApps").items
//     .select("*,Category/Id,Category/CategoryName,EnableAudienceTargeting,Audience/Id,Audience/Title")
//     .expand("Category,Audience")
//     .filter("IsActive eq 1")
//     .orderBy("Order0", true)
//     .getAll()
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//       return [];
//     });

//   // if (res.length > 0) {
//   //   arr = res.filter(item => {
//   //     if (!item.EnableAudienceTargeting) return true; // Include if targeting is disabled
//   //     console.log("User Group IDs:", userGroupIds, item.EnableAudienceTargeting , item.SubTitle);

//   //     // Extract Audience Group IDs (Office 365 Groups)
//   //     const audienceGroupIds = item.Audience?.map(group => group.Id) || [];
//   //     console.log("audienceGroupIds IDs:", audienceGroupIds, item.EnableAudienceTargeting , item.SubTitle);

//   //     // Check if the user is a member of any audience group
//   //     return audienceGroupIds.some(groupId => userGroupIds.includes(groupId));
//   //   });
//   // }
//   if (res.length > 0) {
//     arr = res.filter(item => {
//       if (!item.EnableAudienceTargeting) return true; // Include if targeting is disabled
  
//       // Extract Audience Group IDs
//       const audienceGroupIds = item.Audience?.map(group => String(group.Id)) || [];
  
//       // Debugging logs
//       console.log("Checking Item:", item.SubTitle);
//       console.log("User Group IDs:", userGroupIds);
//       console.log("Audience Group IDs:", audienceGroupIds);
  
//       // Return true if user belongs to any of the Audience groups
//       return audienceGroupIds.some(groupId => userGroupIds.includes(groupId));
//     });
//   }
  
//   console.log("Filtered response-->>>", arr);
//   return arr;
// };
// export const fetchARGAutomationdata = async (_sp, context) => {
//   let arr = []

//   await _sp.web.lists.getByTitle("ARGBusinessApps").items.select("*,Category/Id,Category/CategoryName")
//     .expand("Category")
//     .orderBy("Order0", true)
//     .getAll().then((res) => {
//       console.log("response-->>>", res);

//       //res.filter(x=>x.Category?.Category==str)
//       arr = res;
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }

// export const getMyApprovalsdata = async (_sp,listName,status) => {
//   let arr = []
//   let currentUser;
//   await _sp.web.currentUser()
//     .then(user => {
//       console.log("user",user);
//       currentUser = user.Email; 
//       // Get the current user's Email
//     })
//     .catch(error => {
//       console.error("Error fetching current user: ", error);
//       return [];
//     });

//   if (!currentUser) return arr; 
//   // Return empty array if user fetch failed

//   await _sp.web.lists.getByTitle(listName).items
//     .select("*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail").expand("Author,AssignedTo")
//     .filter(`AssignedTo/EMail eq '${currentUser}' and Status eq '${status}'`)      
//     .orderBy("Created", false).getAll()
//     .then((res) => {
//       console.log(`--MyApproval${listName}`, res);
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
//Add Business Apps


// Initialize PnPjs with the current site collection


// Function to fetch data from a single SharePoint site collection

// export const fetchARGAutomationdata = async (_sp) => {
//   let arr = []

//   await _sp.web.lists.getByTitle("ARGBusinessApps").items.select("*,Category/Id,Category/CategoryName")
//     .expand("Category")
//     .orderBy("Order0", true)
//     .getAll().then((res) => {
//       console.log("response-->>>", res);

//       //res.filter(x=>x.Category?.Category==str)
//       arr = res;
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }

// export const getMyApprovalsdata = async (_sp,listName,status) => {
//   let arr = []
//   let currentUser;
//   await _sp.web.currentUser()
//     .then(user => {
//       console.log("user",user);
//       currentUser = user.Email; 
//       // Get the current user's Email
//     })
//     .catch(error => {
//       console.error("Error fetching current user: ", error);
//       return [];
//     });

//   if (!currentUser) return arr; 
//   // Return empty array if user fetch failed

//   await _sp.web.lists.getByTitle(listName).items
//     .select("*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail").expand("Author,AssignedTo")
//     .filter(`AssignedTo/EMail eq '${currentUser}' and Status eq '${status}'`)      
//     .orderBy("Created", false).getAll()
//     .then((res) => {
//       console.log(`--MyApproval${listName}`, res);
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
//Add Business Apps


// Initialize PnPjs with the current site collection


// Function to fetch data from a single SharePoint site collection




export const getListDataFromSiteCollection = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
  // Setup PnPJs for a specific site collection URL
  let arr = [];
  console.log("sdsssss", sp, _sp)
  
  let apiUrl;
  let FinalStatus = "";
  if (status == "Pending") {
    FinalStatus = "Not Started"
  } else if (status == "Approved") {
    FinalStatus = "Completed"
  }
  if (Actingfor != null && Actingfor != undefined && Actingfor != "") {
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail&$expand=Requestor_x0020_Name,AssignedTo&$filter=${`AssignedTo/EMail eq '${Actingfor}' and TaskStatus eq '${FinalStatus}'`}`;
  } else {
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
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail&$expand=Requestor_x0020_Name,AssignedTo&$filter=${`AssignedTo/EMail eq '${currentUser}' and TaskStatus eq '${FinalStatus}'`}`;
  }
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

export const getDataFromMultipleSites = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
  const allData = [];
  const siteUrls = [
    SiteBaseURL
  ];

  // Loop through each site collection URL and fetch data
  if (portal == "Others"){
    for (const siteUrl of siteUrls) {
      const data = await getListDataFromSiteCollection(_sp, listName, status, Actingfor, portal, siteUrl);
      allData.push(...data);
      console.log("dadadadadad", data);
    }
  } else{
    for (const siteUrl of siteUrls) {
      const data = await getMyApprovalsdata(_sp, listName, status, Actingfor, portal, siteUrl);
      allData.push(...data);
      console.log("dadadadadad", data);
    }
  }


  console.log('Combined data from all site collections:', allData);
  return allData;
}

// Call the function to get the data

export const getMyApprovalsdata = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
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
  if (Actingfor != null && Actingfor != undefined && Actingfor != "") {
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail,RequestedBy/Title,RequestedBy/EMail,RequestedBy/ID&$expand=Author,AssignedTo,RequestedBy&$filter=${`AssignedTo/EMail eq '${Actingfor}' and Status eq '${FinalStatus}'`}`;
  } else {
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
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail,RequestedBy/Title,RequestedBy/EMail,RequestedBy/ID&$expand=Author,AssignedTo,RequestedBy&$filter=${`AssignedTo/EMail eq '${currentUser}' and Status eq '${FinalStatus}'`}`;
  }
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
// export const getMyApprovalsdata = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
//   let arr = []
//   if (Actingfor != null && Actingfor != undefined && Actingfor != "") {
//     await _sp.web.lists.getByTitle(listName).items
//       .select("*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail").expand("Author,AssignedTo")
//       .filter(`AssignedTo/EMail eq '${Actingfor}' and Status eq '${status}'`)
//       .orderBy("Created", false).getAll()
//       .then((res) => {
//         console.log(`--MyApproval${listName}`, res);
//         arr = res
//       })
//       .catch((error) => {
//         console.log("Error fetching data: ", error);
//       });

//     return arr;


//   } else {
//     //let arr = []
//     let currentUser;
//     await _sp.web.currentUser()
//       .then(user => {
//         console.log("user", user);
//         currentUser = user.Email; // Get the current user's Email
//       })
//       .catch(error => {
//         console.error("Error fetching current user: ", error);
//         return [];
//       });

//     if (!currentUser) return arr; // Return empty array if user fetch failed

//     await _sp.web.lists.getByTitle(listName).items
//       .select("*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail").expand("Author,AssignedTo")
//       .filter(`AssignedTo/EMail eq '${currentUser}' and Status eq '${status}'`)
//       .orderBy("Created", false).getAll()
//       .then((res) => {
//         console.log(`--MyApproval${listName}`, res);
//         arr = res
//       })
//       .catch((error) => {
//         console.log("Error fetching data: ", error);
//       });
//     return arr;

//   }
//   console.log("arrrrrrr", arr)
// }

export const updateItem = async (itemData, _sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGBusinessApps').items.getById(id).update(itemData);
    console.log('Item added successfully:', newItem);
    resultArr = newItem
    // Perform any necessary actions after successful addition
  } catch (error) {
    console.log('Error adding item:', error);
    // Handle errors appropriately
    resultArr = null
  }
  return resultArr;
};

export const uploadFileToLibrary = async (file, sp, docLib) => {

  let arrFIleData = [];
  let fileSize = 0
  try {
    const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(file.name, file,

      // const result = await sp.web.lists.getByTitle(docLib).rootFolder.files.addChunked(
      // file.name,
      // file,
      (progress, data) => {
        console.log(progress, data);
        fileSize = progress.fileSize
      },
      true
    );

    const item = await sp.web.getFileByServerRelativePath(result.data.ServerRelativeUrl).getItem("*", "ID", "AuthorId", "Modified")
    console.log(item.Id, 'itemitem');
    let arr = {
      ID: item.Id,
      Createdby: item.AuthorId,
      Modified: item.Modified,
      fileUrl: result.data.ServerRelativeUrl,
      fileSize: fileSize,
      fileType: file.type,
      fileName: file.name,
    }
    arrFIleData.push(arr)
    console.log(arrFIleData);

    return arrFIleData;
  } catch (error) {
    console.log("Error uploading file:", error);
    return null; // Or handle error differently
  }
};
export const uploadFile = async (file, sp, docLib, siteUrl) => {
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
export const getBusinessAppsByID = async (_sp, id) => {

  let arr = []
  let arrs = []
  await _sp.web.lists.getByTitle("ARGMediaGallery").items.getById(id).select("*,EntityMaster/Id,EntityMaster/Entity,MediaGalleryCategory/Id,MediaGalleryCategory/CategoryName").expand("EntityMaster,MediaGalleryCategory")()
    .then((res) => {
      console.log(res, ' let arrs=[]');
      const parsedValues = {
        Title: res.Title,
        ID: res.ID,
        entity: res.EntityMaster?.Id,
        Image: res.Image,
        MediaGalleriesId: res?.MediaGalleriesId,
        MediaGalleryJSON: res?.MediaGalleryJSON,
        Category: res?.MediaGalleryCategory?.Id
        // other fields as needed
      };
      arr.push(parsedValues)
      arrs = arr
      console.log(arrs, 'arr');
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arrs, 'arr');
  return arrs;
}
export const getUrl = async (sp, siteUrl) => {
  let srt = siteUrl;
  let dynamicPart = "/sites/";
  let UrlArr = [];
  try {
    let index = srt.lastIndexOf(dynamicPart); // Find the last occurrence of "/sites/"

    if (index !== -1) {
      let endIndex = srt.indexOf("/", index + dynamicPart.length) !== -1
        ? srt.indexOf("/", index + dynamicPart.length)
        : srt.length;

      let updatedStr = srt.slice(0, index) + srt.slice(endIndex);
      console.log(updatedStr, 'updatedStr');
      const url = await sp.web.currentUser.getContextInfo();
      console.log(url, 'res');

      let UrlArr1 =
      {
        DomainUrl: updatedStr,
        WebFullUrl: url.WebFullUrl

      }
      UrlArr.push(UrlArr1)
    } else {
      console.log("Pattern not found. No replacement was made.");
    }
  } catch (error) {
    console.log("An error occurred:", error.message);
  }

  return UrlArr
}
export const ARGBusinessAppCategory = async (sp) => {

  let arr = []
  let arrs = []
  await sp.web.lists.getByTitle("BusinessAppsCategory").items.orderBy("Created", false).getAll().then((res) => {
    console.log(res, 'Resss');

    arrs = res
    console.log(arrs, 'arr');
  })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  console.log(arrs, 'arr');
  return arrs;
}

//End
//Business apps Master
export const getBusinessApps = async (_sp) => {
  let arr = []
  let str = "Announcements"
  await _sp.web.lists.getByTitle("ARGBusinessApps")
    .items.select("*,EntityMaster/ID,EntityMaster/Entity,Category/ID,Category/CategoryName").expand("EntityMaster,Category").orderBy("Created", false).getAll()
    .then((res) => {
      console.log(res);

      //res.filter(x=>x.Category?.Category==str)
      arr = res;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const DeleteBusinessAppsAPI = async (_sp, id) => {
  let resultArr = []
  try {
    const newItem = await _sp.web.lists.getByTitle('ARGBusinessApps').items.getById(id).delete();
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
//End
// export const getApprovalListsData = async (_sp,status) => {
//   let arr = []

//   await _sp.web.lists.getByTitle("AllApprovalLists").items.orderBy("Created", false).getAll()
//     .then(async (res) => {
//       console.log("AllApprovallists",res);
//       let AllApprovalArr = [];

//       for (let i = 0; i < res.length; i++) {
//        await getMyApprovalsdata(_sp,res[i].Title,status).then((resData)=>{
//           for (let j = 0; j < resData.length; j++) {
//             AllApprovalArr.push(resData[j])
//           }

//         })
//       }
//       console.log("AllApprovalArr",AllApprovalArr);
//       arr = AllApprovalArr;
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }
export const getApprovalListsData = async (_sp, status, Actingfor) => {
  let arr = []
  

  if (!Actingfor) {
    // alert(`acting for ${Actingfor} is not null in Automation`)
    await _sp.web.lists.getByTitle("AllApprovalLists").items.filter(`IsActive eq 'Yes'`).orderBy("Created", false)()
      .then(async (res) => {
        console.log("AllApprovallists", res);
        let AllApprovalArr = [];

        for (let i = 0; i < res.length; i++) {

          if (res[i].RedirectionLinkSource == "Others" && res[i].Portal == "Others") {
            await getDataFromMultipleSites(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL)
              .then((resData) => {
                if (resData && resData.length > 0) {
                  for (let j = 0; j < resData.length; j++) {
                    // AllApprovalArr.push(resData[j])
                    AllApprovalArr.push({
                      ID: resData[j].ID,
                      RequestID: resData[j].Title,
                      // ApprovalTitle: "",
                      ApprovalTitle:resData[j]?.RequestTitle!= ""? resData[j]?.RequestTitle:"",
                      Author: resData[j].Requestor_x0020_Name,
                      ProcessName: res[i].ProcessName,
                      Created: new Date(resData[j].Created),
                      Status: resData[j].TaskStatus,
                      TaskID: resData[j].ID,
                      AppID: res[i].AppId,
                      RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                    })
                  }
                }
                // Handle the combined data here
                console.log("Final combined data:", resData);
              })
              .catch((error) => {
                console.error("Error fetching data:", error);
              });


          } else {
            await getDataFromMultipleSites(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL).then((resData) => {
              if (resData && resData.length > 0) {
                console.log("resDataresDataresDataresData", resData);
                for (let j = 0; j < resData.length; j++) {
                  AllApprovalArr.push({
                    ID: resData[j].ID,
                    RequestID: resData[j].RequestID,
                    ApprovalTitle: resData[j].ApprovalTitle,
                    Author: resData[j].RequestedBy,
                    ProcessName: resData[j].ProcessName,
                    Created: new Date(resData[j].Created),
                    Status: resData[j].Status,
                    TaskID: "",
                    AppID: "",
                    //RedirectionLink: resData[j].RedirectionLink
                     RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}${resData[j].RedirectionLink}`
                  })
                }
              }
            })
          }

        }
        console.log("AllApprovalArr in action for is not null", AllApprovalArr);
        arr = AllApprovalArr;
        console.log("AllApprovalArIF", AllApprovalArr)
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    return arr;
  } else {
    // alert(`acting for ${Actingfor} is null in Automation`)
    await _sp.web.lists.getByTitle("AllApprovalLists").items.filter(`IsActive eq 'Yes'`).orderBy("Created", false).getAll()
      .then(async (res) => {
        console.log("AllApprovallists", res);
        let AllApprovalArr = [];
        if (res.length > 0) {
          for (let i = 0; i < res.length; i++) {
            if (res[i].RedirectionLinkSource == "Others" && res[i].Portal == "Others") {
              await getDataFromMultipleSites(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL)
                .then((resData) => {
                  if (resData && resData.length > 0) {
                    for (let j = 0; j < resData.length; j++) {
                      // AllApprovalArr.push(resData[j])
                      // AllApprovalArr.push({
                      //   ID: resData[j].ID,
                      //   RequestID: resData[j].ID,
                      //   ApprovalTitle: resData[j].Title,
                      //   Author: resData[j].Requestor_x0020_Name,
                      //   ProcessName: res[i].ProcessName,
                      //   Created: resData[j].Created,
                      //   Status: resData[j].TaskStatus,
                      //   TaskID: resData[j].MasterID,
                      //   AppID: res[i].AppId,
                      //   RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                      // })
                      AllApprovalArr.push({
                        ID: resData[j].ID,
                        RequestID: resData[j].Title,
                        ApprovalTitle: "",
                        Author: resData[j].Requestor_x0020_Name,
                        ProcessName: res[i].ProcessName,
                        Created: resData[j].Created,
                        Status: resData[j].TaskStatus,
                        TaskID: resData[j].ID,
                        AppID: res[i].AppId,
                        RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                      })
                    }
                  }
                  // Handle the combined data here
                  console.log("Final combined data:", resData);
                })
                .catch((error) => {
                  console.error("Error fetching data:", error);
                });


            } else {
              await getDataFromMultipleSites(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL).then((resData) => {
                if (resData && resData.length > 0) {
                  console.log("resDataresDataresDataresData", resData);
                  for (let j = 0; j < resData.length; j++) {
                    AllApprovalArr.push({
                      ID: resData[j].ID,
                      RequestID: resData[j].RequestID,
                      ApprovalTitle: resData[j].ApprovalTitle,
                      Author: resData[j].Author,
                      ProcessName: resData[j].ProcessName,
                      Created: resData[j].Created,
                      Status: resData[j].Status,
                      TaskID: "",
                      AppID: "",
                      RedirectionLink: resData[j].RedirectionLink
                    })
                  }
                }
              })
            }
          }
        }
        console.log("AllApprovalArr in action for is not null", AllApprovalArr);
        arr = AllApprovalArr;
        console.log("AllApprovalArrelse", AllApprovalArr)
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    return arr;
  }

}

export const getMyRequestsdata = async (_sp, listName) => {
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
    .filter(`Author/EMail eq '${currentUser}'`)
    .orderBy("Created", false).getAll()
    .then((res) => {
      console.log(`--MyRequest${listName}`, res);
      arr = res
      // arr = res.filter(item => 
      //     // Include public groups or private groups where the current user is in the InviteMembers array
      //     item.GroupType === "Public" || 
      //     (item.GroupType === "Private" && item.InviteMemebers && item.InviteMemebers.some(member => member.Id === currentUser))
      //   );
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}

export const getRequestListsData = async (_sp) => {
  let arr = []

  await _sp.web.lists.getByTitle("AllRequestLists").items.filter(`IsActive eq 'Yes'`).orderBy("Created", false).getAll()
    .then((res) => {
      console.log("AllRequestLists", res);
      let AllRequestArr = [];

      for (let i = 0; i < res.length; i++) {
        getMyRequestsdata(_sp, res[i].Title).then((resData) => {
          for (let j = 0; j < resData.length; j++) {
            AllRequestArr.push(resData[j])
          }

        })
      }
      console.log("AllRequestArr", AllRequestArr);
      arr = AllRequestArr;
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}
export const getApprovalListsDataAudit = async (_sp, status, Actingfor) => {
  let arr = []
  

  if (!Actingfor) {
    // alert(`acting for ${Actingfor} is not null in Automation`)
    await _sp.web.lists.getByTitle("AllApprovalLists").items.filter(`IsActive eq 'Yes'`).orderBy("Created", false).getAll()
      .then(async (res) => {
        console.log("AllApprovallists", res);
        let AllApprovalArr = [];

        for (let i = 0; i < res.length; i++) {

          if (res[i].RedirectionLinkSource == "Others" && res[i].Portal == "Others") {
            await getDataFromMultipleSitesAudit(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL)
              .then((resData) => {
                if (resData && resData.length > 0) {
                  for (let j = 0; j < resData.length; j++) {
                    // AllApprovalArr.push(resData[j])
                    AllApprovalArr.push({
                      ID: resData[j].ID,
                      RequestID: resData[j].Title,
                      // ApprovalTitle: "",
                      ApprovalTitle:resData[j]?.RequestTitle!= ""? resData[j]?.RequestTitle:"",
                      Author: resData[j].Requestor_x0020_Name,
                      ProcessName: res[i].ProcessName,
                      Created: new Date(resData[j].Created),
                      Status: resData[j].TaskStatus,
                      TaskID: resData[j].ID,
                      AppID: res[i].AppId,
                      RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                    })
                  }
                }
                // Handle the combined data here
                console.log("Final combined data:", resData);
              })
              .catch((error) => {
                console.error("Error fetching data:", error);
              });


          } else {
            await getDataFromMultipleSitesAudit(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL).then((resData) => {
              if (resData && resData.length > 0) {
                console.log("resDataresDataresDataresData", resData);
                for (let j = 0; j < resData.length; j++) {
                  AllApprovalArr.push({
                    ID: resData[j].ID,
                    RequestID: resData[j].RequestID,
                    ApprovalTitle: resData[j].ApprovalTitle,
                    Author: resData[j].Author,
                    ProcessName: resData[j].ProcessName,
                    Created: new Date(resData[j].Created),
                    Status: resData[j].Status,
                    TaskID: "",
                    AppID: "",
                    RedirectionLink: resData[j].RedirectionLink
                  })
                }
              }
            })
          }

        }
        console.log("AllApprovalArr in action for is not null", AllApprovalArr);
        arr = AllApprovalArr;
        console.log("AllApprovalArIF", AllApprovalArr)
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    return arr;
  } else {
    // alert(`acting for ${Actingfor} is null in Automation`)
    await _sp.web.lists.getByTitle("AllApprovalLists").items.filter(`IsActive eq 'Yes'`).orderBy("Created", false).getAll()
      .then(async (res) => {
        console.log("AllApprovallists", res);
        let AllApprovalArr = [];
        if (res.length > 0) {
          for (let i = 0; i < res.length; i++) {
            if (res[i].RedirectionLinkSource == "Others" && res[i].Portal == "Others") {
              await getDataFromMultipleSitesAudit(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL)
                .then((resData) => {
                  if (resData && resData.length > 0) {
                    for (let j = 0; j < resData.length; j++) {
                      // AllApprovalArr.push(resData[j])
                      // AllApprovalArr.push({
                      //   ID: resData[j].ID,
                      //   RequestID: resData[j].ID,
                      //   ApprovalTitle: resData[j].Title,
                      //   Author: resData[j].Requestor_x0020_Name,
                      //   ProcessName: res[i].ProcessName,
                      //   Created: resData[j].Created,
                      //   Status: resData[j].TaskStatus,
                      //   TaskID: resData[j].MasterID,
                      //   AppID: res[i].AppId,
                      //   RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                      // })
                      AllApprovalArr.push({
                        ID: resData[j].ID,
                        RequestID: resData[j].Title,
                        ApprovalTitle: "",
                        Author: resData[j].Requestor_x0020_Name,
                        ProcessName: res[i].ProcessName,
                        Created: resData[j].Created,
                        Status: resData[j].TaskStatus,
                        TaskID: resData[j].ID,
                        AppID: res[i].AppId,
                        RedirectionLink: `https://apps.powerapps.com/apps/${res[i].AppId}?hidenavbar=true&RequestNo=${resData[j].MasterID}&TaskNo=${resData[j].ID}`
                      })
                    }
                  }
                  // Handle the combined data here
                  console.log("Final combined data:", resData);
                })
                .catch((error) => {
                  console.error("Error fetching data:", error);
                });


            } else {
              await getDataFromMultipleSitesAudit(_sp, res[i].Title, status, Actingfor, res[i].RedirectionLinkSource, res[i].SiteBaseURL).then((resData) => {
                if (resData && resData.length > 0) {
                  console.log("resDataresDataresDataresData", resData);
                  for (let j = 0; j < resData.length; j++) {
                    AllApprovalArr.push({
                      ID: resData[j].ID,
                      RequestID: resData[j].RequestID,
                      ApprovalTitle: resData[j].ApprovalTitle,
                      Author: resData[j].Author,
                      ProcessName: resData[j].ProcessName,
                      Created: resData[j].Created,
                      Status: resData[j].Status,
                      TaskID: "",
                      AppID: "",
                      RedirectionLink: resData[j].RedirectionLink
                    })
                  }
                }
              })
            }
          }
        }
        console.log("AllApprovalArr in action for is not null", AllApprovalArr);
        arr = AllApprovalArr;
        console.log("AllApprovalArrelse", AllApprovalArr)
      })
      .catch((error) => {
        console.log("Error fetching data: ", error);
      });
    return arr;
  }

}
export const getDataFromMultipleSitesAudit = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
  const allData = [];
  const siteUrls = [
    SiteBaseURL
  ];

  // Loop through each site collection URL and fetch data
  if (portal == "Others"){
    for (const siteUrl of siteUrls) {
      const data = await getListDataFromSiteCollectionAudit(_sp, listName, status, Actingfor, portal, siteUrl);
      allData.push(...data);
      console.log("dadadadadad", data);
    }
  } else{
    for (const siteUrl of siteUrls) {
      const data = await getMyApprovalsdataAudit(_sp, listName, status, Actingfor, portal, siteUrl);
      allData.push(...data);
      console.log("dadadadadad", data);
    }
  }


  console.log('Combined data from all site collections:', allData);
  return allData;
}
export const getListDataFromSiteCollectionAudit = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
  // Setup PnPJs for a specific site collection URL
  let arr = [];
  console.log("sdsssss", sp, _sp)
  
  let apiUrl;
  let FinalStatus = "";
  if (status == "Pending") {
    FinalStatus = "Not Started"
  } else if (status == "Approved") {
    FinalStatus = "Completed"
  }
  if (Actingfor != null && Actingfor != undefined && Actingfor != "") {
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail&$expand=Requestor_x0020_Name,AssignedTo&$filter=${`AssignedTo/EMail eq '${Actingfor}' and TaskStatus eq '${FinalStatus}'`}`;
  } else {
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
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Requestor_x0020_Name/ID,Requestor_x0020_Name/Title,Requestor_x0020_Name/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail&$expand=Requestor_x0020_Name,AssignedTo&$filter=${`AssignedTo/EMail eq '${currentUser}' and TaskStatus eq '${FinalStatus}'`}`;
  }
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
export const getMyApprovalsdataAudit = async (_sp, listName, status, Actingfor, portal, SiteBaseURL) => {
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
  if (Actingfor != null && Actingfor != undefined && Actingfor != "") {
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail&$expand=Author,AssignedTo&$filter=${`AssignedTo/EMail eq '${Actingfor}' and Status eq '${FinalStatus}'`}`;
  } else {
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
    apiUrl = `${SiteBaseURL}/_api/web/lists/getbytitle('${listName}')/items?$select=*,Author/ID,Author/Title,Author/EMail,AssignedTo/ID,AssignedTo/Title,AssignedTo/EMail&$expand=Author,AssignedTo&$filter=${`AssignedTo/EMail eq '${currentUser}' and Status eq '${FinalStatus}'`}`;
  }
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