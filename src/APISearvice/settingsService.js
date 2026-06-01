
// export const getSettingAPI = async (_sp) => {
//     let arr =[]
//    await _sp.web.lists.getByTitle("Settings").items.select("Title,ID,ImageIcon,LinkUrl")()
//     .then((res) => {
//         console.log(res);
//         arr= res;
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return arr;
// }



// export const getSettingAPI = async (_sp) => {
//     let arr =[]
//    await _sp.web.lists.getByTitle("Settings").items.select("Title,ID,ImageIcon,LinkUrl")()
//     .then((res) => {
//         console.log(res);
//         arr= res;
//     })
//     .catch((error) => {
//         console.error("Error fetching data: ", error);
//     });
//     return arr;
// }

export const getSettingAPImanagemaster = async (_sp) => {
    let arr =[]

   await _sp.web.lists.getByTitle("Settings").items
    .select(
      "Title,ID,ImageIcon,LinkUrl,EnableAudienceTargeting,Audience/Title,Category,IsActive,Order"
    )
    .expand("Audience")
    .filter("IsActive eq 'Yes' and Category eq 'Manage Master'")
    .orderBy("Order", true)()

    .then((res) => {
        console.log("Responce of data for get master:",res);
        arr= res;
    })
    .catch((error) => {
        console.log("Error fetching data: ", error);
    });

    return arr;
}

export const getSettingAPIPpowerappsmaster = async (_sp) => {
    let arr =[]

   await _sp.web.lists.getByTitle("Settings").items
    .select(
      "Title,ID,ImageIcon,LinkUrl,EnableAudienceTargeting,Audience/Title,Category,IsActive,Order"
    )
    .expand("Audience")
    .filter("IsActive eq 'Yes' and Category eq 'Power Apps Master'")
    .orderBy("Order", true)()

    .then((res) => {
        console.log("Responce of data for get master:",res);
        arr= res;
    })
    .catch((error) => {
        console.log("Error fetching data: ", error);
    });

    return arr;
}
export const getSettingAPI = async (_sp) => {
    let arr =[]
   await _sp.web.lists.getByTitle("Settings").items.select("Title,ID,ImageIcon,LinkUrl")()
    .then((res) => {
        console.log(res);
        arr= res;
    })
    .catch((error) => {
        console.error("Error fetching data: ", error);
    });
    return arr;
}