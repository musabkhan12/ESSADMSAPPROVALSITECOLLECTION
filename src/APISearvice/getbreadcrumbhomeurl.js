
// import { getSP } from '../webparts/essaApproval/loc/pnpjsConfig';
// import { SPFI } from "@pnp/sp/presets/all";
// import { Web } from "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";

//  const sp = getSP();

// export const  getbreadcrumbhomeurlfunction = async (_context) => {
//      //alert('coming heree')
//     const url = await _context.web.lists.getByTitle("HomeMasterURL").items.select("*").getAll()
//     // alert(url + JSON.stringify(url)); 

//      let arr = ''

//   await _context.web.lists.getByTitle("HomeMasterURL").items
//   .select("*,TabName , URL")
//   .filter(`TabName eq 'Home'`)()
//     .then((res) => {
//       console.log(res);
//        //alert(res + 'res' + JSON.stringify(res) 
//       //  )
//         //alert(res[0].URL + "url [0]")
//       //res.filter(x=>x.Category?.Category==str)
//       if(res.length > 0){
//         arr = res[0].URL;
//       }
     
//     })
//     .catch((error) => {
//       console.log("Error fetching data: ", error);
//     });
//   return arr;
// }




import { getSP } from '../webparts/essaApproval/loc/pnpjsConfig';
import { SPFI } from "@pnp/sp/presets/all";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export const getbreadcrumbhomeurlfunction = async (_context) => {
    try {
        // Get all items from HomeMasterURL list - FIX: Remove .getAll()
        const items = await _context.web.lists.getByTitle("HomeMasterURL").items();
        
        // Filter for items where TabName equals 'Home'
        const filteredItems = items.filter(item => item.TabName === 'Home');
        
        // Return the URL of the first matching item, or empty string if none found
        return filteredItems.length > 0 ? filteredItems[0].URL : '';
    } catch (error) {
        console.log("Error fetching data: ", error);
        return '';
    }
}