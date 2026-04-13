
import { getSP } from '../webparts/essaApproval/loc/pnpjsConfig';
import { SPFI } from "@pnp/sp/presets/all";
import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

 const sp = getSP();

export const  getbreadcrumbhomeurlfunction = async (_context) => {
     //alert('coming heree')
    const url = await _context.web.lists.getByTitle("HomeMasterURL").items.select("*").getAll()
    // alert(url + JSON.stringify(url)); 

     let arr = ''

  await _context.web.lists.getByTitle("HomeMasterURL").items
  .select("*,TabName , URL")
  .filter(`TabName eq 'Home'`)()
    .then((res) => {
      console.log(res);
       //alert(res + 'res' + JSON.stringify(res) 
      //  )
        //alert(res[0].URL + "url [0]")
      //res.filter(x=>x.Category?.Category==str)
      if(res.length > 0){
        arr = res[0].URL;
      }
     
    })
    .catch((error) => {
      console.log("Error fetching data: ", error);
    });
  return arr;
}