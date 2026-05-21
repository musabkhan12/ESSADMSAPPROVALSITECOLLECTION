import { WebPartContext } from "@microsoft/sp-webpart-base";
// import pnp and pnp logging system
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
import "@pnp/sp/profiles";
// import "@pnp/sp/items/get-all";
import "@pnp/sp/items"; 
import "@pnp/sp/folders";
import "@pnp/sp/files/folder";
import "@pnp/sp/fields";
import "@pnp/sp/files";
import "@pnp/sp/security";
import "@pnp/sp/presets/all";
// var _sp: SPFI;
// export const getSP = (context?: WebPartContext): SPFI => {
//   if (context != null && (_sp === undefined ||_sp === null)) {
//     //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
//     // The LogLevel set's at what level a message will be written to the console
//     _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
//   }
//   return _sp;
// };



// ✅ Helper function to get all items (handles pagination)
export const getAllItems = async <T = any>(
  sp: SPFI,
  listTitle: string,
  select: string[] = [],
  expand: string[] = [],
  filter: string = "",
  orderBy?: string,
  ascending: boolean = true
): Promise<T[]> => {
  let allItems: T[] = [];
  let skip = 0;
  const top = 4000; // SharePoint max limit per request

  let query = sp.web.lists.getByTitle(listTitle).items;
  if (select.length) query = query.select(select.join(","));
  if (expand.length) query = query.expand(expand.join(","));
  if (filter) query = query.filter(filter);
  if (orderBy) query = query.orderBy(orderBy, ascending);

  while (true) {
    const batch = await query.top(top).skip(skip)();
    if (batch.length === 0) break;
    allItems = allItems.concat(batch);
    if (batch.length < top) break;
    skip += top;
  }
  return allItems;
};

var _sp: SPFI;
export const getSP = (context?: WebPartContext): SPFI => {
  
    
  if (context !== null && (_sp === undefined ||_sp === null)) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
    _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  
    
  }
  return _sp;
}
var _spurl: SPFI;
export const getSPContext = (context?: WebPartContext): SPFI=> {
  if (context !== null && (_spurl === undefined ||_spurl === null)) {
    //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
    // The LogLevel set's at what level a message will be written to the console
   _spurl = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
  
    
  }
  return _spurl;
  
}