// import pnp and pnp logging system
import { spfi, SPFx } from "@pnp/sp";
import { LogLevel, PnPLogging } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
import "@pnp/sp/profiles";
// import "@pnp/sp/items/get-all";
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
var _sp;
export var getSP = function (context) {
    // Use !!context or context !== undefined to ensure it's actually there
    if (context && (_sp === undefined || _sp === null)) {
        _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _sp;
};
var _spurl;
export var getSPContext = function (context) {
    // Same fix here
    if (context && (_spurl === undefined || _spurl === null)) {
        _spurl = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _spurl;
};
//# sourceMappingURL=pnpjsConfig.js.map