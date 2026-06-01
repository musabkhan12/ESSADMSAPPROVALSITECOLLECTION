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
// let _sp: SPFI | null = null;
// Correct imports for the properties you requested:
import "@pnp/sp/sites"; // for _sp.site
import "@pnp/sp/hubsites"; // for _sp.hubSites
import "@pnp/sp/sputilities"; // for _sp.utility (NOTE the 'sp' prefix)
import "@pnp/sp/navigation"; // for _sp.web.navigation
import "@pnp/sp/site-designs"; // for _sp.siteDesigns and _sp.siteScripts
import "@pnp/sp/appcatalog"; // for _sp.tenantAppcatalog
export var setupPnP = function (context) {
    _sp = spfi().using(SPFx(context));
};
var _sp;
export var getSP = function (context) {
    if (context !== undefined && (_sp === undefined || _sp === null)) {
        //You must add the @pnp/logging package to include the PnPLogging behavior it is no longer a peer dependency
        // The LogLevel set's at what level a message will be written to the console
        _sp = spfi().using(SPFx(context)).using(PnPLogging(LogLevel.Warning));
    }
    return _sp;
};
//# sourceMappingURL=pnpjsConfig.js.map