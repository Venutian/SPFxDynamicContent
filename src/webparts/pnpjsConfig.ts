import {spfi, SPFI} from "@pnp/sp";
import {SPFx} from "@pnp/sp/behaviors/spfx";
import {WebPartContext} from "@microsoft/sp-webpart-base";

// Keep a global reference so we only init once
let _sp: SPFI | undefined;

export const getSP = (context: WebPartContext): SPFI => {
    if (!_sp) {
        // Initialize the PnPjs SP object
        _sp = spfi().using(SPFx(context));
        console.log("PnP.js initialized successfully for:", context.pageContext.web.absoluteUrl);
    }
    return _sp;
};