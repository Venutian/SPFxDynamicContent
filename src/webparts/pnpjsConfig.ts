import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";
import { WebPartContext } from "@microsoft/sp-webpart-base";

let _sp: SPFI | undefined = undefined;

export const getSP = (context?: WebPartContext): SPFI => {
    if (!_sp && context) {
        try {
            _sp = spfi().using(SPFx(context));
            console.log("PnP.js initialized successfully.");
        } catch (error) {
            console.error("Error initializing PnP.js:", error);
        }
    }
    return _sp!;
};
