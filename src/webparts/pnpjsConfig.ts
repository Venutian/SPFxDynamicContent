import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx";
import { WebPartContext } from '@microsoft/sp-webpart-base';

let sp: SPFI;

export const getSP = (context: WebPartContext): SPFI => {
    if (!sp) {
        if (context) {
            try {
                sp = spfi().using(SPFx(context)); // Bind SPFx context
                console.log("PnP.js initialized successfully.");
            } catch (error) {
                console.error("Error initializing PnP.js:", error);
            }
        } else {
            console.error("SPFx context is missing or invalid.");
        }
    }
    return sp;
};