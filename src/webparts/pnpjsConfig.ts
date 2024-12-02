import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx"; // Correct path for SPFx behavior in v4

let sp: SPFI;

export const getSP = (context: any): SPFI => {
    if (!sp) {
        try {
            sp = spfi().using(SPFx(context)); // Bind SPFx context
            console.log("PnP.js initialized successfully.");
        } catch (error) {
            console.error("Error initializing PnP.js:", error);
        }
    }
    return sp;
};
