import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx"; // Correct SPFx behavior

let sp: SPFI;

export const getSP = (context: any): SPFI => {
    if (!sp) {
        console.log("Initializing PnP.js instance...");
        try {
            sp = spfi().using(SPFx(context));
            console.log("PnP.js initialized successfully.");
        } catch (error) {
            console.error("Error initializing PnP.js:", error);
        }
    }

    return sp;
};
