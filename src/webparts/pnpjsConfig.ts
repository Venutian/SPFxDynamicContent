import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/behaviors/spfx"; // Correct path for SPFx behavior in v4

let sp: SPFI;

export const getSP = (context: any): SPFI => {
    if (!sp) {
        sp = spfi().using(SPFx(context)); // Bind SPFx context using the correct path
    }
    return sp;
};
