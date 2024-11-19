import { SPFI } from "@pnp/sp";

// Interface for the web part properties
export interface IDynamicContentWebPartProps {
  description: string; // Description of the web part
  userRole: string; // Role of the current user
  sp: SPFI; // Instance of PnP.js for SharePoint operations
  context: any; // SPFx context
  listName: string; // Name of the SharePoint list to fetch data from
}

// Interface for a single link item (page)
export interface ILinkItem {
  id: number;
  title: string;
  url: string;
  clicks: number;
  roles: string[];
}

