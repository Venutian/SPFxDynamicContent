import { SPFI } from "@pnp/sp";
import { WebPartContext } from '@microsoft/sp-webpart-base';

/**
 * Properties passed to the DynamicContentComponent.
 */
export interface IDynamicContentWebPartProps {
  description: string; // Description of the web part
  userRole: string; // Role of the current user
  sp: SPFI; // Instance of PnP.js for SharePoint operations
  context: WebPartContext; // Strongly typed SPFx context
  listName: string; // Name of the SharePoint list to fetch data from
  demoMode?: boolean; // Optional flag to enable demo mode
}

/**
 * Interface representing a single page/item in the SharePoint list.
 */
export interface ILinkItem {
  id: number; // Unique ID of the list item
  title: string; // Title of the page or link
  url: string; // URL of the page or resource
  clicks: number; // Total clicks for the role
  roles: string[]; // Roles associated with the page
}
