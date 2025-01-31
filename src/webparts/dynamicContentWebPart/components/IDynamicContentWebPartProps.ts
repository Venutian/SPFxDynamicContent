import { SPFI } from "@pnp/sp";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDynamicContentWebPartProps {
  description: string;
  userRole: string;
  sp: SPFI;
  context: WebPartContext;
  listName: string;
}

export interface ILinkItem {
  id: number;
  title: string;
  url: string;
  clicks: number;
  roles: string[];
  icon: string;
}
