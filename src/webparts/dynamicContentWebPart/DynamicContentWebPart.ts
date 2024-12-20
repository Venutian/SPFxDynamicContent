import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import DynamicContentComponent from './components/DynamicContentWebPart';

import { IDynamicContentWebPartProps } from './components/IDynamicContentWebPartProps';
import { getSP } from "../pnpjsConfig"; // PnP.js config import
import { SPFI } from "@pnp/sp";
import { Web } from "@pnp/sp/webs";

export default class DynamicContentWebPart extends BaseClientSideWebPart<IDynamicContentWebPartProps> {

    private sp: SPFI;

    public async onInit(): Promise<void> {
        console.log("Initializing SPFx Web Part...");
        this.sp = getSP(this.context);

        try {
            console.log("Checking connection to SharePoint...");
            await Web(this.context.pageContext.web.absoluteUrl).lists.select("Title")();
            console.log("Connection to SharePoint established.");
        } catch (error) {
            console.error("Error connecting to SharePoint:", error);
        }

        return super.onInit();
    }

    public render(): void {
        const element: React.ReactElement<IDynamicContentWebPartProps> = React.createElement(
            DynamicContentComponent,
            {
                description: this.properties.description,
                userRole: this.properties.userRole || "Admin", // Set a default role for demo
                sp: this.sp,
                context: this.context,
                listName: this.properties.listName,
                demoMode: false, // Disable demo mode for live testing
            }
        );

        ReactDom.render(element, this.domElement);
    }

    public onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: { description: "Web Part Configuration" },
                    groups: [
                        {
                            groupName: "Settings",
                            groupFields: [
                                PropertyPaneTextField("listName", {
                                    label: "List Name",
                                    description: "Enter the name of the SharePoint list",
                                    value: "DailyClickCounts" // Default value
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    }
}