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

export default class DynamicContentWebPart extends BaseClientSideWebPart<IDynamicContentWebPartProps> {

    private sp: SPFI;

    public onInit(): Promise<void> {
        // Initialize PnP.js
        this.sp = getSP(this.context);
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
                demoMode: false, // Enable demo mode
            }
        );

        ReactDom.render(element, this.domElement);
    }


    public onDispose(): void {
        // Unmount React component to avoid memory leaks
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
