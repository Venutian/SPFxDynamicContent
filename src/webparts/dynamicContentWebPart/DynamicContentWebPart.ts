import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
    IPropertyPaneConfiguration,
    PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import DynamicContentComponent from './components/DynamicContentWebPart';
import { IDynamicContentWebPartProps } from './components/IDynamicContentWebPartProps';
import { getSP } from '../pnpjsConfig';
import { SPFI } from '@pnp/sp';

export default class DynamicContentWebPart extends BaseClientSideWebPart<IDynamicContentWebPartProps> {
    private sp: SPFI;

    public onInit(): Promise<void> {
        this.sp = getSP(this.context);
        console.log('onInit: sp =>', this.sp);
        return super.onInit();
    }

    public render(): void {
        const element: React.ReactElement<IDynamicContentWebPartProps> = React.createElement(
            DynamicContentComponent,
            {
                description: "Dynamisk Sidor - Prioritera och spåra länkklick dynamiskt.",
                userRole: this.properties.userRole,
                sp: this.sp,
                context: this.context,
                listName: this.properties.listName,
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
                    header: { description: 'Dynamisk Sidor Konfiguration' },
                    groups: [
                        {
                            groupName: 'Inställningar',
                            groupFields: [
                                PropertyPaneTextField('listName', {
                                    label: 'Listnamn',
                                    description: 'Ange namnet på SharePoint-listan',
                                    value: 'KlickPrioritet',
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    }
}
