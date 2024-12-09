import * as React from 'react';
import {IDynamicContentWebPartProps} from './IDynamicContentWebPartProps';
import styles from './DynamicContentWebPart.module.scss';
import {ILinkItem} from './IDynamicContentWebPartProps';
import {Web} from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

interface IDynamicContentWebPartState {
    pages: ILinkItem[];
}

export default class DynamicContentComponent extends React.Component<IDynamicContentWebPartProps, IDynamicContentWebPartState> {
    private updateInterval: NodeJS.Timeout | null = null;

    constructor(props: IDynamicContentWebPartProps) {
        super(props);
        this.state = {
            pages: []
        };
    }

    public async componentDidMount(): Promise<void> {
        await this.ensureListExists();
        await this.cleanUpOldData();
        await this.loadPages();

        this.updateInterval = setInterval(async () => {
            await this.cleanUpOldData();
            await this.loadPages();
        }, 24 * 60 * 60 * 1000); // Every 24 hours
    }

    public componentWillUnmount(): void {
        if (this.updateInterval) {
            clearInterval(this.updateInterval);
        }
    }

    private async ensureListExists(): Promise<void> {
        const isLocal = this.props.demoMode || !this.props.context.pageContext.web.absoluteUrl.includes("https");

        if (isLocal) {
            console.log("Running in demo mode. Skipping list existence check.");
            return;
        }

        const webUrl = this.props.context.pageContext.web.absoluteUrl;
        const web = Web(webUrl);

        try {
            await web.lists.getByTitle(this.props.listName || "DailyClickCounts").select("Title")();
            console.log("List exists.");
        } catch (error) {
            console.log("List does not exist. Creating...");
            try {
                await web.lists.add(this.props.listName || "DailyClickCounts", "Stores click counts for pages", 100);
                console.log("List created.");
            } catch (createError) {
                console.error("Error creating list:", createError);
            }
        }
    }

    private async cleanUpOldData(): Promise<void> {
        const webUrl = this.props.context.pageContext.web.absoluteUrl;
        const web = Web(webUrl);

        const sevenDaysAgo = new Date();
        sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

        try {
            const items = await web.lists
                .getByTitle(this.props.listName || "DailyClickCounts")
                .items.select("Id", "ClickCounts")();

            for (const item of items) {
                const clickCounts = JSON.parse(item.ClickCounts || "{}");
                for (const role in clickCounts) {
                    if (Object.prototype.hasOwnProperty.call(clickCounts, role)) {
                        clickCounts[role] = clickCounts[role].filter((entry: { timestamp: string }) =>
                            new Date(entry.timestamp) >= sevenDaysAgo
                        );
                    }
                }
                await web.lists.getByTitle(this.props.listName || "DailyClickCounts").items.getById(item.Id).update({
                    ClickCounts: JSON.stringify(clickCounts)
                });
            }
            console.log("Old data cleaned up.");
        } catch (error) {
            console.error("Error cleaning up old data:", error);
        }
    }

    private async loadPages(): Promise<void> {
        const isLocal = this.props.demoMode || !this.props.context.pageContext.web.absoluteUrl.includes("https");
        const userRole = this.props.userRole || "Admin"; // Default to a valid role for demo

        if (isLocal) {
            console.log("Using demo mode with mocked data...");
            const mockedItems = [
                {
                    Id: 1,
                    Title: "Admin Dashboard",
                    URL: "/sites/admin",
                    ClickCounts: JSON.stringify({
                        Admin: [{ timestamp: "2024-11-15T10:00:00Z" }],
                    }),
                    Roles: "Admin,User",
                },
                {
                    Id: 2,
                    Title: "User Profile",
                    URL: "/sites/userprofile",
                    ClickCounts: JSON.stringify({
                        User: [{ timestamp: "2024-11-15T11:00:00Z" }],
                    }),
                    Roles: "User",
                },
                {
                    Id: 3,
                    Title: "Reports",
                    URL: "/sites/reports",
                    ClickCounts: JSON.stringify({
                        Admin: [{ timestamp: "2024-11-15T12:00:00Z" }],
                        User: [{ timestamp: "2024-11-15T12:30:00Z" }],
                    }),
                    Roles: "Admin,User",
                },
            ];

            // Debug logging for mocked items and userRole
            console.log("Mocked Items:", mockedItems);
            console.log("User Role:", userRole);

            const pages = mockedItems.map((item) => {
                const clickCounts = JSON.parse(item.ClickCounts || "{}");
                const totalClicks = (clickCounts[userRole] || []).length;

                return {
                    id: item.Id,
                    title: item.Title,
                    url: item.URL,
                    clicks: totalClicks,
                    roles: item.Roles.split(","),
                };
            }) as ILinkItem[];

            const roleFilteredPages = pages.filter((page) => page.roles.includes(userRole));
            roleFilteredPages.sort((a, b) => b.clicks - a.clicks);

            this.setState({ pages: roleFilteredPages }, () => {
                console.log("Filtered Pages:", this.state.pages); // Debug filtered pages
            });
        } else {
            console.log("Fetching live SharePoint data...");
            const webUrl = this.props.context.pageContext.web.absoluteUrl;
            const web = Web(webUrl);

            try {
                const items = await web.lists
                    .getByTitle(this.props.listName || "DailyClickCounts")
                    .items.select("Id", "Title", "URL", "ClickCounts", "Roles")();

                const pages = items.map((item) => {
                    const clickCounts = JSON.parse(item.ClickCounts || "{}");
                    const totalClicks = (clickCounts[userRole] || []).length;

                    return {
                        id: item.Id,
                        title: item.Title,
                        url: item.URL,
                        clicks: totalClicks,
                        roles: item.Roles.split(","),
                    };
                }) as ILinkItem[];

                const roleFilteredPages = pages.filter((page) => page.roles.includes(userRole));
                roleFilteredPages.sort((a, b) => b.clicks - a.clicks);

                this.setState({ pages: roleFilteredPages });
            } catch (error) {
                console.error("Error fetching live data:", error);
            }
        }
    }

    private async handlePageClick(pageId: number, userRole: string): Promise<void> {
        const webUrl = this.props.context.pageContext.web.absoluteUrl;
        const web = Web(webUrl);

        try {
            const item = await web.lists
                .getByTitle(this.props.listName || "DailyClickCounts")
                .items.getById(pageId)
                .select("ClickCounts")();

            const clickCounts = JSON.parse(item.ClickCounts || "{}");

            const currentTimestamp = new Date().toISOString();
            if (!clickCounts[userRole]) {
                clickCounts[userRole] = [];
            }
            clickCounts[userRole].push({timestamp: currentTimestamp});

            await web.lists.getByTitle(this.props.listName || "DailyClickCounts").items.getById(pageId).update({
                ClickCounts: JSON.stringify(clickCounts)
            });

            console.log("Click count updated.");
            await this.loadPages(); // Refresh the UI
        } catch (error) {
            console.error("Error updating click count:", error);
        }
    }

    public render(): React.ReactElement<IDynamicContentWebPartProps> {
        const { pages } = this.state;

        return (
            <section className={styles.dynamicContentWebPart}>
                {pages.length > 0 ? (
                    pages.map((page) => (
                        <a
                            key={page.id}
                            href={page.url}
                            target="_blank"
                            rel="noopener noreferrer"
                            onClick={() => this.handlePageClick(page.id, this.props.userRole)}
                            className={styles.pageButton}
                        >
                            <div className={styles.icon}>
                                {/* Add icon dynamically based on the page or use default */}
                                <i className="ms-Icon ms-Icon--Globe" aria-hidden="true"></i>
                            </div>
                            <div className={styles.title}>
                                {page.title} <br /> ({page.clicks} clicks)
                            </div>
                        </a>
                    ))
                ) : (
                    <p>No pages available to display.</p>
                )}
            </section>
        );
    }

}
