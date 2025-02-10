import * as React from 'react';
import {IDynamicContentWebPartProps} from './IDynamicContentWebPartProps';
import styles from './DynamicContentWebPart.module.scss';
import {ILinkItem} from './IDynamicContentWebPartProps';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs/index";
import "@pnp/sp/fields/list";

interface IDynamicContentWebPartState {
    pages: ILinkItem[];
}

export default class DynamicContentComponent extends React.Component<
    IDynamicContentWebPartProps,
    IDynamicContentWebPartState
> {
    private updateInterval: NodeJS.Timeout | null = null;

    constructor(props: IDynamicContentWebPartProps) {
        super(props);
        this.state = {pages: []};
    }

    public async componentDidMount(): Promise<void> {
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

    private async cleanUpOldData(): Promise<void> {
        const sp = this.props.sp;
        const listName = this.props.listName || "KlickPrioritet";
        const sevenDaysAgo = new Date();
        sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

        try {
            const items = await sp.web.lists
                .getByTitle(listName)
                .items.select("Id", "ClickCounts")();

            for (const item of items) {
                const clickCounts = JSON.parse(item.ClickCounts || "{}");
                for (const role in clickCounts) {
                    if (Object.prototype.hasOwnProperty.call(clickCounts, role)) {
                        clickCounts[role] = clickCounts[role].filter(
                            (entry: { timestamp: string }) =>
                                new Date(entry.timestamp) >= sevenDaysAgo
                        );
                    }
                }
                await sp.web.lists.getByTitle(listName).items.getById(item.Id).update({
                    ClickCounts: JSON.stringify(clickCounts),
                });
            }
            console.log("Old data cleaned up.");
        } catch (error) {
            console.error("Error cleaning up old data:", error);
        }
    }

    private async loadPages(): Promise<void> {
        const sp = this.props.sp;
        const listName = this.props.listName || "KlickPrioritet";
        const userRole = this.props.userRole || "Admin";

        if (!userRole) {
            console.error("User role is missing.");
            return;
        }

        console.log("Fetching live SharePoint data...");

        try {
            const items = await sp.web.lists
                .getByTitle(listName)
                .items.select("Id", "Title", "URL", "ClickCounts", "Roles", "Icon")();

            const pages = items.map((item) => {
                const clickCounts = JSON.parse(item.ClickCounts || "{}");
                const totalClicks = (clickCounts[userRole] || []).length;
                return {
                    id: item.Id,
                    title: item.Title,
                    url: item.URL,
                    clicks: totalClicks,
                    roles: item.Roles.split(","),
                    icon: item.Icon, // Include the icon
                };
            }) as ILinkItem[];

            const roleFilteredPages = pages.filter((page) =>
                page.roles.includes(userRole)
            );
            roleFilteredPages.sort((a, b) => b.clicks - a.clicks);

            // Ensure only 10 buttons are shown, with "Other" always as the last one
            const displayedPages = roleFilteredPages.slice(0, 11);
            const otherButton = roleFilteredPages.find((page) => page.title === "Övriga System");
            if (otherButton) {
                displayedPages.push(otherButton);
            }

            this.setState({pages: displayedPages});
        } catch (error) {
            console.error("Error fetching live data:", error);
        }
    }

    private async handlePageClick(pageId: number, userRole: string): Promise<void> {
        const sp = this.props.sp;
        const listName = this.props.listName || "KlickPrioritet";

        try {
            const item = await sp.web.lists
                .getByTitle(listName)
                .items.getById(pageId)
                .select("ClickCounts")();

            const clickCounts = JSON.parse(item.ClickCounts || "{}");
            if (!clickCounts[userRole]) {
                clickCounts[userRole] = [];
            }

            clickCounts[userRole].push({timestamp: new Date().toISOString()});
            await sp.web.lists.getByTitle(listName).items.getById(pageId).update({
                ClickCounts: JSON.stringify(clickCounts),
            });
            console.log("Click count updated.");

            const updatedPages = this.state.pages.map(page => {
                if (page.id === pageId) {
                    // Calculate clicks from the updated SharePoint data (not just +1)
                    const currentClicks = (clickCounts[userRole] || []).length;
                    return {...page, clicks: currentClicks};
                }
                return page;
            });

            updatedPages.sort((a, b) => b.clicks - a.clicks);

            const slicedPages = updatedPages.slice(0, 11);
            const otherButton = updatedPages.find((page) => page.title === "Övriga System");
            if (otherButton) {
                slicedPages.push(otherButton);
            }

            this.setState({ pages: slicedPages });

        } catch (error) {
            console.error("Error updating click count:", error);
        }
    }
    private async cleanUpOldDataBasedOnLatestEntry(): Promise<void> {
        const sp = this.props.sp;
        const listName = this.props.listName || "KlickPrioritet";

        try {
            // Get all items
            const items = await sp.web.lists
                .getByTitle(listName)
                .items.select("Id", "ClickCounts")();

            for (const item of items) {
                const clickCounts = JSON.parse(item.ClickCounts || "{}");
                let shouldUpdate = false;

                // For each role
                for (const role of Object.keys(clickCounts)) {
                    if (!Array.isArray(clickCounts[role])) continue;

                    // Sort timestamps descending so [0] is the newest
                    clickCounts[role].sort((a: any, b: any) =>
                        new Date(b.timestamp).valueOf() - new Date(a.timestamp).valueOf()
                    );

                    // If there are no timestamps, skip
                    if (clickCounts[role].length === 0) continue;

                    // The latest timestamp in this role
                    const latestTimestamp = new Date(clickCounts[role][0].timestamp);

                    // Our cutoff is (latestTimestamp - 7 days)
                    const cutoff = new Date(latestTimestamp.valueOf() - 7 * 24 * 60 * 60 * 1000);

                    // Filter out any entry older than the cutoff
                    const filtered = clickCounts[role].filter((entry: { timestamp: string }) => {
                        return new Date(entry.timestamp) >= cutoff;
                    });

                    // If the array changed in length, mark item for update
                    if (filtered.length !== clickCounts[role].length) {
                        shouldUpdate = true;
                        clickCounts[role] = filtered;
                    }
                }

                // Only update if something changed
                if (shouldUpdate) {
                    await sp.web.lists.getByTitle(listName)
                        .items.getById(item.Id)
                        .update({
                            ClickCounts: JSON.stringify(clickCounts)
                        });
                }
            }

            console.log("Old data cleaned up based on the latest entry minus 7 days.");
        } catch (error) {
            console.error("Error cleaning up old data:", error);
        }
    }

    public render(): React.ReactElement<IDynamicContentWebPartProps> {
        const {pages} = this.state;

        return (
            <section className={styles.dynamicContentWebPart}>
                {pages.length > 0 ? (
                    pages.map((page) => (
                        <a
                            key={page.id}
                            href={page.url}
                            target="_blank"
                            rel="noopener noreferrer"
                            onClick={() => page.id !== -1 && this.handlePageClick(page.id, this.props.userRole)}
                            className={styles.pageButton}
                        >
                            <div className={styles.icon}>
                                <i className={`ms-Icon ${page.icon}`} aria-hidden="true"/>
                            </div>
                            <div className={styles.title}>
                                {page.title} {}
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
