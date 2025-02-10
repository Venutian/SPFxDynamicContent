import * as React from 'react';
import { IDynamicContentWebPartProps, ILinkItem } from './IDynamicContentWebPartProps';
import styles from './DynamicContentWebPart.module.scss';
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/webs/index";
import "@pnp/sp/fields/list";

// Define an interface for each click entry
interface IClickEntry {
    timestamp: string;
}

// Define a type for the ClickCounts object mapping roles to arrays of click entries
type IClickCounts = {
    [role: string]: IClickEntry[];
};

interface IDynamicContentWebPartState {
    pages: ILinkItem[];
}

export default class DynamicContentComponent extends React.Component<
    IDynamicContentWebPartProps,
    IDynamicContentWebPartState
> {
    constructor(props: IDynamicContentWebPartProps) {
        super(props);
        this.state = { pages: [] };
    }

    public async componentDidMount(): Promise<void> {
        if (!this.props.sp) {
            console.error("SharePoint instance (sp) is undefined. Please ensure that 'sp' is passed as a prop.");
            return;
        }
        await this.cleanUpOldDataBasedOnLatestEntry();
        await this.loadPages();
    }

    /**
     * Cleans up old click count entries for each item.
     * For each role’s click entries, it sorts them (newest first)
     * and removes any entry older than 7 days before the latest entry.
     */
    private async cleanUpOldDataBasedOnLatestEntry(): Promise<void> {
        const sp = this.props.sp;
        if (!sp) {
            console.error("SharePoint instance (sp) is undefined in cleanUpOldDataBasedOnLatestEntry.");
            return;
        }
        const listName = this.props.listName || "KlickPrioritet";

        try {
            const items = await sp.web.lists
                .getByTitle(listName)
                .items.select("Id", "ClickCounts")();

            for (const item of items) {
                const clickCounts: IClickCounts = JSON.parse(item.ClickCounts || "{}");
                let shouldUpdate = false;

                // Iterate over each role’s click entries
                for (const role in clickCounts) {
                    if (!Array.isArray(clickCounts[role])) continue;

                    // Sort click entries so that index 0 is the newest
                    clickCounts[role].sort((a: IClickEntry, b: IClickEntry) =>
                        new Date(b.timestamp).valueOf() - new Date(a.timestamp).valueOf()
                    );

                    if (clickCounts[role].length === 0) continue;

                    // Use the latest timestamp as reference and set cutoff to 7 days before
                    const latestTimestamp = new Date(clickCounts[role][0].timestamp);
                    const cutoff = new Date(latestTimestamp.valueOf() - 7 * 24 * 60 * 60 * 1000);

                    // Filter out any entries older than the cutoff
                    const filtered = clickCounts[role].filter((entry: IClickEntry) =>
                        new Date(entry.timestamp) >= cutoff
                    );

                    if (filtered.length !== clickCounts[role].length) {
                        shouldUpdate = true;
                        clickCounts[role] = filtered;
                    }
                }

                // If any role’s entries were updated, write the changes back to SharePoint
                if (shouldUpdate) {
                    await sp.web.lists
                        .getByTitle(listName)
                        .items.getById(item.Id)
                        .update({
                            ClickCounts: JSON.stringify(clickCounts),
                        });
                }
            }

            console.log("Old data cleaned up (latest entry minus 7 days).");
        } catch (error) {
            console.error("Error cleaning up old data:", error);
        }
    }

    /**
     * Loads pages from SharePoint, calculates click counts based on the user's role,
     * and prepares the list of pages. The "Övriga System" button is separated out
     * so it is not sorted or ranked and appears only once at the end.
     */
    private async loadPages(): Promise<void> {
        const sp = this.props.sp;
        if (!sp) {
            console.error("SharePoint instance (sp) is undefined in loadPages.");
            return;
        }
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
                const clickCounts: IClickCounts = JSON.parse(item.ClickCounts || "{}");
                const totalClicks = (clickCounts[userRole] || []).length;
                return {
                    id: item.Id,
                    title: item.Title,
                    url: item.URL,
                    clicks: totalClicks,
                    roles: item.Roles.split(","),
                    icon: item.Icon,
                } as ILinkItem;
            });

            // Filter pages visible to the current user
            const roleFilteredPages = pages.filter((page) => page.roles.includes(userRole));
            // Separate the "Övriga System" button from ranked pages
            const otherButton = roleFilteredPages.find((page) => page.title === "Övriga System");
            const rankedPages = roleFilteredPages.filter((page) => page.title !== "Övriga System");

            // Sort ranked pages descending by click count
            rankedPages.sort((a, b) => b.clicks - a.clicks);

            // Limit the ranked pages (e.g., top 11) and then append the "Övriga System" button if it exists
            const displayedPages = rankedPages.slice(0, 11);
            if (otherButton) {
                displayedPages.push(otherButton);
            }

            this.setState({ pages: displayedPages });
        } catch (error) {
            console.error("Error fetching live data:", error);
        }
    }

    /**
     * Handles a page click by updating its click count in SharePoint
     * (if it is not the "Övriga System" button) and then updating local state.
     */
    private async handlePageClick(pageId: number, userRole: string): Promise<void> {
        const sp = this.props.sp;
        if (!sp) {
            console.error("SharePoint instance (sp) is undefined in handlePageClick.");
            return;
        }
        const listName = this.props.listName || "KlickPrioritet";

        try {
            const item = await sp.web.lists
                .getByTitle(listName)
                .items.getById(pageId)
                .select("ClickCounts")();

            const clickCounts: IClickCounts = JSON.parse(item.ClickCounts || "{}");
            if (!clickCounts[userRole]) {
                clickCounts[userRole] = [];
            }

            // Add a new click entry with the current timestamp
            clickCounts[userRole].push({ timestamp: new Date().toISOString() });

            await sp.web.lists.getByTitle(listName).items.getById(pageId).update({
                ClickCounts: JSON.stringify(clickCounts),
            });
            console.log("Click count updated.");

            // Update local state with the new click count for the clicked page
            const updatedPages = this.state.pages.map((page) => {
                if (page.id === pageId) {
                    const currentClicks = (clickCounts[userRole] || []).length;
                    return { ...page, clicks: currentClicks };
                }
                return page;
            });

            // Re-sort the ranked pages and then re-append the "Övriga System" button
            const rankedPages = updatedPages.filter((page) => page.title !== "Övriga System");
            rankedPages.sort((a, b) => b.clicks - a.clicks);
            const slicedPages = rankedPages.slice(0, 11);
            const otherButton = updatedPages.find((page) => page.title === "Övriga System");
            if (otherButton) {
                slicedPages.push(otherButton);
            }

            this.setState({ pages: slicedPages });
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
                            onClick={(e) => {
                                e.preventDefault();
                                // If this is the "Övriga System" button, simply open the URL without updating clicks
                                if (page.title === "Övriga System") {
                                    window.open(page.url, '_blank');
                                } else {
                                    this.handlePageClick(page.id, this.props.userRole)
                                        .then(() => {
                                            window.open(page.url, '_blank');
                                        })
                                        .catch((error) => {
                                            console.error('Error during click handling:', error);
                                            window.open(page.url, '_blank');
                                        });
                                }
                            }}
                            className={styles.pageButton}
                        >
                            <div className={styles.icon}>
                                <i className={`ms-Icon ${page.icon}`} aria-hidden="true" />
                            </div>
                            <div className={styles.title}>{page.title}</div>
                        </a>
                    ))
                ) : (
                    <p>No pages available to display.</p>
                )}
            </section>
        );
    }
}
