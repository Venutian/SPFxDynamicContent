import * as React from 'react';
import { IDynamicContentWebPartProps } from './IDynamicContentWebPartProps';
import styles from './DynamicContentWebPart.module.scss';
import { ILinkItem } from './IDynamicContentWebPartProps';
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
        this.state = { pages: [] };
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
        const sp = this.props.sp;

        try {
            console.log("Checking if list exists:", this.props.listName || "DailyClickCounts");
            const list = await sp.web.lists
                .getByTitle(this.props.listName || "DailyClickCounts")
                .select("Title")();

            console.log("List exists:", list);
            await this.ensureListColumnsExist();
            await this.insertSampleDataIfEmpty();
        } catch (error) {
            console.error("Error checking list existence:", error);
            console.log("List does not exist. Creating...");

            try {
                const newList = await sp.web.lists.add(
                    this.props.listName || "DailyClickCounts",
                    "Stores click counts for pages",
                    100, // Template type (100 = custom list)
                    false // Enable content types (false)
                );
                console.log("List created:", newList);

                await this.ensureListColumnsExist();
                await this.insertSampleData();
            } catch (createError) {
                console.error("Error creating list:", createError);
            }
        }
    }

    private async ensureListColumnsExist(): Promise<void> {
        const sp = this.props.sp;
        const listTitle = this.props.listName || "DailyClickCounts";
        const list = sp.web.lists.getByTitle(listTitle);

        // URL column
        try {
            await list.fields.getByTitle("URL")();
            console.log("URL column already exists.");
        } catch (error) {
            console.log("URL column does not exist. Creating...", error);
            await list.fields.addText("URL", {
                Group: "Custom Columns",
                Description: "Page URL",
            });
            console.log("URL column created.");
        }

        // ClickCounts column (multiline for JSON data)
        try {
            await list.fields.getByTitle("ClickCounts")();
            console.log("ClickCounts column already exists.");
        } catch (error) {
            console.log("ClickCounts column does not exist. Creating...", error);
            await list.fields.addMultilineText("ClickCounts", {
                Group: "Custom Columns",
                Description: "Stores click counts in JSON format",
            });
            console.log("ClickCounts column created.");
        }

        // Roles column
        try {
            await list.fields.getByTitle("Roles")();
            console.log("Roles column already exists.");
        } catch (error) {
            console.log("Roles column does not exist. Creating...", error);
            await list.fields.addText("Roles", {
                Group: "Custom Columns",
                Description: "Comma-separated roles",
            });
            console.log("Roles column created.");
        }
    }

    private async insertSampleDataIfEmpty(): Promise<void> {
        const sp = this.props.sp;
        const listTitle = this.props.listName || "DailyClickCounts";

        const items = await sp.web.lists.getByTitle(listTitle).items.select("Id")();
        if (items.length === 0) {
            console.log("List is empty. Inserting sample data...");
            await this.insertSampleData();
        }
    }

    private async insertSampleData(): Promise<void> {
        const sp = this.props.sp;
        const listTitle = this.props.listName || "DailyClickCounts";

        const sampleItem = {
            Title: "Admin Dashboard",
            URL: "/sites/admin",
            ClickCounts: JSON.stringify({
                Admin: [{ timestamp: new Date().toISOString() }],
            }),
            Roles: "Admin,User",
        };

        await sp.web.lists.getByTitle(listTitle).items.add(sampleItem);
        console.log("Sample data inserted.");
    }

    private async cleanUpOldData(): Promise<void> {
        const sp = this.props.sp;
        const listName = this.props.listName || "DailyClickCounts";
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
        const listName = this.props.listName || "DailyClickCounts";
        const userRole = this.props.userRole || "Admin";

        if (!userRole) {
            console.error("User role is missing.");
            return;
        }

        console.log("Fetching live SharePoint data...");

        try {
            const items = await sp.web.lists
                .getByTitle(listName)
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

            const roleFilteredPages = pages.filter((page) =>
                page.roles.includes(userRole)
            );
            roleFilteredPages.sort((a, b) => b.clicks - a.clicks);

            this.setState({ pages: roleFilteredPages });
        } catch (error) {
            console.error("Error fetching live data:", error);
        }
    }

    private async handlePageClick(pageId: number, userRole: string): Promise<void> {
        const sp = this.props.sp;
        const listName = this.props.listName || "DailyClickCounts";

        try {
            const item = await sp.web.lists
                .getByTitle(listName)
                .items.getById(pageId)
                .select("ClickCounts")();

            const clickCounts = JSON.parse(item.ClickCounts || "{}");
            if (!clickCounts[userRole]) {
                clickCounts[userRole] = [];
            }

            clickCounts[userRole].push({ timestamp: new Date().toISOString() });
            await sp.web.lists.getByTitle(listName).items.getById(pageId).update({
                ClickCounts: JSON.stringify(clickCounts),
            });
            console.log("Click count updated.");

            const updatedPages = this.state.pages.map((page) => {
                if (page.id === pageId) {
                    return { ...page, clicks: page.clicks + 1 };
                }
                return page;
            });

            this.setState({ pages: updatedPages });
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
                                <i className="ms-Icon ms-Icon--Globe" aria-hidden="true" />
                            </div>
                            <div className={styles.title}>
                                {page.title}
                                <br /> ({page.clicks} clicks)
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
