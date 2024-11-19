import * as React from 'react';
import { IDynamicContentWebPartProps } from './IDynamicContentWebPartProps';
import styles from './DynamicContentWebPart.module.scss';
import { ILinkItem } from './IDynamicContentWebPartProps';
import { Web } from "@pnp/sp/webs"; // Import Web for accessing SharePoint lists
import "@pnp/sp/lists"; // Import List functionality explicitly
import "@pnp/sp/items"; // Import Item functionality explicitly

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

  public async componentDidMount() {
    await this.ensureListExists(); // Create the list if it doesn't exist
    await this.cleanUpOldData(); // Clean up old data
    await this.loadPages(); // Load pages

    // Set up daily updates
    this.updateInterval = setInterval(async () => {
      await this.cleanUpOldData(); // Clean up old data
      await this.loadPages(); // Refresh pages
    }, 24 * 60 * 60 * 1000); // Every 24 hours
  }

  public componentWillUnmount() {
    if (this.updateInterval) {
      clearInterval(this.updateInterval);
    }
  }

  // Check if the list exists, and create it if not
  private async ensureListExists() {
    const webUrl = this.props.context.pageContext.web.absoluteUrl;
    const web = Web(webUrl);

    try {
      // Check if the list exists
      await web.lists.getByTitle(this.props.listName || "DailyClickCounts").select("Title")();
      console.log("List exists.");
    } catch (error) {
      console.log("List does not exist. Creating...");
      // Create the list
      await web.lists.add(this.props.listName || "DailyClickCounts", "Stores click counts for pages", 100);
      console.log("List created.");
    }
  }

  // Clean up data older than 7 days
  private async cleanUpOldData() {
    const webUrl = this.props.context.pageContext.web.absoluteUrl;
    const web = Web(webUrl);

    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    const sevenDaysAgoISO = sevenDaysAgo.toISOString();

    try {
      const oldItems = await web.lists
          .getByTitle(this.props.listName || "DailyClickCounts")
          .items.filter(`LastUpdated lt datetime'${sevenDaysAgoISO}'`)
          .select("Id")();

      for (const item of oldItems) {
        await web.lists.getByTitle(this.props.listName || "DailyClickCounts").items.getById(item.Id).delete();
      }

      console.log("Old data cleaned up.");
    } catch (error) {
      console.error("Error cleaning up old data:", error);
    }
  }

  // Load pages and sort by popularity
  private async loadPages() {
    const userRole = this.props.userRole || "DefaultRole";
    const webUrl = this.props.context.pageContext.web.absoluteUrl;
    const web = Web(webUrl);

    const sevenDaysAgo = new Date();
    sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
    const sevenDaysAgoISO = sevenDaysAgo.toISOString();

    try {
      const items = await web.lists
          .getByTitle(this.props.listName || "DailyClickCounts")
          .items.filter(`LastUpdated ge datetime'${sevenDaysAgoISO}'`)
          .select("Id", "Title", "URL", "ClickCounts", "Roles")();

      const pages = items.map((item) => {
        const clickCounts = JSON.parse(item.ClickCounts || "{}");
        const clicksForRole = clickCounts[userRole] || 0;
        return {
          id: item.Id,
          title: item.Title,
          url: item.URL.Url,
          clicks: clicksForRole,
          roles: item.Roles.split(",")
        };
      }) as ILinkItem[];

      const roleFilteredPages = pages.filter((page) => page.roles.includes(userRole));
      roleFilteredPages.sort((a, b) => b.clicks - a.clicks);

      this.setState({ pages: roleFilteredPages });
    } catch (error) {
      console.error("Error loading pages:", error);
    }
  }

  // Handle page clicks and update ClickCounts
  private async handlePageClick(pageId: number, userRole: string) {
    const webUrl = this.props.context.pageContext.web.absoluteUrl;
    const web = Web(webUrl);

    try {
      const item = await web.lists
          .getByTitle(this.props.listName || "DailyClickCounts")
          .items.getById(pageId)
          .select("ClickCounts")();

      const clickCounts = JSON.parse(item.ClickCounts || "{}");
      clickCounts[userRole] = (clickCounts[userRole] || 0) + 1;

      await web.lists.getByTitle(this.props.listName || "DailyClickCounts").items.getById(pageId).update({
        ClickCounts: JSON.stringify(clickCounts)
      });

      console.log("Click count updated.");
    } catch (error) {
      console.error("Error updating click count:", error);
    }
  }

  public render(): React.ReactElement<IDynamicContentWebPartProps> {
    const { pages } = this.state;

    return (
        <section className={styles.dynamicContentWebPart}>
          <div>
            <h2>Popular Pages</h2>
            <ul>
              {pages.map((page) => (
                  <li key={page.id}>
                    <a
                        href={page.url}
                        target="_blank"
                        rel="noopener noreferrer"
                        onClick={() => this.handlePageClick(page.id, this.props.userRole)}
                    >
                      {page.title} ({page.clicks} clicks)
                    </a>
                  </li>
              ))}
            </ul>
          </div>
        </section>
    );
  }
}
