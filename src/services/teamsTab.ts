import { graphFetch, sleep } from "./graphClient";

// The generic website tab renders the SharePoint OneNote view, which shows the section list.
const WEBSITE_TAB_APP_ID = "com.microsoft.teamspace.tab.web";

// Teams provisions the team itself after the group, so the channel is not available right away.
const ATTEMPTS = 6;
const DELAY_MS = 5000;

/** Pins the copied notebook as a tab in the team's primary channel. */
export async function pinWebsiteTab(
  groupId: string,
  displayName: string,
  url: string,
  token: string
): Promise<void> {
  let lastError = "unknown error";

  for (let attempt = 0; attempt < ATTEMPTS; attempt++) {
    const channelResponse = await graphFetch(`/teams/${groupId}/primaryChannel?$select=id`, token);

    if (channelResponse.ok) {
      const channel = (await channelResponse.json()) as { id: string };
      const response = await graphFetch(`/teams/${groupId}/channels/${channel.id}/tabs`, token, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          displayName,
          "teamsApp@odata.bind": `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${WEBSITE_TAB_APP_ID}`,
          configuration: { entityId: "", contentUrl: url, websiteUrl: url, removeUrl: "" },
        }),
      });

      if (response.ok) {
        return;
      }
      lastError = `${response.status} ${await response.text()}`;
    } else {
      lastError = `${channelResponse.status} ${await channelResponse.text()}`;
    }

    if (attempt < ATTEMPTS - 1) {
      await sleep(DELAY_MS);
    }
  }

  throw new Error(`Pinning the Teams tab "${displayName}" failed: ${lastError}`);
}
