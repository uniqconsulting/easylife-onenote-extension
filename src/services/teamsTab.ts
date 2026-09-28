import { graphFetch, sleep } from "./graphClient";

// Teams renders these natively. A website tab only embeds pages that allow framing, which
// SharePoint does not, so such a tab always opens the browser instead.
export const TAB_APPS = {
  website: "com.microsoft.teamspace.tab.web",
  library: "com.microsoft.teamspace.tab.files.sharepoint",
  onenote: "0d820ecd-def2-4297-adad-78056cde7c78",
} as const;

export type TabType = keyof typeof TAB_APPS;

export function isTabType(value: string): value is TabType {
  return value in TAB_APPS;
}

export interface PinTabOptions {
  groupId: string;
  displayName: string;
  tabType: TabType;
  /** Doc.aspx embed view for website tabs, the folder url for library tabs. */
  contentUrl: string;
  websiteUrl: string;
  entityId: string;
  /** Only available when the OneNote API accepted the token. */
  oneNote?: { notebookId: string; notebookName: string; webUrl: string };
  token: string;
}

// Teams provisions the team itself after the group, so the channel is not available right away.
const ATTEMPTS = 6;
const DELAY_MS = 5000;

function buildBody(options: PinTabOptions): Record<string, unknown> {
  const { displayName, tabType, contentUrl, websiteUrl, entityId, groupId, oneNote } = options;
  const body: Record<string, unknown> = {
    displayName,
    "teamsApp@odata.bind": `https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/${TAB_APPS[tabType]}`,
  };

  if (tabType === "website") {
    body.configuration = { entityId: "", contentUrl, websiteUrl, removeUrl: null };
  } else if (tabType === "library") {
    body.configuration = { entityId, contentUrl, websiteUrl: null, removeUrl: null };
  } else if (oneNote) {
    const selfUrl = `https://www.onenote.com/api/v1.0/myOrganization/groups/${groupId}/notes/notebooks/${oneNote.notebookId}`;
    const tabContentUrl =
      `https://www.onenote.com/teams/TabContent?entityid=%7BentityId%7D&subentityid=%7BsubEntityId%7D` +
      `&auth_upn=%7Bupn%7D&notebookSource=Pick&notebookSelfUrl=${encodeURIComponent(selfUrl)}` +
      `&oneNoteWebUrl=${encodeURIComponent(oneNote.webUrl)}&notebookName=${encodeURIComponent(oneNote.notebookName)}` +
      `&ui={locale}&tenantId={tid}`;

    body.configuration = {
      entityId: oneNote.notebookId,
      contentUrl: tabContentUrl,
      removeUrl: null,
      websiteUrl: `https://www.onenote.com/teams/TabRedirect?redirectUrl=${encodeURIComponent(oneNote.webUrl)}`,
    };
  }
  // Without OneNote ids the tab stays unconfigured and the first user picks the notebook once.

  return body;
}

/** Pins the copied notebook as a tab in the team's primary channel. */
export async function pinTab(options: PinTabOptions): Promise<void> {
  const { groupId, displayName, token } = options;
  let lastError = "unknown error";

  for (let attempt = 0; attempt < ATTEMPTS; attempt++) {
    const channelResponse = await graphFetch(`/teams/${groupId}/primaryChannel?$select=id`, token);

    if (channelResponse.ok) {
      const channel = (await channelResponse.json()) as { id: string };
      const response = await graphFetch(`/teams/${groupId}/channels/${channel.id}/tabs`, token, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(buildBody(options)),
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
