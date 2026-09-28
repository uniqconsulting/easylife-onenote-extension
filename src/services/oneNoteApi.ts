import { graphFetch } from "./graphClient";

export interface OneNoteNotebook {
  id: string;
  displayName: string;
  webUrl: string;
}

interface NotebookResponse {
  value: { id: string; displayName: string; links?: { oneNoteWebUrl?: { href?: string } } }[];
}

/**
 * Looks up a group notebook through the OneNote API. App-only tokens may be rejected with
 * error 40001, so callers must handle the failure instead of relying on this.
 */
export async function findGroupNotebook(
  groupId: string,
  notebookName: string,
  token: string
): Promise<OneNoteNotebook> {
  const response = await graphFetch(`/groups/${groupId}/onenote/notebooks?$select=id,displayName,links`, token);
  if (!response.ok) {
    throw new Error(`Listing OneNote notebooks of group ${groupId} failed: ${response.status} ${await response.text()}`);
  }

  const body = (await response.json()) as NotebookResponse;
  const match =
    body.value.find((n) => n.displayName?.toLowerCase() === notebookName.toLowerCase()) ?? body.value[0];

  if (!match) {
    throw new Error(`Group ${groupId} exposes no OneNote notebook named "${notebookName}".`);
  }

  return {
    id: match.id,
    displayName: match.displayName,
    webUrl: match.links?.oneNoteWebUrl?.href ?? "",
  };
}
