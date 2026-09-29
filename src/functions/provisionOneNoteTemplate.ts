import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getGraphAccessToken } from "../services/graphClient";
import {
  copyTemplateNotebookToGroup,
  copyTemplateSectionsToGroup,
  TemplateSource,
} from "../services/notebookSectionCopier";
import { findGroupNotebook } from "../services/oneNoteApi";
import { isTabType, listTabs, pinTab, PinTabResult, removeTabs, TabType } from "../services/teamsTab";

/** Extracts the newly created group's id from the EasyLife 365 webhook payload. */
function extractGroupId(body: unknown): string | undefined {
  if (!body || typeof body !== "object") {
    return undefined;
  }
  const record = body as Record<string, unknown>;
  const group = record.group as Record<string, unknown> | undefined;
  const candidate = group?.id ?? record.groupId ?? record.resourceId ?? record.id;
  return typeof candidate === "string" ? candidate : undefined;
}

/** Extracts the newly created group's display name from the EasyLife 365 webhook payload. */
function extractGroupName(body: unknown): string | undefined {
  if (!body || typeof body !== "object") {
    return undefined;
  }
  const record = body as Record<string, unknown>;
  const group = record.group as Record<string, unknown> | undefined;
  const candidate = group?.displayName ?? record.displayName;
  return typeof candidate === "string" ? candidate : undefined;
}

/** OneNote numbers notebooks that share a display name, so names can be made unique. */
function applyPlaceholders(value: string, groupName: string | undefined): string {
  return value.replace(/\{group(?:Name)?\}/gi, groupName ?? "").replace(/\s{2,}/g, " ").trim();
}

function splitNames(value: string | null | undefined): string[] {
  return (value ?? "")
    .split(",")
    .map((name) => name.trim())
    .filter((name) => name.length > 0);
}

/** Accepts singular and plural spellings, repeated query parameters, and comma-separated lists. */
function readList(request: HttpRequest, queryNames: string[], envNames: string[]): string[] {
  const fromQuery = queryNames.flatMap((name) => request.query.getAll(name).flatMap(splitNames));
  if (fromQuery.length) {
    return fromQuery;
  }
  return envNames.flatMap((name) => splitNames(process.env[name]));
}

function buildSectionMappings(from: string[], to: string[]): { from: string; to: string }[] {
  return from.map((name, index) => ({ from: name, to: to[index] ?? name }));
}

const TRUE_VALUES = ["true", "1", "yes", "ja", "on"];
const FALSE_VALUES = ["false", "0", "no", "nein", "off"];

/** The switch doubles as the tab name, so "createTeamsTab=true" reuses the notebook name. */
function resolveTabName(values: string[], fallback: string): string | undefined {
  const value = values[0];
  if (!value || FALSE_VALUES.includes(value.toLowerCase())) {
    return undefined;
  }
  return TRUE_VALUES.includes(value.toLowerCase()) ? fallback : value;
}

export async function provisionOneNoteTemplate(
  request: HttpRequest,
  context: InvocationContext
): Promise<HttpResponseInit> {
  const body = await request.json().catch(() => undefined);
  context.log("EasyLife webhook payload", JSON.stringify(body));

  const targetGroupId = extractGroupId(body);
  if (!targetGroupId) {
    context.warn("No group id found in webhook payload.");
    return { status: 400, jsonBody: { error: "Could not determine groupId from webhook payload." } };
  }

  // Everything below can be set per EasyLife automation step via the webhook URL query string.
  const templateSiteUrls = readList(
    request,
    ["templateSiteUrl", "templateSiteUrls", "templateSite", "templateSites"],
    ["DEFAULT_TEMPLATE_SITE_URL", "DEFAULT_TEMPLATE_SITE_URLS"]
  );
  const templateGroupIds = readList(
    request,
    ["templateGroupId", "templateGroupIds", "templateGroup", "templateGroups"],
    ["DEFAULT_TEMPLATE_GROUP_ID", "DEFAULT_TEMPLATE_GROUP_IDS"]
  );
  const notebookNames = readList(
    request,
    ["templateNotebookName", "templateNotebookNames", "templateNotebook", "templateNotebooks"],
    ["DEFAULT_TEMPLATE_NOTEBOOK_NAME", "DEFAULT_TEMPLATE_NOTEBOOK_NAMES"]
  );

  const templateSectionNames = readList(
    request,
    ["templateSectionName", "templateSectionNames", "templateSection", "templateSections"],
    ["DEFAULT_TEMPLATE_SECTION_NAMES", "DEFAULT_TEMPLATE_SECTION_NAME"]
  );
  const targetSectionNames = readList(
    request,
    ["targetSectionName", "targetSectionNames", "targetSection", "targetSections"],
    ["DEFAULT_TARGET_SECTION_NAMES", "DEFAULT_TARGET_SECTION_NAME"]
  );

  // Setting a target notebook switches from copying sections to cloning the whole notebook.
  const targetNotebookNames = readList(
    request,
    ["targetNotebookName", "targetNotebookNames", "targetNotebook", "targetNotebooks"],
    ["DEFAULT_TARGET_NOTEBOOK_NAME", "DEFAULT_TARGET_NOTEBOOK_NAMES"]
  );
  const tabValues = readList(
    request,
    [
      "createTeamsTab",
      "createTab",
      "pinTab",
      "addTab",
      "tabName",
      "tabNames",
      "teamsTabName",
      "teamsTabNames",
    ],
    ["DEFAULT_CREATE_TEAMS_TAB", "DEFAULT_TAB_NAME", "DEFAULT_TAB_NAMES"]
  );
  const tabTypeValue = readList(request, ["tabType", "teamsTabType"], ["DEFAULT_TAB_TYPE"])[0] ?? "onenote";
  const tabType: TabType = isTabType(tabTypeValue) ? tabTypeValue : "onenote";
  const tabNotebookId = readList(request, ["tabNotebookId", "notebookId"], ["DEFAULT_TAB_NOTEBOOK_ID"])[0];
  const removeTabNames = readList(
    request,
    ["removeTab", "removeTabs", "removeTabName", "removeTabNames"],
    ["DEFAULT_REMOVE_TAB_NAMES"]
  );
  const targetLibrary = readList(request, ["targetLibrary", "library"], ["DEFAULT_TARGET_LIBRARY"])[0];
  const targetFolder = readList(request, ["targetFolder", "folder"], ["DEFAULT_TARGET_FOLDER"])[0];

  const sources: TemplateSource[] = [
    ...templateSiteUrls.map((siteUrl) => ({ kind: "site" as const, siteUrl, notebookNames })),
    ...templateGroupIds.map((groupId) => ({ kind: "group" as const, groupId, notebookNames })),
  ];

  if (!sources.length) {
    context.warn(
      "No template source configured. Set templateSiteUrl/templateGroupId in the webhook URL or DEFAULT_TEMPLATE_SITE_URL/DEFAULT_TEMPLATE_GROUP_ID in the app settings."
    );
    return {
      status: 400,
      jsonBody: { error: "No template source configured. Provide templateSiteUrl or templateGroupId." },
    };
  }

  context.log(
    `Copying into group ${targetGroupId} from ${sources.length} template source(s)`,
    JSON.stringify({
      templateSiteUrls,
      templateGroupIds,
      notebookNames,
      templateSectionNames,
      targetSectionNames,
      targetNotebookNames,
      tabValues,
      tabType,
      tabNotebookId,
      removeTabNames,
      targetLibrary,
      targetFolder,
    })
  );

  try {
    const token = await getGraphAccessToken();

    if (targetNotebookNames.length) {
      let tabsRemoved: string[] | undefined;

      // Runs before the clone so the empty notebook disappears as early as possible.
      if (removeTabNames.length) {
        try {
          tabsRemoved = await removeTabs(targetGroupId, removeTabNames, token);
        } catch (err) {
          context.warn("Could not remove Teams tabs", (err as Error).message);
        }
      }

      const groupName = extractGroupName(body);
      const notebookName = applyPlaceholders(targetNotebookNames[0], groupName);
      const result = await copyTemplateNotebookToGroup({
        token,
        sources,
        targetGroupId,
        notebookName,
        libraryName: targetLibrary,
        folderPath: targetFolder,
      });

      let tabPinned: string | undefined;
      let tabError: string | undefined;
      let oneNoteApiError: string | undefined;
      let tabNames: PinTabResult | undefined;

      const tabName = resolveTabName(tabValues, notebookName);
      if (tabName) {
        // Reveals whether app-only access to the OneNote API is possible in this tenant.
        let oneNote: { notebookId: string; notebookName: string; webUrl: string } | undefined;
        if (tabType === "onenote") {
          try {
            const notebook = await findGroupNotebook(targetGroupId, notebookName, token);
            oneNote = { notebookId: notebook.id, notebookName: notebook.displayName, webUrl: notebook.webUrl };
          } catch (err) {
            oneNoteApiError = (err as Error).message;
            context.warn("OneNote API unavailable, falling back", oneNoteApiError);
          }

          if (!oneNote && tabNotebookId) {
            // OneNote ids of SharePoint notebooks are the sourcedoc id with a "1-" prefix.
            const notebookId =
              tabNotebookId.toLowerCase() === "auto" ? `1-${result.notebookItemId}` : tabNotebookId;
            oneNote = {
              notebookId,
              notebookName,
              webUrl: result.notebookUrl,
            };
          }
        }

        try {
          tabNames = await pinTab({
            groupId: targetGroupId,
            displayName: tabName,
            tabType,
            contentUrl: tabType === "library" ? result.notebookFolderUrl : result.notebookEmbedUrl,
            websiteUrl: result.notebookUrl,
            entityId: result.notebookItemId,
            oneNote,
            token,
          });
          tabPinned = tabNames.finalName ?? tabName;
        } catch (err) {
          // The notebook is already in place, so a missing tab permission must not fail the run.
          tabError = (err as Error).message;
          context.warn("Could not pin the Teams tab", tabError);
        }
      }

      const tabsInChannel = await listTabs(targetGroupId, token);

      context.log(
        `Cloned notebook into group ${targetGroupId}.`,
        JSON.stringify({ ...result, tabPinned, tabError, oneNoteApiError, tabsRemoved, tabNames, tabsInChannel })
      );
      return {
        status: 200,
        jsonBody: {
          status: "ok",
          ...result,
          tabPinned,
          tabError,
          oneNoteApiError,
          tabsRemoved,
          tabNames,
          tabsInChannel,
        },
      };
    }

    const result = await copyTemplateSectionsToGroup({
      token,
      sources,
      sections: buildSectionMappings(templateSectionNames, targetSectionNames),
      targetGroupId,
    });

    context.log(`Copied ${result.sectionsCopied.length} section(s) into group ${targetGroupId}.`, JSON.stringify(result));
    return { status: 200, jsonBody: { status: "ok", ...result } };
  } catch (err) {
    context.error("Failed to copy OneNote template", err);
    return { status: 500, jsonBody: { error: (err as Error).message } };
  }
}

app.http("provisionOneNoteTemplate", {
  methods: ["POST"],
  authLevel: "function",
  route: "onenote-template",
  handler: provisionOneNoteTemplate,
});
