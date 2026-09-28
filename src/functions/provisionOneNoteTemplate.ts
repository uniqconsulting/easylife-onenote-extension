import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { getGraphAccessToken } from "../services/graphClient";
import {
  copyTemplateNotebookToGroup,
  copyTemplateSectionsToGroup,
  TemplateSource,
} from "../services/notebookSectionCopier";
import { pinWebsiteTab } from "../services/teamsTab";

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
  const tabNames = readList(
    request,
    ["tabName", "tabNames", "teamsTabName", "teamsTabNames"],
    ["DEFAULT_TAB_NAME", "DEFAULT_TAB_NAMES"]
  );

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
      tabNames,
    })
  );

  try {
    const token = await getGraphAccessToken();

    if (targetNotebookNames.length) {
      const result = await copyTemplateNotebookToGroup({
        token,
        sources,
        targetGroupId,
        notebookName: targetNotebookNames[0],
      });

      let tabPinned: string | undefined;
      let tabError: string | undefined;
      if (tabNames.length) {
        try {
          await pinWebsiteTab(targetGroupId, tabNames[0], result.notebookUrl, token);
          tabPinned = tabNames[0];
        } catch (err) {
          // The notebook is already in place, so a missing tab permission must not fail the run.
          tabError = (err as Error).message;
          context.warn("Could not pin the Teams tab", tabError);
        }
      }

      context.log(`Cloned notebook into group ${targetGroupId}.`, JSON.stringify({ ...result, tabPinned, tabError }));
      return { status: 200, jsonBody: { status: "ok", ...result, tabPinned, tabError } };
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
