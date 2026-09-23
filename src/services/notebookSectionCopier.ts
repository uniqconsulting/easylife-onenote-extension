import { graphFetch, sleep } from "./graphClient";

// The OneNote API rejects app-only tokens (Graph error 40001), so sections are copied as
// the underlying .one files through the SharePoint Drive API instead.

export type TemplateSource =
  | { kind: "site"; siteUrl: string; notebookNames: string[] }
  | { kind: "group"; groupId: string; notebookNames: string[] };

export interface CopyTemplateOptions {
  token: string;
  sources: TemplateSource[];
  sections: { from: string; to: string }[];
  targetGroupId: string;
}

export interface CopyTemplateResult {
  sectionsCopied: { from: string; to: string }[];
  templateNotebooks: string[];
  targetNotebook: string;
  filesInTargetNotebook: string[];
}

interface DriveItem {
  id: string;
  name: string;
  folder?: { childCount?: number };
  file?: { mimeType?: string };
}

interface DriveRef {
  id: string;
  name: string;
}

interface NotebookLocation {
  driveId: string;
  driveName: string;
  folderId: string;
  folderName: string;
}

interface TemplateSection {
  notebook: NotebookLocation;
  item: DriveItem;
}

const SECTION_EXTENSION = ".one";
const NOTEBOOK_MARKER_EXTENSION = ".onetoc2";

// Teams pins its channel tab to the notebook's default section, whose name is localised.
export const DEFAULT_SECTION_TOKEN = "@default";
// Keeps the template's own section instead of merging into an existing one.
export const SOURCE_SECTION_TOKEN = "@source";
const DEFAULT_SECTION_NAMES = [
  "general",
  "allgemein",
  "général",
  "generale",
  "algemeen",
  "generelt",
  "allmänt",
  "yleinen",
  "ogólny",
  "obecné",
  "általános",
  "genel",
];

// EasyLife fires the webhook before SharePoint has finished provisioning the group notebook.
const PROVISIONING_ATTEMPTS = 6;
const PROVISIONING_DELAY_MS = 5000;

function normalizeName(value: string): string {
  return value.trim().replace(/\.one$/i, "").toLowerCase();
}

async function getJson<T>(path: string, token: string, what: string): Promise<T> {
  const response = await graphFetch(path, token);
  if (!response.ok) {
    throw new Error(`${what} failed: ${response.status} ${await response.text()}`);
  }
  return (await response.json()) as T;
}

async function resolveSiteIdFromUrl(siteUrl: string, token: string): Promise<string> {
  const url = new URL(siteUrl);
  const sitePath = url.pathname.replace(/\/+$/, "");
  const site = await getJson<{ id: string }>(
    `/sites/${url.hostname}:${sitePath}?$select=id`,
    token,
    `Resolving SharePoint site ${siteUrl}`
  );
  return site.id;
}

async function resolveSiteIdFromGroup(groupId: string, token: string): Promise<string> {
  const site = await getJson<{ id: string }>(
    `/groups/${groupId}/sites/root?$select=id`,
    token,
    `Resolving site of group ${groupId}`
  );
  return site.id;
}

async function resolveSiteId(source: TemplateSource, token: string): Promise<string> {
  return source.kind === "site"
    ? resolveSiteIdFromUrl(source.siteUrl, token)
    : resolveSiteIdFromGroup(source.groupId, token);
}

/** Some tenants answer this endpoint without an id, which used to produce "/drives/undefined". */
async function resolveListDrive(
  siteId: string,
  list: string,
  fallbackName: string,
  token: string
): Promise<DriveRef | undefined> {
  const response = await graphFetch(`/sites/${siteId}/lists/${list}/drive`, token);
  if (!response.ok) {
    return undefined;
  }

  const drive = (await response.json().catch(() => undefined)) as { id?: string; name?: string } | undefined;
  return drive?.id ? { id: drive.id, name: drive.name ?? fallbackName } : undefined;
}

async function listDrives(siteId: string, token: string): Promise<DriveRef[]> {
  const drives = new Map<string, DriveRef>();

  const direct = await getJson<{ value: { id?: string; name?: string }[] }>(
    `/sites/${siteId}/drives?$select=id,name`,
    token,
    `Listing document libraries of site ${siteId}`
  );
  for (const drive of direct.value) {
    if (drive.id) {
      drives.set(drive.id, { id: drive.id, name: drive.name ?? drive.id });
    }
  }

  // /drives omits some libraries (notably "Site Assets"), so also resolve them through /lists.
  const lists = await getJson<{ value: { id: string; displayName: string }[] }>(
    `/sites/${siteId}/lists?$select=id,displayName&$top=200`,
    token,
    `Listing lists of site ${siteId}`
  );

  for (const list of lists.value) {
    const drive = await resolveListDrive(siteId, list.id, list.displayName, token);
    if (drive && !drives.has(drive.id)) {
      drives.set(drive.id, drive);
    }
  }

  // Graph hides some system libraries from /drives and /lists; SiteAssets holds the notebooks.
  for (const knownList of ["SiteAssets", "Site Assets", "Shared Documents", "Documents"]) {
    const drive = await resolveListDrive(siteId, encodeURIComponent(knownList), knownList, token);
    if (drive && !drives.has(drive.id)) {
      drives.set(drive.id, drive);
    }
  }

  return [...drives.values()];
}

async function listChildren(driveId: string, itemPath: string, token: string): Promise<DriveItem[]> {
  const body = await getJson<{ value: DriveItem[] }>(
    `/drives/${driveId}/${itemPath}?$select=id,name,folder,file&$top=200`,
    token,
    `Listing ${itemPath} of drive ${driveId}`
  );
  return body.value;
}

function isNotebookFolder(children: DriveItem[]): boolean {
  return children.some((c) => c.name.toLowerCase().endsWith(NOTEBOOK_MARKER_EXTENSION));
}

const MAX_FOLDER_DEPTH = 3;

async function collectNotebooksInDrive(
  drive: DriveRef,
  token: string,
  inspected: string[]
): Promise<NotebookLocation[]> {
  const found: NotebookLocation[] = [];

  async function walk(folders: DriveItem[], depth: number): Promise<void> {
    for (const folder of folders) {
      const children = await listChildren(drive.id, `items/${folder.id}/children`, token);
      inspected.push(`${drive.name}/${folder.name}`);

      if (isNotebookFolder(children)) {
        found.push({ driveId: drive.id, driveName: drive.name, folderId: folder.id, folderName: folder.name });
        continue;
      }

      if (depth < MAX_FOLDER_DEPTH) {
        await walk(children.filter((c) => c.folder), depth + 1);
      }
    }
  }

  const rootFolders = (await listChildren(drive.id, "root/children", token)).filter((item) => item.folder);
  await walk(rootFolders, 1);
  return found;
}

/** Fallback when the notebook sits deeper than the folder walk reaches. */
async function searchNotebookInDrive(
  drive: DriveRef,
  notebookName: string,
  token: string
): Promise<NotebookLocation | undefined> {
  const response = await graphFetch(
    `/drives/${drive.id}/root/search(q='${encodeURIComponent(notebookName)}')?$select=id,name,folder`,
    token
  );
  if (!response.ok) {
    return undefined;
  }

  const body = (await response.json()) as { value: DriveItem[] };
  for (const item of body.value.filter((i) => i.folder)) {
    const children = await listChildren(drive.id, `items/${item.id}/children`, token);
    if (isNotebookFolder(children)) {
      return { driveId: drive.id, driveName: drive.name, folderId: item.id, folderName: item.name };
    }
  }
  return undefined;
}

function notebookNotFound(
  siteId: string,
  notebookName: string | undefined,
  drives: DriveRef[],
  inspected: string[]
): Error {
  const libraries = drives.map((d) => d.name).join(", ") || "none";
  const folders = inspected.join(", ") || "none";
  return new Error(
    `Notebook ${notebookName ? `"${notebookName}" ` : ""}not found in site ${siteId}. ` +
      `Libraries: ${libraries}. Folders inspected: ${folders}`
  );
}

async function listNotebookSections(notebook: NotebookLocation, token: string): Promise<string[]> {
  const children = await listChildren(notebook.driveId, `items/${notebook.folderId}/children`, token);
  return children
    .filter((c) => c.file && c.name.toLowerCase().endsWith(SECTION_EXTENSION))
    .map((c) => normalizeName(c.name));
}

/** Lists the notebooks and their sections so a wrong name is obvious in the error message. */
async function describeNotebooks(notebooks: NotebookLocation[], token: string): Promise<string> {
  const described = await Promise.all(
    notebooks.map(async (notebook) => {
      const sections = await listNotebookSections(notebook, token);
      return `${notebook.driveName}/${notebook.folderName} [${sections.join(", ") || "no sections"}]`;
    })
  );
  return described.join("; ") || "none";
}

/** Notebooks may live in any document library of the site, not only in "Site Assets". */
async function findNotebooks(
  siteId: string,
  notebookNames: string[],
  token: string
): Promise<NotebookLocation[]> {
  const drives = await listDrives(siteId, token);
  // "Site Assets" holds notebooks in most tenants, so check it first.
  const ordered = [...drives].sort((a, b) => Number(/site\s*assets/i.test(b.name)) - Number(/site\s*assets/i.test(a.name)));
  const inspected: string[] = [];

  const notebooks: NotebookLocation[] = [];
  for (const drive of ordered) {
    notebooks.push(...(await collectNotebooksInDrive(drive, token, inspected)));
  }

  // Without explicit names every notebook of the site is treated as a template.
  if (!notebookNames.length) {
    if (!notebooks.length) {
      throw notebookNotFound(siteId, undefined, drives, inspected);
    }
    return notebooks;
  }

  const matched: NotebookLocation[] = [];
  for (const name of notebookNames) {
    let match = notebooks.find((n) => normalizeName(n.folderName) === normalizeName(name));

    for (const drive of ordered) {
      if (match) {
        break;
      }
      match = await searchNotebookInDrive(drive, name, token);
    }

    if (!match) {
      // OneNote links expose section names, so a "notebook" name that is really a section
      // means the notebook filter was not intended.
      const sections = await Promise.all(notebooks.map((n) => listNotebookSections(n, token)));
      if (sections.some((names) => names.includes(normalizeName(name)))) {
        return notebooks;
      }

      throw new Error(
        `Notebook "${name}" not found in site ${siteId}. ` +
          `Notebooks found: ${await describeNotebooks(notebooks, token)}. ` +
          `Omit templateNotebookName to search every notebook of the site.`
      );
    }
    matched.push(match);
  }

  return matched;
}

/** Retries while the group's site or notebook is still being provisioned. */
async function waitForProvisioned<T>(what: string, resolve: () => Promise<T>): Promise<T> {
  let lastError: unknown;

  for (let attempt = 0; attempt < PROVISIONING_ATTEMPTS; attempt++) {
    try {
      return await resolve();
    } catch (err) {
      lastError = err;
      if (attempt < PROVISIONING_ATTEMPTS - 1) {
        await sleep(PROVISIONING_DELAY_MS);
      }
    }
  }

  throw new Error(
    `${what} is not available after ${PROVISIONING_ATTEMPTS} attempts: ${(lastError as Error).message}`
  );
}

/** Graph reports copy progress on an unauthenticated monitor URL. */
async function waitForCopyToFinish(monitorUrl: string): Promise<void> {
  for (let attempt = 0; attempt < 10; attempt++) {
    const response = await fetch(monitorUrl);
    if (!response.ok) {
      return;
    }

    const status = (await response.json()) as { status?: string; error?: unknown };
    if (status.status === "completed") {
      return;
    }
    if (status.status === "failed") {
      throw new Error(`Copy failed: ${JSON.stringify(status.error)}`);
    }
    await sleep(2000);
  }
}

async function copySectionFile(
  source: NotebookLocation,
  sectionItemId: string,
  target: NotebookLocation,
  newName: string,
  token: string
): Promise<void> {
  const fileName = newName.toLowerCase().endsWith(SECTION_EXTENSION) ? newName : `${newName}${SECTION_EXTENSION}`;

  async function requestCopy(): Promise<Response> {
    return graphFetch(`/drives/${source.driveId}/items/${sectionItemId}/copy`, token, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        parentReference: { driveId: target.driveId, id: target.folderId },
        name: fileName,
        // EasyLife pre-creates empty sections with the same name, so overwrite them.
        "@microsoft.graph.conflictBehavior": "replace",
      }),
    });
  }

  let response = await requestCopy();

  // Not every drive honours conflictBehavior on copy, so remove the placeholder and retry.
  if (response.status === 409) {
    const existing = (await listChildren(target.driveId, `items/${target.folderId}/children`, token)).find(
      (item) => item.name.toLowerCase() === fileName.toLowerCase()
    );
    if (existing) {
      await graphFetch(`/drives/${target.driveId}/items/${existing.id}`, token, { method: "DELETE" });
      response = await requestCopy();
    }
  }

  // Graph answers 202 Accepted and completes the copy asynchronously.
  if (!response.ok && response.status !== 202) {
    throw new Error(`Copying section "${newName}" failed: ${response.status} ${await response.text()}`);
  }

  const monitorUrl = response.headers.get("Location");
  if (monitorUrl) {
    await waitForCopyToFinish(monitorUrl);
  }
}

/** Resolves the section Teams shows by default, whatever it is called in this tenant. */
function findDefaultSectionName(targetFiles: DriveItem[]): string | undefined {
  return targetFiles
    .filter((f) => f.name.toLowerCase().endsWith(SECTION_EXTENSION))
    .find((f) => DEFAULT_SECTION_NAMES.includes(normalizeName(f.name)))?.name;
}

/** Teams creates its channel section after the webhook fires, so give it time to appear. */
async function resolveDefaultSectionName(target: NotebookLocation, token: string): Promise<string> {
  let files: DriveItem[] = [];

  for (let attempt = 0; attempt < PROVISIONING_ATTEMPTS; attempt++) {
    files = await listChildren(target.driveId, `items/${target.folderId}/children`, token);
    const defaultSection = findDefaultSectionName(files);
    if (defaultSection) {
      return defaultSection;
    }
    await sleep(PROVISIONING_DELAY_MS);
  }

  const fallback = files.find((f) => f.name.toLowerCase().endsWith(SECTION_EXTENSION))?.name;
  if (!fallback) {
    throw new Error("No section found to use as the default target section.");
  }
  return fallback;
}

/** Keeps the first notebook that provides a section name when several templates overlap. */
function dedupeSections(sections: TemplateSection[]): TemplateSection[] {
  const seen = new Set<string>();
  return sections.filter((section) => {
    const key = normalizeName(section.item.name);
    if (seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

async function collectTemplateSections(
  sources: TemplateSource[],
  token: string
): Promise<{ sections: TemplateSection[]; notebooks: string[] }> {
  const sections: TemplateSection[] = [];
  const notebooks: string[] = [];
  const failures: string[] = [];

  for (const source of sources) {
    try {
      const siteId = await resolveSiteId(source, token);
      for (const notebook of await findNotebooks(siteId, source.notebookNames, token)) {
        notebooks.push(`${notebook.driveName}/${notebook.folderName}`);
        const items = await listChildren(notebook.driveId, `items/${notebook.folderId}/children`, token);
        for (const item of items.filter((i) => i.file && i.name.toLowerCase().endsWith(SECTION_EXTENSION))) {
          sections.push({ notebook, item });
        }
      }
    } catch (err) {
      // One unreachable template must not break the remaining sources.
      failures.push((err as Error).message);
    }
  }

  if (!sections.length) {
    throw new Error(`No template sections found. ${failures.join(" | ") || "The templates contain no sections."}`);
  }

  return { sections, notebooks: [...new Set(notebooks)] };
}

/** Copies template sections (.one files) into the notebook of the newly provisioned group. */
export async function copyTemplateSectionsToGroup(options: CopyTemplateOptions): Promise<CopyTemplateResult> {
  const { token, sources, sections, targetGroupId } = options;

  const { sections: sourceSections, notebooks: templateNotebooks } = await collectTemplateSections(sources, token);

  const targetSiteId = await waitForProvisioned(`Site of group ${targetGroupId}`, () =>
    resolveSiteIdFromGroup(targetGroupId, token)
  );
  const targetNotebook = await waitForProvisioned(`Notebook of group ${targetGroupId}`, async () =>
    (await findNotebooks(targetSiteId, [], token))[0]
  );

  const mappings = sections.length
    ? sections
    : dedupeSections(sourceSections).map((s) => ({ from: s.item.name, to: s.item.name }));
  const copied: { from: string; to: string }[] = [];

  const needsDefaultSection = mappings.some((m) => normalizeName(m.to) === DEFAULT_SECTION_TOKEN);
  const defaultSectionName = needsDefaultSection
    ? await resolveDefaultSectionName(targetNotebook, token)
    : undefined;

  for (const mapping of mappings) {
    const sourceSection = sourceSections.find((s) => normalizeName(s.item.name) === normalizeName(mapping.from));
    if (!sourceSection) {
      const available =
        sourceSections.map((s) => `${s.notebook.folderName}/${normalizeName(s.item.name)}`).join(", ") || "none";
      throw new Error(`Template section "${mapping.from}" not found. Available sections: ${available}`);
    }

    const requestedTarget = normalizeName(mapping.to);
    const targetName =
      requestedTarget === DEFAULT_SECTION_TOKEN
        ? (defaultSectionName as string)
        : requestedTarget === SOURCE_SECTION_TOKEN
          ? sourceSection.item.name
          : mapping.to;

    await copySectionFile(sourceSection.notebook, sourceSection.item.id, targetNotebook, targetName, token);
    copied.push({ from: normalizeName(sourceSection.item.name), to: normalizeName(targetName) });
  }

  const filesInTargetNotebook = (
    await listChildren(targetNotebook.driveId, `items/${targetNotebook.folderId}/children`, token)
  ).map((item) => item.name);

  return {
    sectionsCopied: copied,
    templateNotebooks,
    targetNotebook: `${targetNotebook.driveName}/${targetNotebook.folderName}`,
    filesInTargetNotebook,
  };
}
