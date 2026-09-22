# EasyLife 365 OneNote Template Function

[Deutsche Anleitung](README_DE.md)

This Azure Function copies OneNote templates into the new OneNote notebook created by EasyLife 365 when provisioning a Microsoft 365 group or team.

Templates can be stored in a SharePoint site or a Microsoft 365 group. The function copies one or more sections, including their pages. The target group does not need to be known in advance: EasyLife sends its ID to the webhook after provisioning.

**Authors:** Andy Bui, uniQconsulting ag

## Requirements

- Microsoft 365 tenant with an existing OneNote template
- Azure subscription
- GitHub account
- EasyLife 365 permission to create a Team & Group automation
- Azure Function App running Node.js on Flex Consumption or another supported plan
- Microsoft Entra app registration for Microsoft Graph

## Quick Start

1. Fork this repository on GitHub.
2. Add the three Azure deployment secrets under **Settings → Secrets and variables → Actions** in your fork.
3. Connect the Function App to your fork through Azure Deployment Center, or use the workflow already included in this repository.
4. Configure the Microsoft Graph app registration and Function App settings.
5. Create a Function key.
6. Create and activate an EasyLife OneNote automation step with a webhook.
7. Provision a test group and verify the new OneNote notebook.

## 1. Fork the Repository

Open this repository on GitHub and select **Fork**. Use your fork as the deployment source. The steps below apply to your fork, not the upstream repository.

Cloning locally is optional. If you want to develop locally:

```powershell
git clone https://github.com/<organisation-or-user>/<your-fork>.git
cd easylife-onenote-extension
npm install
npm run build
```

## 2. Create the Azure Function App

Create a Function App in the Azure portal with these settings:

- Runtime stack: **Node.js**
- Node version: **22 LTS** or another LTS version supported by the repository
- Region: choose the region required by your organisation
- Hosting: **Flex Consumption** is supported and requires OIDC deployment
- Operating system: Linux

Open **Deployment Center** and connect the Function App to your GitHub fork:

1. Source: **GitHub**
2. Select the organisation or user
3. Select the repository
4. Branch: `main`
5. Build provider: **GitHub Actions**
6. For Flex Consumption, select **User-assigned identity** or OIDC authentication

Azure normally creates a workflow under `.github/workflows/`. It must use `azure/login@v2` with OIDC and include:

```yaml
permissions:
  id-token: write
  contents: read
```

Publish Profiles are not suitable for Flex Consumption because Kudu and ZipDeploy are not available there.

### GitHub Deployment Secrets

Deployment Center can create these secrets automatically. If you configure the workflow manually, add the following secrets to your fork under **Settings → Secrets and variables → Actions**:

| Secret | Value |
|---|---|
| `AZURE_CLIENT_ID` | Client ID of the deployment identity |
| `AZURE_TENANT_ID` | Microsoft Entra tenant ID |
| `AZURE_SUBSCRIPTION_ID` | Azure subscription ID |

The deployment identity needs at least **Contributor** or **Website Contributor** on the Function App. A federated credential for your fork and the `main` branch is also required.

## 3. Configure Microsoft Graph

The Function App uses a separate Microsoft Entra app registration for Microsoft Graph. The existing EasyLife app registration is not required for the Function Key webhook configuration.

In **API permissions → Add a permission → Microsoft Graph → Application permissions**, add:

- `Sites.ReadWrite.All`
- `Group.Read.All`

Then select **Grant admin consent** for the tenant.

> **Why not the OneNote API?** Microsoft Graph rejects app-only tokens for the OneNote API (error `40001`, enforced since 31 March 2025). This project therefore copies OneNote sections as their underlying `.one` files through the SharePoint Drive API, which still supports app-only access. `Notes.ReadWrite.All` is not required.

Under **Certificates & secrets**, create a client secret. Copy the secret value immediately; it cannot be fully displayed again later.

## 4. Configure the Function App

In Azure, open **Function App → Settings → Environment variables** and add these Application Settings:

| Setting | Value |
|---|---|
| `GRAPH_TENANT_ID` | Microsoft 365 tenant ID |
| `GRAPH_CLIENT_ID` | Client ID of the Graph app registration |
| `GRAPH_CLIENT_SECRET` | Client secret of the Graph app registration |

### SharePoint Site Template

For a template stored in a site such as `https://<tenant>.sharepoint.com/sites/<site>`:

| Setting | Value |
|---|---|
| `DEFAULT_TEMPLATE_SITE_URL` | Full SharePoint site URL |
| `DEFAULT_TEMPLATE_NOTEBOOK_NAME` | OneNote notebook name, for example `Template` |
| `DEFAULT_TEMPLATE_SECTION_NAMES` | Optional comma-separated section names; empty means all sections |
| `DEFAULT_TARGET_SECTION_NAMES` | Optional comma-separated target names; empty keeps the source names |

Example:

```text
DEFAULT_TEMPLATE_SITE_URL=https://contoso.sharepoint.com/sites/OneNoteTemplates
DEFAULT_TEMPLATE_NOTEBOOK_NAME=Template
DEFAULT_TEMPLATE_SECTION_NAMES=Meetings,Documentation
DEFAULT_TARGET_SECTION_NAMES=Meetings,Documentation
```

### Microsoft 365 Group Template

Alternatively, use a template group:

| Setting | Value |
|---|---|
| `DEFAULT_TEMPLATE_GROUP_ID` | Group ID of the template group |
| `DEFAULT_TEMPLATE_NOTEBOOK_NAME` | Optional template notebook name |
| `DEFAULT_TEMPLATE_SECTION_NAMES` | Optional comma-separated section names |
| `DEFAULT_TARGET_SECTION_NAMES` | Optional comma-separated target names |

Template settings can remain empty when all values are supplied in the webhook URL.

## 5. Configure the EasyLife Webhook

Open or create a **Team & Group** automation step in EasyLife 365 and select **OneNote**:

1. Enable **OneNote**.
2. Under **Sections**, define at least one section, for example `Meetings`. EasyLife creates this section in the new notebook.
3. Keep **Provision default notebook** enabled under **Naming**.
4. Enable **Notify via webhook**.
5. Select **Code authentication**.
6. Enter the Function key in **Authentication code**.
7. Select **Activate and save** to activate the automation step.

Use this as the webhook URL:

```text
https://<function-app-name>.azurewebsites.net/api/onenote-template
```

If your EasyLife version does not send the authentication code as the `x-functions-key` header, the Function also accepts a `code` query parameter. Prefer the **Authentication code** field, because keys in URLs can end up in logs and configuration views.

Find the Function key under **Function App → Functions → provisionOneNoteTemplate → Function keys**.

### Select a Template Per Automation

Query parameters override the Function App settings. For a SharePoint notebook named `Template`:

```text
https://<function-app-name>.azurewebsites.net/api/onenote-template?templateSiteUrl=https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FOneNoteTemplates&templateNotebookName=Template&templateSectionName=Meetings&targetSectionName=Meetings
```

To copy all sections, omit `templateSectionName`:

```text
https://<function-app-name>.azurewebsites.net/api/onenote-template?templateSiteUrl=https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FOneNoteTemplates&templateNotebookName=Template
```

To copy multiple sections:

```text
...&templateSectionName=Meetings,Documentation,Agenda
```

### Combine Several Templates

`templateNotebookName`, `templateSiteUrl`, and `templateGroupId` accept comma-separated lists. Several notebooks of the same site:

```text
...&templateNotebookName=Template,Status%20meeting
```

Several sites:

```text
...&templateSiteUrl=https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FA,https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FB
```

Without `templateNotebookName`, every notebook of the given sites is treated as a template. Combined with `templateSectionName` this picks a section without knowing its notebook:

```text
...&templateSiteUrl=https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FOneNoteTemplates&templateSectionName=Meeting%20Notes&targetSectionName=%40default
```

If a section name exists in several notebooks, the first match wins. If one site or notebook is unreachable, the remaining sources are still processed.

Every parameter accepts singular and plural spellings as well as repeated query parameters. `templateNotebook`, `templateNotebooks`, `templateNotebookName`, and `templateNotebookNames` are equivalent; the same applies to `templateSite`, `templateGroup`, `templateSection`, and `targetSection`. These variants are valid too:

```text
...&templateNotebooks=Template,Status%20meeting
...&templateNotebook=Template&templateNotebook=Status%20meeting
```

Without any parameter, every section of every template notebook found is copied 1:1 under its original name.

The Teams channel tab is pinned to the notebook's default section, which is named differently per tenant language (`General`, `Allgemein`, ...). Use `@default` to overwrite exactly that section, so the template appears where users look first:

```text
...&templateSectionName=Meetings&targetSectionName=@default
```

The target group is not configured in the URL. EasyLife sends the new group ID after provisioning, normally under `group.id`.

## 6. Test the Installation

1. Verify that the template notebook exists and its sections contain at least one page.
2. Verify that the Graph app has received admin consent.
3. Verify that the Function App is deployed and its Application Settings are saved.
4. Save the EasyLife step with **Activate and save**.
5. Provision a test group.
6. In Azure, open **Function App → Functions → provisionOneNoteTemplate → Invocations**.
7. Check the HTTP result and the copied pages in the target OneNote notebook.

A successful response looks like this:

```json
{
  "status": "ok",
  "sectionsCopied": [{ "from": "meetings", "to": "general" }],
  "templateNotebooks": ["Site Assets/Template Notebook"],
  "targetNotebook": "Site Assets/Contoso Project Notebook",
  "filesInTargetNotebook": ["General.one", "Contoso Project Notebook.onetoc2"]
}
```

## Troubleshooting

| Result | Cause or solution |
|---|---|
| `401` | The Function key is missing or incorrect. Select **Code authentication** in EasyLife and use the current key. |
| `400` | The EasyLife payload has no group ID, or `templateSiteUrl` / `templateGroupId` is missing. |
| `403` | The Graph app is missing `Sites.ReadWrite.All` or admin consent. |
| `Notebook "..." not found` | Use the **notebook** name, not the section name. The error lists every library and folder that was inspected. |
| Section not found | Use the exact section name from the template notebook. If no section is specified, all sections are copied. |
| Sections do not appear in OneNote | See below. |

### Sections do not appear in OneNote

The function copies section files into the notebook folder. OneNote only adds them to its table of contents (`.onetoc2`) when the notebook is opened and synchronised, so newly copied sections can take a moment to show up.

Check the invocation result first. `filesInTargetNotebook` lists what is physically present in the target folder:

```json
{ "filesInTargetNotebook": ["Open Notebook.onetoc2", "Vorlage.one"] }
```

If the `.one` files are listed, the copy worked and the issue is OneNote indexing:

1. Close the notebook in Teams and open it again.
2. Open the notebook once in the OneNote desktop app, which forces a folder rescan.
3. In SharePoint, open **Site Assets → &lt;group&gt; Notebook** to confirm the files are there.

EasyLife may retry a failed webhook several times, so repeated identical errors are expected.

## Security

- Never commit client secrets, Function keys, or GitHub secrets to Git, README files, or screenshots.
- `local.settings.json` is local only and is excluded by `.gitignore`.
- Rotate production secrets regularly, or use certificates and managed identities where appropriate.
- The Function only writes to the newly provisioned group received from EasyLife; the target group ID is not configured in advance.

## Development

```powershell
npm install
npm run build
```

For local execution with Azure Functions Core Tools:

```powershell
func start
```

The HTTP Function is defined in [src/functions/provisionOneNoteTemplate.ts](src/functions/provisionOneNoteTemplate.ts). Microsoft Graph and OneNote logic is in [src/services](src/services).

## License and Ownership

This project was created for uniQconsulting ag. Licensing and reuse should be agreed with uniQconsulting ag.
