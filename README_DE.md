# EasyLife 365 OneNote Template Function

Diese Azure Function kopiert OneNote-Vorlagen automatisch in das neue OneNote-Notizbuch, das EasyLife 365 beim Erstellen einer Microsoft-365-Gruppe oder eines Teams anlegt.

Die Vorlage kann in einer SharePoint-Site oder in einer Microsoft-365-Gruppe liegen. Es können eine oder mehrere Sections inklusive ihrer Seiten kopiert werden. Die Zielgruppe muss nicht vorher bekannt sein: EasyLife sendet ihre ID nach der Bereitstellung an den Webhook.

**Autoren:** Andy Bui, uniQconsulting ag

## Was wird benötigt?

- Microsoft-365-Tenant mit einer OneNote-Vorlage
- Azure-Abonnement
- GitHub-Konto
- EasyLife 365 mit Berechtigung zum Erstellen einer Team-/Gruppen-Automation
- Azure Function App mit Node.js auf Flex Consumption oder einem anderen unterstützten Plan
- eine Microsoft-Entra-App-Registrierung für Microsoft Graph

## Schnellstart

1. Dieses Repository auf GitHub forken.
2. Im Fork unter **Settings → Secrets and variables → Actions** die drei Azure-Deployment-Secrets eintragen.
3. Im Azure Portal eine Function App mit GitHub Deployment Center verbinden oder den vorhandenen Workflow verwenden.
4. Die Microsoft-Graph-App-Registrierung und Function-App-Einstellungen konfigurieren.
5. Einen Function Key erzeugen.
6. In EasyLife einen aktiven OneNote-Automation-Step mit Webhook anlegen.
7. Eine Testgruppe provisionieren und das neue OneNote prüfen.

## 1. Repository forken

Öffne dieses Repository auf GitHub und wähle **Fork**. Verwende anschliessend deinen eigenen Fork als Quelle für das Deployment. Die weiteren Schritte beziehen sich auf den Fork, nicht auf das Original-Repository.

Lokal ist kein Klonen erforderlich. Falls du lokal arbeiten möchtest:

```powershell
git clone https://github.com/<organisation-oder-benutzer>/<dein-fork>.git
cd easylife-onenote-extension
npm install
npm run build
```

## 2. Azure Function App erstellen

Erstelle im Azure Portal eine Function App mit diesen Eckdaten:

- Runtime stack: **Node.js**
- Node-Version: **22 LTS** oder eine vom Repository unterstützte LTS-Version
- Region: nach deiner Umgebung
- Hosting: **Flex Consumption** ist möglich und benötigt OIDC-Deployment
- Betriebssystem: Linux

Danach öffnest du **Deployment Center** und verbindest die Function App mit deinem GitHub-Fork:

1. Source: **GitHub**
2. Organisation oder Benutzer auswählen
3. Repository auswählen
4. Branch: `main`
5. Build provider: **GitHub Actions**
6. Für Flex Consumption: **User-assigned identity** beziehungsweise OIDC verwenden

Azure erzeugt dabei normalerweise einen Workflow unter `.github/workflows/`. Dieser muss `azure/login@v2` mit OIDC verwenden. Der Workflow braucht die Berechtigung:

```yaml
permissions:
  id-token: write
  contents: read
```

Ein Publish Profile ist für Flex Consumption nicht geeignet, weil dort Kudu und ZipDeploy nicht verfügbar sind.

### GitHub Deployment Secrets

Der Workflow verwendet diese drei Secrets. Azure Deployment Center kann sie beim Verbinden automatisch anlegen. Falls du den Workflow manuell einrichtest, müssen sie im Fork unter **Settings → Secrets and variables → Actions** exakt so vorhanden sein:

| Secret | Inhalt |
|---|---|
| `AZURE_CLIENT_ID` | Client-ID der Deployment-Identität |
| `AZURE_TENANT_ID` | ID des Microsoft-Entra-Tenants |
| `AZURE_SUBSCRIPTION_ID` | ID des Azure-Abonnements |

Der Workflow deployt auf die Function App, die in der Repository-Variable `AZURE_FUNCTIONAPP_NAME` steht. Diese unter **Settings → Secrets and variables → Actions → Variables** im Fork anlegen, sonst zeigt der Workflow auf die Function App des Original-Repositories und das Deployment schlägt fehl.

| Variable | Wert |
|---|---|
| `AZURE_FUNCTIONAPP_NAME` | Name der eigenen Function App |

Wird die Function App stattdessen über das Azure Deployment Center verbunden, legt Azure einen eigenen Workflow mit eigenen Secret-Namen an. In dem Fall den hier mitgelieferten Workflow löschen oder deaktivieren, damit nicht beide laufen.

Die Deployment-Identität benötigt auf der Function App mindestens **Contributor** oder **Website Contributor**. Zusätzlich muss eine Federated Credential für deinen Fork und den Branch `main` existieren.

## 3. Microsoft Graph konfigurieren

Die Function benötigt eine separate Microsoft-Entra-App-Registrierung für Microsoft Graph. Die bestehende EasyLife-App-Registrierung wird für den Webhook mit Function Key nicht benötigt.

In **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions** hinzufügen:

- `Sites.ReadWrite.All`
- `Group.Read.All`

Danach unbedingt **Grant admin consent** für den Tenant ausführen.

> **Warum nicht die OneNote-API?** Microsoft Graph lehnt App-only-Token für die OneNote-API ab (Fehler `40001`, seit 31.03.2025 erzwungen). Dieses Projekt kopiert OneNote-Sections deshalb als zugrunde liegende `.one`-Dateien über die SharePoint-Drive-API, die App-only weiterhin unterstützt. `Notes.ReadWrite.All` wird nicht benötigt.

Unter **Certificates & secrets** ein Client Secret erstellen. Den Secret-Wert direkt kopieren; er wird später nicht erneut vollständig angezeigt.

## 4. Function App konfigurieren

In Azure: **Function App → Settings → Environment variables**. Diese Werte als Application Settings eintragen:

| Einstellung | Wert |
|---|---|
| `GRAPH_TENANT_ID` | Tenant-ID des Microsoft-365-Tenants |
| `GRAPH_CLIENT_ID` | Client-ID der Graph-App-Registrierung |
| `GRAPH_CLIENT_SECRET` | Client-Secret der Graph-App-Registrierung |

### Vorlage aus einer SharePoint-Site

Für die Vorlage aus `https://<tenant>.sharepoint.com/sites/<site>`:

| Einstellung | Wert |
|---|---|
| `DEFAULT_TEMPLATE_SITE_URL` | Vollständige SharePoint-Site-URL |
| `DEFAULT_TEMPLATE_NOTEBOOK_NAME` | Name des OneNote-Notizbuchs, z. B. `Vorlage` |
| `DEFAULT_TEMPLATE_SECTION_NAMES` | Optional, kommagetrennt; leer bedeutet alle Sections |
| `DEFAULT_TARGET_SECTION_NAMES` | Optional, kommagetrennt; leer verwendet die Originalnamen |

Beispiel:

```text
DEFAULT_TEMPLATE_SITE_URL=https://uniqconsultingch.sharepoint.com/sites/T-int-M365
DEFAULT_TEMPLATE_NOTEBOOK_NAME=Vorlage
DEFAULT_TEMPLATE_SECTION_NAMES=Besprechungen,Dokumentation
DEFAULT_TARGET_SECTION_NAMES=Besprechungen,Dokumentation
```

### Vorlage aus einer Microsoft-365-Gruppe

Alternativ kann eine Vorlagen-Gruppe verwendet werden:

| Einstellung | Wert |
|---|---|
| `DEFAULT_TEMPLATE_GROUP_ID` | Gruppen-ID der Vorlagen-Gruppe |
| `DEFAULT_TEMPLATE_NOTEBOOK_NAME` | Optionaler Name des Vorlage-Notizbuchs |
| `DEFAULT_TEMPLATE_SECTION_NAMES` | Optional, kommagetrennt |
| `DEFAULT_TARGET_SECTION_NAMES` | Optional, kommagetrennt |

Die Vorlagen-Einstellungen dürfen leer bleiben, wenn sie vollständig in der Webhook-URL übergeben werden.

## 5. Webhook in EasyLife einrichten

In EasyLife 365 einen Automation Step vom Typ **Team & Group** öffnen oder erstellen und zu **OneNote** wechseln:

1. **OneNote** aktivieren.
2. Unter **Sections** mindestens eine Section definieren, zum Beispiel `Besprechungen`. Diese Section wird von EasyLife im neuen Notizbuch angelegt.
3. Unter **Naming** die Standard-Notebook-Bereitstellung aktiviert lassen.
4. Unter **Webhook** **Notify via webhook** aktivieren.
5. Bei **Authentication** **Code authentication** wählen.
6. Den Function Key bei **Authentication code** eintragen.
7. Den Automation Step mit **Activate and save** aktivieren und speichern.

Die Webhook-URL lautet:

```text
https://<function-app-name>.azurewebsites.net/api/onenote-template
```

Falls deine EasyLife-Version den Code nicht als Header `x-functions-key` sendet, akzeptiert die Function auch einen `code`-Query-Parameter. Bevorzugt wird das Feld **Authentication code**, weil Keys in URLs in Logs und Konfigurationsansichten auftauchen können.

Den Function Key findest du unter **Function App → Functions → provisionOneNoteTemplate → Function keys**.

### Vorlage pro Automation auswählen

Query-Parameter in der Webhook-URL überschreiben die Application Settings. Für ein SharePoint-Notebook `Vorlage`:

```text
https://<function-app-name>.azurewebsites.net/api/onenote-template?templateSiteUrl=https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FT-int-M365&templateNotebookName=Vorlage&templateSectionName=Besprechungen&targetSectionName=Besprechungen
```

Für alle Sections das Feld `templateSectionName` weglassen:

```text
https://<function-app-name>.azurewebsites.net/api/onenote-template?templateSiteUrl=https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FT-int-M365&templateNotebookName=Vorlage
```

Für mehrere Sections:

```text
...&templateSectionName=Besprechungen,Dokumentation,Traktandenliste
```

### Mehrere Vorlagen kombinieren

`templateNotebookName`, `templateSiteUrl` und `templateGroupId` akzeptieren kommagetrennte Listen. Mehrere Notizbücher derselben Site:

```text
...&templateNotebookName=Vorlage,Status%20meeting
```

Mehrere Sites:

```text
...&templateSiteUrl=https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FA,https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FB
```

Ohne `templateNotebookName` werden alle Notizbücher der angegebenen Sites als Vorlage verwendet. Mit `templateSectionName` lässt sich daraus gezielt eine Section übernehmen, ohne das Notizbuch zu kennen:

```text
...&templateSiteUrl=https%3A%2F%2Fcontoso.sharepoint.com%2Fsites%2FT-int-M365&templateSectionName=Meeting%20Notes&targetSectionName=%40default
```

Kommt eine Section in mehreren Notizbüchern vor, gewinnt der erste Treffer. Ist eine Site oder ein Notizbuch nicht erreichbar, werden die übrigen Quellen trotzdem verarbeitet.

Jeder Parameter akzeptiert Einzel- und Mehrzahlschreibweise sowie mehrfach wiederholte Query-Parameter. `templateNotebook`, `templateNotebooks`, `templateNotebookName` und `templateNotebookNames` sind gleichwertig; dasselbe gilt für `templateSite`, `templateGroup`, `templateSection` und `targetSection`. Diese Varianten sind ebenfalls gültig:

```text
...&templateNotebooks=Vorlage,Status%20meeting
...&templateNotebook=Vorlage&templateNotebook=Status%20meeting
```

Wird gar kein Parameter gesetzt, werden alle Sections aller gefundenen Vorlage-Notizbücher 1:1 mit ihren Originalnamen kopiert.

Die Teams-Registerkarte eines Kanals ist fest an den Standardabschnitt des Notizbuchs gebunden, dessen Name je nach Tenant-Sprache variiert (`General`, `Allgemein`, ...). Mit `@default` wird genau dieser Abschnitt überschrieben, damit die Vorlage dort erscheint, wo Anwender zuerst hinschauen:

```text
...&templateSectionName=Vorlage&targetSectionName=@default
```

Das Ziel wird nicht in der URL konfiguriert. EasyLife sendet die neue Gruppen-ID nach der Bereitstellung im Payload, normalerweise unter `group.id`.

## 6. Testen

1. Prüfe, dass das Vorlage-Notizbuch existiert und die Vorlage-Sections mindestens eine Seite enthalten.
2. Prüfe, dass die Graph-App Admin Consent erhalten hat.
3. Prüfe, dass die Function App deployed und die Application Settings gespeichert sind.
4. Speichere den EasyLife-Step mit **Activate and save**.
5. Provisioniere eine Testgruppe.
6. Azure Portal: **Function App → Functions → provisionOneNoteTemplate → Invocations** öffnen.
7. Prüfe den HTTP-Ergebniscode und die kopierten Seiten im Ziel-OneNote.

Ein erfolgreicher Aufruf sieht ungefähr so aus:

```json
{
  "status": "ok",
  "sectionsCopied": [{ "from": "vorlage", "to": "allgemein" }],
  "templateNotebooks": ["Site Assets/T-int-M365 Notebook"],
  "targetNotebook": "Site Assets/Projekt Contoso Notebook",
  "filesInTargetNotebook": ["Allgemein.one", "Projekt Contoso Notebook.onetoc2"]
}
```

## Fehlerbehebung

| Ergebnis | Ursache oder Lösung |
|---|---|
| `401` | Function Key fehlt oder ist falsch. In EasyLife **Code authentication** und den aktuellen Function Key verwenden. |
| `400` | EasyLife-Payload enthält keine Gruppen-ID oder es fehlt `templateSiteUrl` beziehungsweise `templateGroupId`. |
| `403` | Der Graph-App fehlt `Sites.ReadWrite.All` oder der Admin Consent. |
| `Notebook "..." not found` | Name des **Notizbuchs** verwenden, nicht den des Abschnitts. Die Fehlermeldung listet alle geprüften Bibliotheken und Ordner auf. |
| Section nicht gefunden | Exakten Abschnittsnamen aus dem Vorlage-Notizbuch verwenden. Ohne Angabe werden alle Abschnitte kopiert. |
| Abschnitte erscheinen nicht in OneNote | Siehe unten. |

### Abschnitte erscheinen nicht in OneNote

Die Function kopiert Abschnittsdateien in den Notizbuch-Ordner. OneNote nimmt sie erst in sein Inhaltsverzeichnis (`.onetoc2`) auf, wenn das Notizbuch geöffnet und synchronisiert wird. Neu kopierte Abschnitte können daher verzögert erscheinen.

Prüfe zuerst das Aufrufergebnis. `filesInTargetNotebook` zeigt, was physisch im Zielordner liegt:

```json
{ "filesInTargetNotebook": ["Open Notebook.onetoc2", "Vorlage.one"] }
```

Sind die `.one`-Dateien aufgeführt, hat das Kopieren funktioniert und es liegt an der OneNote-Indizierung:

1. Notizbuch in Teams schliessen und erneut öffnen.
2. Notizbuch einmal in der OneNote-Desktop-App öffnen; das erzwingt einen erneuten Scan des Ordners.
3. In SharePoint unter **Site Assets → &lt;Gruppe&gt; Notebook** prüfen, ob die Dateien vorhanden sind.

EasyLife wiederholt fehlgeschlagene Webhooks mehrfach; identische Fehlermeldungen kurz nacheinander sind daher normal.

## Sicherheit

- Niemals Client Secrets, Function Keys oder GitHub Secrets in Git, README-Dateien oder Screenshots speichern.
- `local.settings.json` ist lokal und wird durch `.gitignore` nicht versioniert.
- Für produktive Umgebungen Secrets regelmässig erneuern oder Zertifikate beziehungsweise Managed Identity verwenden.
- Die Function kopiert nur in die neu von EasyLife gemeldete Zielgruppe; diese ID wird nicht vorher konfiguriert.

## Entwicklung

```powershell
npm install
npm run build
```

Für die lokale Ausführung mit Azure Functions Core Tools:

```powershell
func start
```

Die HTTP-Function ist in [src/functions/provisionOneNoteTemplate.ts](src/functions/provisionOneNoteTemplate.ts) definiert. Die Graph- und OneNote-Logik befindet sich in [src/services](src/services).

## Lizenz und Eigentum

Dieses Projekt wurde für uniQconsulting ag erstellt. Lizenzierung und Weiterverwendung sind mit uniQconsulting ag zu klären.
