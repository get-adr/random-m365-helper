# Microsoft 365 Brand Center – Statusbericht

Zwei gleichwertige Scripte, die den Zustand des SharePoint Brand Centers eines Tenants auslesen und als Konsolentext oder als eigenständige HTML-Datei ausgeben. Die Scripte lesen nur und ändern nichts im Tenant.

| Script | Basis |
|---|---|
| `m365-brand-center-info.sh` | Bash, [CLI for Microsoft 365](https://pnp.github.io/cli-microsoft365/) (`m365`) und `jq` |
| `m365-brand-center-info.ps1` | PowerShell 7, [PnP.PowerShell](https://pnp.github.io/powershell/) |

## Inhalt des Berichts

- Tenant-Grunddaten
- Brand Center: Site-URL, Feature-Status, Markenfarben, Schriften für SharePoint / Viva Connections, Themes
- Office-Schriften (OfficeFontLibrary): Dateien mit Familie, Größe und Änderungsdatum
- Office-Vorlagen (OfficeTemplateLibrary)
- Bilder (ImageDocumentLibrary): Copilot-Suche, Dateianzahl
- CDN Public und Private: Status, Origins, Policies

Jede Org Asset Library erscheint mit ihrem OAL-Eintrag (Typ, URL, List-ID, CDN-Typ). Es werden nur Angaben angezeigt, die SharePoint oder der Font-Katalog selbst liefern. Abrufe, die fehlschlagen, werden als Fehler gemeldet und nicht als „leer“ dargestellt. Der Bericht nennt außerdem, mit welchem Werkzeug und in welcher Version er erstellt wurde.

## Voraussetzungen

- Bash-Variante: `m365` (angemeldet mit `m365 login`) und `jq`
- PowerShell-Variante: PowerShell 7 und das Modul `PnP.PowerShell` (`Install-Module PnP.PowerShell -Scope CurrentUser`)
- Leserechte auf Tenant und Brand-Center-Site; die CDN-Abfragen setzen in der Regel SharePoint-Administratorrechte voraus

## Anmeldung

- **Bash:** Das Script stellt beim Start mit `m365 login --ensure --authType browser` sicher, dass eine Sitzung besteht: Ist sie gültig, passiert nichts, sonst öffnet sich der Browser zur Anmeldung. Scheitert die Anmeldung, endet das Script mit Exit-Code 2. Wer ein anderes Verfahren (Zertifikat, Secret, Identität) nutzt, meldet sich vorher selbst an, zum Beispiel `m365 login --authType certificate …`. App-ID und Tenant lassen sich mit `--appId` und `--tenant` oder über die Umgebungsvariablen `CLIMICROSOFT365_ENTRAAPPID` und `CLIMICROSOFT365_TENANT` festlegen.
- **PowerShell:** Das Script meldet selbst im Browser an (`-Url`, optional `-ClientId`). Der Gerätecode wird nicht unterstützt, weil er in vielen Tenants gesperrt ist. Seit September 2024 verlangt PnP.PowerShell eine eigene Entra-ID-App-ID; alternativ lässt sie sich einmalig mit `Set-PnPManagedAppId` hinterlegen.

## Aufruf

```bash
# Bash: Text in der Konsole
./m365-brand-center-info.sh

# Bash: HTML-Bericht (mit -o implizit HTML)
./m365-brand-center-info.sh -f html -o bericht.html
```

```powershell
# PowerShell: Text in der Konsole
pwsh ./m365-brand-center-info.ps1 -Url https://<tenant>.sharepoint.com

# PowerShell: HTML-Bericht
pwsh ./m365-brand-center-info.ps1 -Url https://<tenant>.sharepoint.com -Format html -Output bericht.html
```

Ohne Dateiname entsteht `brand-center-report-<Datum>_<Uhrzeit>.html` im Scriptverzeichnis.

| Option | Bash | PowerShell |
|---|---|---|
| Format `text` oder `html` | `-f`, `--format` | `-Format` |
| Zieldatei (impliziert HTML) | `-o`, `--output` | `-Output` |
| Tenant-URL | ergibt sich aus der m365-Anmeldung | `-Url` |
| Anmeldeverfahren | Browser, sonst vorab `m365 login --authType …` | Browser |
| Eigene App-Registrierung (Entra-ID-App) | beim Login: `m365 login --appId <ID>` oder `CLIMICROSOFT365_ENTRAAPPID` | `-ClientId` |
| Getrennte Anmeldung für die Brand-Center-Site | nicht nötig, die Anmeldung gilt für alle Sites des Tenants | `-SeparateLogin` (sonst wird das Token der ersten Anmeldung wiederverwendet) |
| Abrufe nacheinander statt parallel | `M365_SEQUENTIAL=1` | nicht nötig, die Abrufe laufen immer nacheinander |
| Farben abschalten | `NO_COLOR=1` | `NO_COLOR=1` |

## Exit-Codes

| Code | Bedeutung |
|---|---|
| 0 | vollständig |
| 1 | mindestens ein Abruf ist fehlgeschlagen |
| 2 | Aufruf- oder Anmeldefehler |

## Hinweise

- Die Themes stammen aus dem Endpunkt `_api/brandcenter/GetTenantThemes` der Brand-Center-Site. Er ist nicht öffentlich dokumentiert und kann sich ohne Ankündigung ändern.
