#!/usr/bin/env pwsh
# ============================================================================
# m365-brand-center-info.ps1 – SharePoint Online Brand-Center-Bericht
# ============================================================================
# Zeigt mit PnP.PowerShell:
#   • Tenant-Grunddaten
#   • Brand Center (Site-URL, Feature-Status, Markenfarben, Schriften für SharePoint/Viva, Themes)
#   • Office-Schriften (OfficeFontLibrary) und Office-Vorlagen (OfficeTemplateLibrary)
#   • Brand Kits für Clipchamp (BrandKitLibrary)
#   • OAL (Org Asset Library): Bilder-Libraries, Asset-Typ, CDN-Typ, Dateianzahl
#   • CDN Config (Public & Private): Status, Origins, Policies
#
# Ablauf: anmelden -> alle Daten einmal abrufen -> auswerten -> Text oder HTML.
#
# Voraussetzung: PnP.PowerShell (SharePoint-Admin-Rechte).
#   Install:  Install-Module -Name PnP.PowerShell -Scope CurrentUser
#   Anmelden: erfolgt durch das Script (-Url, optional -ClientId); eine
#             Entra-ID-App-ID ist seit September 2024 zwingend (alternativ per
#             Set-PnPManagedAppId hinterlegen).
#
# Die Brand-Center-Site nutzt das Token der ersten Anmeldung; -SeparateLogin erzwingt eine eigene Anmeldung.
# Exit-Codes: 0 = vollständig, 1 = mindestens ein Abruf fehlgeschlagen,
#             2 = Aufruf-/Anmeldefehler.
# Umgebung:   NO_COLOR=1 schaltet Farben ab.
# ============================================================================

#Requires -Version 7.0

# Write-Host ist hier gewollt: farbige Konsolenausgabe eines Berichts, kein Pipeline-Ergebnis
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Farbige Konsolenausgabe')]
[CmdletBinding()]
param(
  [ValidateSet("text", "html")]
  [string]$Format,
  [string]$Output = "",
  [string]$Url = "",
  [string]$ClientId = "",
  [switch]$SeparateLogin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# --- Farben (nur bei Terminal und ohne NO_COLOR) --------------------------------

$esc = [char]27
if ((-not [Console]::IsOutputRedirected) -and (-not $env:NO_COLOR)) {
  $RED = "$esc[0;31m"; $GREEN = "$esc[0;32m"; $YELLOW = "$esc[1;33m"
  $BLUE = "$esc[0;34m"; $CYAN = "$esc[0;36m"; $BOLD = "$esc[1m"
  $DIM = "$esc[2m"; $RESET = "$esc[0m"
} else {
  $RED = ""; $GREEN = ""; $YELLOW = ""; $BLUE = ""; $CYAN = ""; $BOLD = ""; $DIM = ""; $RESET = ""
}

function Exit-Script {
  param([string]$Message, [int]$Code = 2)
  [Console]::Error.WriteLine("${RED}FEHLER:${RESET} $Message")
  exit $Code
}

# --- Ausgabeformat & Zieldatei ---------------------------------------------------

if (-not $PSBoundParameters.ContainsKey('Format')) {
  $Format = if ($Output) { "html" } else { "text" }
}
if (($Format -eq "text") -and $Output) {
  Exit-Script "-Output wird nur mit -Format html unterstützt."
}
if ($Format -eq "html" -and -not $Output) {
  $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
  $Output = Join-Path $scriptDir ("brand-center-report-{0}.html" -f (Get-Date -Format "yyyy-MM-dd_HH-mm-ss"))
}

# --- Hilfsfunktionen ---------------------------------------------------------------

# Safe-Eigenschaftszugriff (StrictMode-immun, Hashtable + PSObject);
# das Komma verhindert das Aufrollen von Arrays bei der Rückgabe.
function Get-PropValue {
  param($Object, $Name)
  if ($null -eq $Object) { return $null }
  if ($Object -is [System.Collections.IDictionary]) {
    # Schlüssel einzeln vergleichen: generische Dictionaries (z. B. mit Enum-Schlüsseln) haben kein passendes Contains(string)
    foreach ($k in $Object.Keys) { if ("$k" -eq "$Name") { return ,$Object[$k] } }
    return $null
  }
  $p = $Object.PSObject.Properties | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
  if ($null -ne $p) { return ,$p.Value }
  return $null
}

# Skalarzugriff: erster Name mit nicht-leerem Wert (verschluckt $false; für Booleans Get-PropValue nutzen)
function Get-GProp {
  param($Object, $Names)
  foreach ($n in $Names) {
    $v = Get-PropValue $Object $n
    if (($null -ne $v) -and ("$v" -ne '')) { return $v }
  }
  return $null
}

function Test-Enumerable {
  param($Value)
  return (($null -ne $Value) -and ($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string]) -and ($Value -isnot [System.Collections.IDictionary]))
}

function Get-CollectionValue {
  param($Object, $Name)
  if ($null -eq $Object) { return $null }
  return Get-PropValue $Object $Name
}

function ConvertTo-Bool {
  param($Value)
  if ($null -eq $Value) { return $false }
  if ($Value -is [bool]) { return $Value }
  $inner = Get-PropValue $Value 'Value'          # ClientResult<bool>
  if ($null -ne $inner) { return [bool]$inner }
  return ("$Value" -eq "true")
}

# ResourcePath (.DecodedUrl) oder String
function Get-UrlText {
  param($Value)
  if ($null -eq $Value) { return "" }
  $d = Get-PropValue $Value 'DecodedUrl'
  if ($null -ne $d) { return "$d" }
  return "$Value"
}

function Get-OrDash {
  param($Value)
  if (($null -eq $Value) -or ("$Value" -eq "")) { return "-" }
  return "$Value"
}

# Größe in MB (zwei Nachkommastellen, Komma) und Datum deutsch in lokaler Zeit
$script:DeCulture = [Globalization.CultureInfo]::GetCultureInfo("de-DE")
function Format-SizeMb {
  param($Bytes)
  try {
    $mb = [double]"$Bytes" / 1MB
    if (($mb -gt 0) -and ([Math]::Round($mb, 2) -eq 0)) { return "< 0,01" }
    return $mb.ToString("N2", $script:DeCulture)
  } catch { return "-" }
}
function Format-DeDate {
  param($Iso)
  try {
    return [datetime]::Parse("$Iso", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToLocalTime().ToString("dd.MM.yyyy HH:mm", $script:DeCulture)
  } catch { return "$Iso" }
}

function ConvertTo-HtmlEscape {
  param($s)
  if ($null -eq $s) { return "" }
  return [System.Net.WebUtility]::HtmlEncode("$s")
}

function Test-Hex { param([string]$s) return ($s -match '^[0-9A-Fa-f]{3,8}$') }
function Get-YesNo { param($v) if (ConvertTo-Bool $v) { "Ja" } else { "Nein" } }

# Pfade/Origins normalisieren: klein, ohne Schema/Host, ohne führendes '*/' und Slashes
function ConvertTo-NormPath {
  param([string]$p)
  if (-not $p) { return "" }
  $p = $p.ToLowerInvariant() -replace '^https?://[^/]+', ''
  return ($p -replace '^/+|/+$', '')
}
function ConvertTo-NormOrigin {
  param([string]$p)
  if (-not $p) { return "" }
  $p = $p.ToLowerInvariant() -replace '^\*/', ''
  return ($p -replace '^/+|/+$', '')
}
function Test-OriginCover {
  param([string]$Path, $Origins)
  if (-not $Path) { return $false }
  foreach ($o in @($Origins)) {
    $n = ConvertTo-NormOrigin "$o"
    if ($n -and (($Path -eq $n) -or $Path.StartsWith("$n/"))) { return $true }
  }
  return $false
}

# Fehlersammlung
$script:Errors = [System.Collections.Generic.List[string]]::new()
function Add-Err { param([string]$Message) $script:Errors.Add($Message) }
function Add-StepErr {
  param([string]$Label, $ErrorRecord)
  $msg = ($ErrorRecord.Exception.Message -split "`n")[0]
  if ($msg.Length -gt 200) { $msg = $msg.Substring(0, 200) }
  Add-Err "${Label}: $msg"
}

# Tenant-Grunddaten aus Graph /organization (Wrapper + verifiedDomains)
function Get-TenantInfo {
  param($TenantData)
  $orgObj = $null
  if (($null -ne $TenantData) -and (@($TenantData).Count)) {
    $wrapper = @($TenantData)[0]
    $valueArr = Get-CollectionValue $wrapper 'value'
    # 'value' ist je nach Antwort eine Liste, ein Array oder (bei genau einem Eintrag) das Objekt selbst
    if (($null -ne $valueArr) -and (@($valueArr).Count)) {
      $orgObj = @($valueArr)[0]
    } elseif ((Get-GProp $wrapper @('displayName', 'id'))) {
      $orgObj = $wrapper
    }
  }
  $orgName  = Get-GProp $orgObj @('displayName', 'Name'); if (-not $orgName) { $orgName = "N/A" }
  $tenantId = Get-GProp $orgObj @('id', 'TenantId');     if (-not $tenantId) { $tenantId = "N/A" }
  $domain = "N/A"
  $verified = Get-CollectionValue $orgObj 'verifiedDomains'
  if ((Test-Enumerable $verified) -and (@($verified).Count)) {
    $verified = @($verified)
    $def = $verified | Where-Object { (Get-PropValue $_ 'isDefault') -eq $true } | Select-Object -First 1
    if (-not $def) { $def = $verified[0] }
    $domain = Get-GProp $def @('name'); if (-not $domain) { $domain = "N/A" }
  }
  return [pscustomobject]@{ OrgName = $orgName; Domain = $domain; TenantId = $tenantId }
}

# Dateien einer Library (rekursiv, ohne Ordner) über die Listenelemente
function Get-LibraryFile {
  param($ListId, $Connection)
  $items = Get-PnPListItem -List $ListId -PageSize 500 -Connection $Connection `
    -Fields "FileLeafRef", "FileRef", "File_x0020_Size", "Modified", "FSObjType"
  foreach ($it in @($items)) {
    $fv = $it.FieldValues
    $fs = if ($fv.ContainsKey('FSObjType')) { $fv['FSObjType'] } else { $null }
    if (($null -ne $fs) -and ("$fs" -eq "1")) { continue }       # Ordner überspringen
    $mod = if ($fv.ContainsKey('Modified') -and ($fv['Modified'] -is [datetime])) {
      $fv['Modified'].ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    } else { "-" }
    [pscustomobject]@{
      Name     = Get-OrDash $(if ($fv.ContainsKey('FileLeafRef')) { $fv['FileLeafRef'] } else { $null })
      Path     = if ($fv.ContainsKey('FileRef')) { "$($fv['FileRef'])" } else { "" }
      Size     = Get-OrDash $(if ($fv.ContainsKey('File_x0020_Size')) { $fv['File_x0020_Size'] } else { $null })
      Modified = $mod
    }
  }
}

function Get-FieldValue {
  param($Item, [string[]]$Names)
  $fv = $Item.FieldValues
  foreach ($n in $Names) { if ($fv.ContainsKey($n) -and ($null -ne $fv[$n])) { return $fv[$n] } }
  return $null
}

function Get-AssetTypeLabel {
  param($Type, $Flags)
  $name = "$Type"
  $de = switch ($name) {
    "ImageDocumentLibrary"  { "Bilder" }
    "OfficeTemplateLibrary" { "Office-Vorlagen" }
    "OfficeFontLibrary"     { "Office-Schriften" }
    "BrandFontsLibrary"     { "SharePoint-Schriften" }
    "BrandKitLibrary"       { "Brand Kit" }
    "BrandColorsList"       { "Markenfarben" }
    "OrgSkillsLibrary"      { "Skills" }
    default                 { $null }
  }
  $label = if ($de) { "$de ($name)" } else { "Typ $name" }
  $f = "$Flags"
  # CopilotSearchable steht in einer eigenen Spalte; nur weitere Flags hängen am Typ
  $other = @($f -split ',\s*' | Where-Object { $_ -and ($_ -notin @("None", "0", "1", "CopilotSearchable")) })
  if ($other.Count) { $label += " (Flags: $($other -join ', '))" }
  return $label
}

# --- Anmeldung -----------------------------------------------------------------------

if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
  Exit-Script "PnP.PowerShell ist nicht installiert: Install-Module -Name PnP.PowerShell -Scope CurrentUser"
}
if (-not $Url) {
  $Url = Read-Host "SharePoint Root-Site URL (z.B. https://<tenant>.sharepoint.com)"
}
if (-not $Url) { Exit-Script "Keine SharePoint-URL angegeben." }

function Connect-Site {
  param([string]$SiteUrl)
  $p = @{ Url = $SiteUrl; ReturnConnection = $true; ErrorAction = "Stop" }
  $p.Interactive = $true      # Anmeldung im Browser; der Gerätecode ist in vielen Tenants gesperrt
  if ($ClientId) { $p.ClientId = $ClientId }
  return (Connect-PnPOnline @p)
}

try {
  $rootConn = Connect-Site $Url
} catch {
  Exit-Script "Anmeldung fehlgeschlagen: $($_.Exception.Message)"
}
$who = $Url

# --- Datenabruf (einmalig) -------------------------------------------------------------

[Console]::Error.Write("Lade Daten von Microsoft 365 ...")

# Tenant
$tenantOk = $false
$orgName = "N/A"; $domain = "N/A"; $tenantId = "N/A"
try {
  $tenantData = @(Invoke-PnPGraphMethod -Url "/organization" -Connection $rootConn)
  if ($tenantData.Count) {
    $info = Get-TenantInfo $tenantData
    $orgName = $info.OrgName; $domain = $info.Domain; $tenantId = $info.TenantId
    $tenantOk = $true
    if ($orgName -eq "N/A") {
      $first = $tenantData[0]
      $props = if ($first -is [System.Collections.IDictionary]) { @($first.Keys) -join "," } else { @($first.PSObject.Properties.Name) -join "," }
      Add-Err "Tenant-Info: Antwortformat nicht erkannt (Typ $($first.GetType().Name); Eigenschaften: $props)"
    }
  } else { Add-Err "Tenant-Info: leere Antwort" }
} catch { Add-StepErr "Tenant-Info" $_ }

# Brand Center (Konfiguration, Farben, Schriften)
$bcOk = $false; $bcConn = $null; $cfg = $null
$bcUrl = "-"; $bcEnabled = $false
$bcLibs = @()
try {
  $cfg = Get-PnPBrandCenterConfig -Connection $rootConn
  $bcUrl       = Get-OrDash (Get-PropValue $cfg 'SiteUrl')
  $bcEnabled   = ConvertTo-Bool (Get-PropValue $cfg 'IsBrandCenterSiteFeatureEnabled')
  $oa = Get-PropValue $cfg 'OrgAssets'
  $oaLibs = Get-PropValue $oa 'OrgAssetsLibraries'
  if ($null -ne $oaLibs) { $bcLibs = @($oaLibs) }
  $bcOk = $true
} catch { Add-StepErr "Brand Center" $_ }

# Verbindung zur Brand-Center-Site: das SharePoint-Token der Root-Verbindung gilt für alle Site Collections
# des Tenants und wird wiederverwendet (keine zweite Anmeldung). Mit -SeparateLogin oder bei einem
# Fehler wird stattdessen eine eigene Anmeldung durchgeführt.
if ($bcOk -and $bcUrl -ne "-") {
  if (-not $SeparateLogin) {
    try {
      $token = Get-PnPAccessToken -ResourceTypeName SharePoint -Connection $rootConn
      $bcConn = Connect-PnPOnline -Url $bcUrl -AccessToken $token -ReturnConnection -ErrorAction Stop
    } catch { $bcConn = $null }
  }
  if (-not $bcConn) {
    try { $bcConn = Connect-Site $bcUrl } catch { Add-StepErr "Brand Center (Anmeldung an der Site)" $_ }
  }
}
$siteConn = if ($bcConn) { $bcConn } else { $rootConn }

# ListId -> Typ/Flags (für die OAL-Tabelle)
$typeMap = @{}
foreach ($l in $bcLibs) {
  $typeMap[("$(Get-PropValue $l 'ListId')").ToLowerInvariant()] = $l
}

# Markenfarben
$colors = @()
if ($bcOk -and $bcConn) {
  $colorsListId = Get-PropValue $cfg 'BrandColorsListId'
  if ($colorsListId -and ("$colorsListId" -ne "00000000-0000-0000-0000-000000000000")) {
    try {
      foreach ($it in @(Get-PnPListItem -List $colorsListId -Connection $bcConn)) {
        $code = Get-FieldValue $it @('_SPColorCode', 'OData__SPColorCode')
        $vis  = Get-FieldValue $it @('_SPColorVisible', 'OData__SPColorVisible')
        $colors += [pscustomobject]@{
          Title   = Get-OrDash (Get-FieldValue $it @('_SPColorTitle', 'OData__SPColorTitle'))
          Code    = (Get-OrDash $code) -replace '^#', ''
          Visible = [bool]$vis
        }
      }
    } catch { Add-StepErr "Brand Center (Farben)" $_ }
  }
}

# Schriften für SharePoint/Viva (Brand-Fonts-Library)
$fonts = @(); $fontsFetched = $false
if ($bcOk) {
  $fontsFetched = $true
  try {
    $fontObjs = @(Get-PnPBrandCenterFont -Connection $siteConn)
    # Ein Eintrag je Objekt des Brand Centers, Werte unverändert
    $fonts = @($fontObjs | ForEach-Object {
      $hidden = Get-PropValue $_ 'IsHidden'
      $styles = Get-PropValue $_ 'FontStyles'
      $vis = Get-PropValue $_ 'IsVisible'
      [pscustomobject]@{
        File    = Get-OrDash (Get-PropValue $_ 'FileName')
        Family  = Get-OrDash (Get-PropValue $_ 'Name')
        Faces   = Get-OrDash ((@($styles) | Where-Object { $_ }) -join ", ")
        Visible = $(if ($null -ne $hidden) { Get-YesNo (-not (ConvertTo-Bool $hidden)) } elseif ($null -ne $vis) { Get-YesNo $vis } else { "-" })
      }
    } | Sort-Object File)
  } catch { Add-StepErr "Brand Center (Schriften)" $_; $fontsFetched = $false }
}

# Dateien je Asset-Library: Typ 4 (Office-Schriften + Katalog), Typ 2 (Vorlagen), Typ 1 (nur Anzahl)
$ofLibs = @()
$tplLibs = @()
$bkLibs = @()      # BrandKitLibrary (Brand Kits für Clipchamp)
$fileCounts = @{}
if ($bcOk -and $bcConn) {
  foreach ($l in $bcLibs) {
    $type = "$(Get-PropValue $l 'OrgAssetType')"
    if ($type -notin @("ImageDocumentLibrary", "OfficeTemplateLibrary", "OfficeFontLibrary", "BrandKitLibrary")) { continue }
    $name = Get-OrDash (Get-PropValue $l 'DisplayName')
    $libUrl = Get-UrlText (Get-PropValue $l 'LibraryUrl')
    $listId = Get-PropValue $l 'ListId'
    try {
      $files = @(Get-LibraryFile $listId $bcConn)
    } catch {
      Add-StepErr "Dateien in '$name'" $_
      if ($type -eq "OfficeTemplateLibrary") {
        $tplLibs += [pscustomobject]@{ Oal = $null; FilesOk = $false; Files = @() }
      }
      if ($type -eq "BrandKitLibrary") {
        $bkLibs += [pscustomobject]@{ Oal = $null; FilesOk = $false; Kits = @() }
      }
      if ($type -eq "OfficeFontLibrary") {
        $ofLibs += [pscustomobject]@{ Oal = $null; FilesOk = $false; Files = @(); CatalogPath = ""; Families = @() }
      }
      continue
    }
    $realFiles = @($files | Where-Object { $_.Path -notmatch '/Forms/' })
    $base = "/" + $libUrl.TrimStart('/') + "/"
    switch ($type) {
      "ImageDocumentLibrary" {
        $fileCounts[(ConvertTo-NormPath $libUrl)] = $realFiles.Count
      }
      "OfficeTemplateLibrary" {
        # Mehrere Libraries dieses Typs möglich: je Library ein Eintrag
        $tplFilesOfLib = @()
        foreach ($f in $realFiles) {
          $rel = if ($f.Path.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) { $f.Path.Substring($base.Length) } else { $f.Path }
          $tplFilesOfLib += [pscustomobject]@{ File = $rel; SizeMb = Format-SizeMb $f.Size; Modified = Format-DeDate $f.Modified }
        }
        $tplLibs += [pscustomobject]@{ Oal = $null; FilesOk = $true; Files = @($tplFilesOfLib | Sort-Object File) }
      }
      "BrandKitLibrary" {
        # Mehrere Libraries dieses Typs möglich: je Library ein Eintrag
        # Jedes Brand Kit ist ein Ordner der ersten Ebene (nur Ordner mit Inhalt)
        $bkNames = @($realFiles | ForEach-Object {
          $rel = if ($_.Path.StartsWith($base, [StringComparison]::OrdinalIgnoreCase)) { $_.Path.Substring($base.Length) } else { $_.Path }
          if ($rel.Contains('/')) { $rel.Split('/')[0] }
        } | Sort-Object -Unique)
        $bkLibs += [pscustomobject]@{ Oal = $null; FilesOk = $true; Kits = $bkNames }
      }
      "OfficeFontLibrary" {
        # Mehrere Libraries dieses Typs möglich: je Library ein Eintrag
        $ofLibs += [pscustomobject]@{
          Oal         = $null
          FilesOk     = $true
          Files       = @($realFiles | Where-Object { $_.Name -notmatch '\.fontconfig$' })
          CatalogPath = ""
          Families    = @()
        }
        $sfs = @()
        $cur = $ofLibs[$ofLibs.Count - 1]
        # Gelesen wird die Katalogdatei im Ordner mit der höchsten Nummer (Annahme, nicht belegt)
        $catalog = $realFiles | Where-Object { $_.Path -match '/(\d+)/fontcatalog\.fontconfig$' } |
          Sort-Object { [int]([regex]::Match($_.Path, '/(\d+)/fontcatalog\.fontconfig$').Groups[1].Value) } |
          Select-Object -Last 1
        if ($catalog) {
          $cur.CatalogPath = $catalog.Path
          try {
            $json = Get-PnPFile -Url $catalog.Path -AsString -Connection $bcConn | ConvertFrom-Json
            # Get-PropValue liefert Arrays eingepackt: erst einer Variablen zuweisen, dann mit @() umschließen
            $famList = Get-PropValue $json 'Fonts'
            $sfs = @()
            foreach ($fam in @($famList)) {
              $famName = Get-GProp $fam @('f')
              $cur.Families += $famName
              $sfList = Get-PropValue $fam 'sf'
              foreach ($sf in @($sfList)) { $sfs += [pscustomobject]@{ Id = "$(Get-GProp $sf @('id'))"; Fs = "$(Get-GProp $sf @('fs'))"; Family = $famName } }
            }
          } catch { Add-StepErr "Office-Schriften (Katalog in '$libUrl')" $_ }
        }
        # Eine Zeile je Datei: Familie nur, wenn der Katalog einen Schnitt mit gleicher ID (Dateiname ohne Endung)
        # und gleicher Dateigröße enthält; Größe in MB, Datum deutsch im lokalen Zeitraum
        $cur.Files = @($cur.Files | Sort-Object Name | ForEach-Object {
          $file = $_
          $baseName = [IO.Path]::GetFileNameWithoutExtension($file.Name)
          $fileSize = "$($file.Size)"
          $m = @($sfs | Where-Object { ($_.Id -eq $baseName) -and ($_.Fs -eq $fileSize) }) | Select-Object -First 1
          [pscustomobject]@{ File = $file.Name; Family = $(if ($m) { Get-OrDash $m.Family } else { "-" }); SizeMb = Format-SizeMb $fileSize; Modified = Format-DeDate $file.Modified }
        })
      }
    }
  }
}

$bkKitCount = @($bkLibs | ForEach-Object { $_.Kits }).Count
$tplFileCount = @($tplLibs | ForEach-Object { $_.Files }).Count
$ofFamilyCount = @($ofLibs | ForEach-Object { $_.Families }).Count
$ofFileCount = @($ofLibs | ForEach-Object { $_.Files }).Count

# Gesamtzahl der Bilder (alle Libraries vom Typ ImageDocumentLibrary)
$imageTotal = 0
foreach ($v in $fileCounts.Values) { $imageTotal += [int]$v }

# CDN: Status, Origins, Policies je Typ
$policyNames = @("IncludeFileExtensions", "ExcludeRestrictedSiteClassifications", "ExcludeIfNoScriptDisabled")
$policyOptional = @("ExcludeRestrictedSiteClassificationsFileExtensions")   # nur anzeigen, wenn der Tenant sie liefert
$blockedClassPolicy = "ExcludeRestrictedSiteClassifications"
$cdn = @{}
foreach ($t in @("Public", "Private")) {
  $entry = [ordered]@{ Status = "error"; Origins = @(); Policies = @() }
  try {
    $entry.Status = if (ConvertTo-Bool (Get-PnPTenantCdnEnabled -CdnType $t -Connection $rootConn -ErrorAction Stop)) { "on" } else { "off" }
  } catch { Add-StepErr "CDN $t (Status)" $_ }
  try {
    $entry.Origins = @(Get-PnPTenantCdnOrigin -CdnType $t -Connection $rootConn -ErrorAction Stop | ForEach-Object {
      if ($_ -is [string]) { $_ } else { Get-GProp $_ @('Url', 'Origin') }
    } | Where-Object { $_ })
  } catch { Add-StepErr "CDN $t (Origins)" $_ }
  try {
    $map = @{}
    foreach ($p in @(Get-PnPTenantCdnPolicies -CdnType $t -Connection $rootConn -ErrorAction Stop)) {
      if ($null -eq $p) { continue }
      if ($p -is [System.Collections.IDictionary]) {
        # Get-PnPTenantCdnPolicies liefert ein Dictionary<SPOTenantCdnPolicyType, string>
        foreach ($k in $p.Keys) { $map["$k"] = "$($p[$k])" }
      } elseif ($p -is [string]) {
        $i = $p.IndexOf(':'); if ($i -gt 0) { $map[$p.Substring(0, $i)] = $p.Substring($i + 1) }
      } else {
        $n = Get-GProp $p @('Policy', 'PolicyType'); $v = Get-GProp $p @('Value', 'PolicyValue')
        if ($n) { $map["$n"] = "$v" }
      }
    }
    $names = @($policyNames) + @($policyOptional | Where-Object { $map.ContainsKey($_) })
    $entry.Policies = @($names | ForEach-Object {
      [pscustomobject]@{ Name = $_; Value = $(if ($map.ContainsKey($_) -and $map[$_]) { $map[$_] } else { "(leer)" }) }
    })
  } catch { Add-StepErr "CDN $t (Policies)" $_ }
  $cdn[$t] = $entry
}

# OAL inkl. Asset-Typ, CDN-Typ und Dateianzahl
$libRows = @(); $oalOk = $true
try {
  $oalData = @(Get-PnPOrgAssetsLibrary -Connection $rootConn -ErrorAction Stop)
} catch {
  $oalData = @()
  if ($_.Exception.Message -notmatch 'no libraries') { $oalOk = $false; Add-StepErr "Org Asset Library" $_ }
}
foreach ($lib in $oalData) {
  $url = Get-UrlText (Get-PropValue $lib 'LibraryUrl')
  $listIdRaw = "$(Get-PropValue $lib 'ListId')"
  $m = [regex]::Match($listIdRaw, 'Guid\((?<g>[^)]+)\)')
  $listId = if ($m.Success) { $m.Groups['g'].Value } else { $listIdRaw }
  $norm = ConvertTo-NormPath $url
  $types = @()
  foreach ($t in @("Public", "Private")) { if (Test-OriginCover $norm $cdn[$t].Origins) { $types += $t } }
  $assetType = "-"
  $entry = $typeMap[$listId.ToLowerInvariant()]
  if ($entry) { $assetType = Get-AssetTypeLabel (Get-PropValue $entry 'OrgAssetType') (Get-PropValue $entry 'OrgAssetFlags') }
  $libRows += [pscustomobject]@{
    Name     = Get-OrDash (Get-PropValue $lib 'DisplayName')
    Url      = Get-OrDash $url
    ListId   = Get-OrDash $listId
    CdnType  = $(if ($types.Count) { $types -join ", " } else { "-" })
    Type     = $assetType
    Copilot  = $(if ($entry) { if ("$(Get-PropValue $entry 'OrgAssetFlags')" -match 'CopilotSearchable') { "Ja" } else { "Nein" } } else { "-" })
    Files    = $(if ($fileCounts.ContainsKey($norm)) { "$($fileCounts[$norm])" } else { "-" })
  }
}

# Die vom Brand Center automatisch angelegten OAL-Einträge (BrandColorsList, BrandFontsLibrary, OfficeFontLibrary)
function Get-BrandCenterOalRow {
  param([string]$TypeName, [int]$Index = 0)
  $l = $bcLibs | Where-Object { "$(Get-PropValue $_ 'OrgAssetType')" -eq $TypeName } | Select-Object -Skip $Index -First 1
  if (-not $l) { return $null }
  $url = Get-UrlText (Get-PropValue $l 'LibraryUrl')
  $norm = ConvertTo-NormPath $url
  $types = @()
  foreach ($t in @("Public", "Private")) { if (Test-OriginCover $norm $cdn[$t].Origins) { $types += $t } }
  return [pscustomobject]@{
    Name    = Get-OrDash (Get-PropValue $l 'DisplayName')
    Url     = Get-OrDash $url
    ListId  = Get-OrDash (Get-PropValue $l 'ListId')
    CdnType = $(if ($types.Count) { $types -join ", " } else { "-" })
    Type    = Get-AssetTypeLabel (Get-PropValue $l 'OrgAssetType') (Get-PropValue $l 'OrgAssetFlags')
  }
}
$bcOalColors = Get-BrandCenterOalRow "BrandColorsList"
$bcOalFonts  = Get-BrandCenterOalRow "BrandFontsLibrary"
# OAL-Einträge der Office-Schriften-Libraries (gleiche Reihenfolge wie bei der Datensammlung)
for ($i = 0; $i -lt @($ofLibs).Count; $i++) { $ofLibs[$i].Oal = Get-BrandCenterOalRow "OfficeFontLibrary" $i }
for ($i = 0; $i -lt @($tplLibs).Count; $i++) { $tplLibs[$i].Oal = Get-BrandCenterOalRow "OfficeTemplateLibrary" $i }
for ($i = 0; $i -lt @($bkLibs).Count; $i++) { $bkLibs[$i].Oal = Get-BrandCenterOalRow "BrandKitLibrary" $i }

# Nur Bilder-Libraries (ImageDocumentLibrary): Office-Schriften und -Vorlagen haben eigene Abschnitte
$imgRows = @($libRows | Where-Object { $_.Type -match 'ImageDocumentLibrary' })

# Themes: Brand-Center-Endpunkt (liefert die im Brand Center sichtbaren Themes; Get-PnPTenantTheme liefert dagegen
# nur die klassischen Tenant-Themes und war im Tenant leer)
$themeRows = @(); $themesOk = $true
try {
  $resp = Invoke-PnPSPRestMethod -Method Get -Url "/_api/brandcenter/GetTenantThemes" -Connection $siteConn -ErrorAction Stop
  $themeData = Get-PropValue $resp 'themeData'
  foreach ($theme in @($themeData)) {
    if ($null -eq $theme) { continue }
    $tj = Get-PropValue $theme 'themeJson'
    $t = $null
    if ($tj) { try { $t = if ($tj -is [string]) { $tj | ConvertFrom-Json } else { $tj } } catch { $t = $null } }
    $palette = Get-PropValue $t 'palette'
    $primary = "$(Get-GProp $palette @('themePrimary'))" -replace '^#', ''
    $inv = Get-PropValue $t 'isInverted'
    $vis = Get-PropValue $theme 'isVisible'
    $themeRows += [pscustomobject]@{
      Name     = Get-OrDash (Get-GProp $theme @('name'))
      Inverted = $(if ($null -ne $inv) { Get-YesNo $inv } else { "-" })
      Primary  = Get-OrDash $primary
      Visible  = $(if ($null -ne $vis) { Get-YesNo $vis } else { "-" })
    }
  }
} catch { $themesOk = $false; Add-StepErr "Themes" $_ }

[Console]::Error.Write("`r$esc[K")

$now = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
# Werkzeuge, mit denen der Bericht erstellt wird (geladene Modulversion, sonst die höchste installierte)
$pnpMod = Get-Module -Name PnP.PowerShell | Select-Object -First 1
if (-not $pnpMod) { $pnpMod = Get-Module -ListAvailable -Name PnP.PowerShell | Sort-Object Version -Descending | Select-Object -First 1 }
$toolInfo = "PnP.PowerShell $(if ($pnpMod) { $pnpMod.Version } else { 'unbekannt' }) (PowerShell $($PSVersionTable.PSVersion))"
$errCount = $script:Errors.Count

# Kommagetrennten Policy-Wert in Einträge zerlegen; "(leer)" ergibt eine leere Liste
function Get-ListItem {
  param([string]$Value)
  if ((-not $Value) -or ($Value -eq "(leer)")) { return @() }
  return @($Value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Get-StatusLabel {
  param($s)
  switch ($s) { "on" { "Aktiviert" } "off" { "Nicht aktiviert" } default { "Fehler beim Abruf" } }
}

# ============================================================================
# Textausgabe
# ============================================================================

function Write-Rule { param([string]$Title, [string]$Color, [string]$Indent = "") Write-Host "$Indent$Color$('-' * ($Title.Length + 2))$RESET" }
function Get-Plural { param($n, [string]$one, [string]$many) if ("$n" -eq "1") { "$n $one" } else { "$n $many" } }
# Tabelle im Terminal: Spaltenbreiten aus dem Inhalt, Kopfzeile fett
function print_table {
  # $Right wird nur in der verschachtelten Funktion fmt gelesen (dynamischer Geltungsbereich)
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter', 'Right', Justification = 'Verwendung in fmt')]
  param([string[]]$Headers, $Rows, [int[]]$Right = @())
  $rows = @($Rows)
  # Zelle "@RRGGBB|Text" ergibt ein Farbfeld (4 Zeichen + Leerzeichen) vor dem Text
  function Split-Cell($cell) {
    $c = "$cell"
    if ($c -match '^@([0-9A-Fa-f]{3,8})\|(.*)$') { return [pscustomobject]@{ Sw = $Matches[1]; Txt = $Matches[2]; Len = $Matches[2].Length + 5 } }
    return [pscustomobject]@{ Sw = ""; Txt = $c; Len = $c.Length }
  }
  $w = @(0) * $Headers.Count
  for ($i = 0; $i -lt $Headers.Count; $i++) {
    $w[$i] = $Headers[$i].Length
    foreach ($r in $rows) { $l = (Split-Cell $r[$i]).Len; if ($l -gt $w[$i]) { $w[$i] = $l } }
  }
  function fmt($cells, $style) {
    $parts = for ($i = 0; $i -lt $Headers.Count; $i++) {
      $c = Split-Cell $cells[$i]
      $shown = if ($c.Sw) { "$(Get-Swatch $c.Sw) $($c.Txt)" } else { $c.Txt }
      $pad = ' ' * [Math]::Max(0, $w[$i] - $c.Len)
      if ($Right -contains $i) { "$pad$shown" } elseif ($i -lt $Headers.Count - 1) { "$shown$pad" } else { $shown }
    }
    Write-Host "     $style$($parts -join '  ')$RESET"
  }
  fmt $Headers $BOLD
  Write-Host "     $DIM$('-' * ((($w | Measure-Object -Sum).Sum) + 2 * ($Headers.Count - 1)))$RESET"
  foreach ($r in $rows) { fmt $r "" }
}
function print_header {
  param([string]$title)
  Write-Host ""
  Write-Rule $title $BLUE
  Write-Host "$BLUE$BOLD $title $RESET"
  Write-Rule $title $BLUE
  Write-Host ""
}
function print_subheader {
  param([string]$title)
  Write-Host ""
  Write-Rule $title $CYAN " "
  Write-Host " $CYAN$BOLD $title $RESET"
  Write-Rule $title $CYAN " "
}
function print_ok    { param($msg) Write-Host "  ${GREEN}OK:${RESET}      $msg" }
function print_warn  { param($msg) Write-Host "  ${YELLOW}WARN:${RESET}    $msg" }
function print_error { param($msg) Write-Host "  ${RED}FEHLER:${RESET}  $msg" }
function print_kv    { param([string]$key, $value) Write-Host "     ${BOLD}${key}:${RESET} $value" }

function Get-Swatch {
  param([string]$Hex)
  if ($RESET -and ($Hex.Length -eq 6)) {
    $r = [Convert]::ToInt32($Hex.Substring(0, 2), 16); $g = [Convert]::ToInt32($Hex.Substring(2, 2), 16); $b = [Convert]::ToInt32($Hex.Substring(4, 2), 16)
    return "$esc[48;2;${r};${g};${b}m    $RESET"
  }
  return "    "
}

# Hervorgehobener OAL-Eintrag (Brand-Center-Library) wie im HTML-Bericht
function print_oal_entry {
  param($Row)
  if ($null -eq $Row) { return }
  Write-Host ""
  Write-Host "     $CYAN${BOLD}OAL-Eintrag: $($Row.Name)$RESET"
  print_kv "Typ" $Row.Type
  if ($Row.PSObject.Properties['Copilot']) { print_kv "Copilot-Suche" $Row.Copilot }
  if ($Row.PSObject.Properties['Files']) { print_kv "Dateien" $Row.Files }
  print_kv "URL" $Row.Url
  print_kv "List-ID" $Row.ListId
  print_kv "CDN-Typ" $Row.CdnType
}

function render_text {
  Write-Host ""
  Write-Host "$BOLD  Brand Center – Statusbericht$RESET"
  print_ok "Verbunden mit $who."
  print_kv "Erstellt mit" $toolInfo

  print_header "Tenant-Grunddaten"
  if ($tenantOk) {
    print_kv "Organisation" $orgName
    print_kv "Tenant-Domäne" $domain
    print_kv "Tenant-ID" $tenantId
  } else { print_error "Tenant-Konfiguration konnte nicht abgerufen werden." }

  print_header "Brand Center"
  if ($bcOk) {
    print_kv "Site-URL" $bcUrl
    print_kv "Feature aktiv" (Get-YesNo $bcEnabled)
  } else { print_error "Brand-Center-Konfiguration konnte nicht abgerufen werden." }

  print_subheader "Markenfarben ($($colors.Count))"
  print_oal_entry $bcOalColors
  if ($colors.Count) {
    Write-Host ""
    print_table @("Name", "Farbe", "Sichtbar") @($colors | ForEach-Object { ,@($_.Title, $(if (Test-Hex $_.Code) { "@$($_.Code)|#$($_.Code)" } else { $_.Code }), (Get-YesNo $_.Visible)) })
  } elseif ($bcOk -and $bcConn) { print_warn "Keine Markenfarben definiert." }

  print_subheader "Schriften – SharePoint / Viva Connections ($(Get-Plural (@($fonts).Count) Datei Dateien))"
  print_oal_entry $bcOalFonts
  if ($fonts.Count) {
    Write-Host ""
    print_table @("Datei", "Familie", "Schnitte", "Sichtbar") @($fonts | ForEach-Object { ,@($_.File, $_.Family, $_.Faces, $_.Visible) })
  } elseif ($fontsFetched) { print_warn "Keine Schriften hochgeladen." }
  else { print_error "Schriften konnten nicht abgerufen werden." }

  print_subheader "Themes ($($themeRows.Count))"
  if ($themeRows.Count) {
    Write-Host ""
    print_table @("Name", "Primärfarbe", "Invertiert", "Sichtbar") @($themeRows | ForEach-Object { ,@($_.Name, $(if (Test-Hex $_.Primary) { "@$($_.Primary)|$($_.Primary)" } else { $_.Primary }), $_.Inverted, $_.Visible) })
  } elseif ($themesOk) { print_warn "Keine Themes vorhanden." }
  else { print_error "Themes konnten nicht abgerufen werden." }

  print_header "Office-Schriften – OfficeFontLibrary ($(Get-Plural (@($ofLibs).Count) Library Libraries), $(Get-Plural $ofFamilyCount Familie Familien), $(Get-Plural $ofFileCount Datei Dateien))"
  if (@($ofLibs).Count) {
    foreach ($lib in $ofLibs) {
      print_oal_entry $lib.Oal
      if (-not $lib.FilesOk) { print_error "Dateien der Library konnten nicht abgerufen werden."; continue }
      if ($lib.CatalogPath) { print_kv "Katalogdatei" $lib.CatalogPath }
      else { print_warn "Kein Font-Katalog gefunden." }
      if (@($lib.Files).Count) {
        Write-Host ""
        print_table @("Datei", "Familie", "Größe (MB)", "Geändert") @($lib.Files | ForEach-Object { ,@($_.File, $_.Family, $_.SizeMb, $_.Modified) }) @(2)
      } else { print_warn "Keine Schriftdateien vorhanden." }
    }
  } else { print_warn "Keine Library vom Typ OfficeFontLibrary gefunden oder Brand Center nicht abrufbar." }

  print_header "Office-Vorlagen – OfficeTemplateLibrary ($(Get-Plural (@($tplLibs).Count) Library Libraries), $(Get-Plural $tplFileCount Vorlage Vorlagen))"
  if (@($tplLibs).Count) {
    foreach ($lib in $tplLibs) {
      print_oal_entry $lib.Oal
      if (-not $lib.FilesOk) { print_error "Dateien der Library konnten nicht abgerufen werden."; continue }
      if (@($lib.Files).Count) {
        Write-Host ""
        print_table @("Vorlage", "Größe (MB)", "Geändert") @($lib.Files | ForEach-Object { ,@($_.File, $_.SizeMb, $_.Modified) }) @(1)
      } else { print_warn "Keine Vorlagen in der Library vorhanden." }
    }
  } else { print_warn "Keine Library vom Typ OfficeTemplateLibrary gefunden." }

  print_header "Brand Kits (Clipchamp) – BrandKitLibrary ($(Get-Plural (@($bkLibs).Count) Library Libraries), $(Get-Plural $bkKitCount "Brand Kit" "Brand Kits"))"
  if (@($bkLibs).Count) {
    foreach ($lib in $bkLibs) {
      print_oal_entry $lib.Oal
      if (-not $lib.FilesOk) { print_error "Dateien der Library konnten nicht abgerufen werden."; continue }
      if (@($lib.Kits).Count) {
        Write-Host ""
        print_table @("Brand Kit") @($lib.Kits | ForEach-Object { ,@($_) })
      } else { print_warn "Keine Brand Kits in der Library vorhanden." }
    }
  } else { print_warn "Keine Library vom Typ BrandKitLibrary gefunden." }

  print_header "Bilder – ImageDocumentLibrary ($(Get-Plural (@($imgRows).Count) Library Libraries), $(Get-Plural $imageTotal Datei Dateien))"
  if ($imgRows.Count) {
    foreach ($r in $imgRows) { print_oal_entry $r }
  } elseif ($oalOk) { print_warn "Keine Bilder-Libraries (ImageDocumentLibrary) registriert." }
  else { print_error "Org Asset Library konnte nicht abgerufen werden." }

  print_header "CDN Configuration"
  foreach ($t in @("Public", "Private")) {
    print_subheader "$t CDN"
    switch ($cdn[$t].Status) {
      "on"  { print_ok "$t CDN ist aktiviert." }
      "off" { print_warn "$t CDN ist deaktiviert." }
      default { print_error "$t CDN: Status konnte nicht abgerufen werden." }
    }
    Write-Host ""
    print_kv "Anzahl Origins" $cdn[$t].Origins.Count
    foreach ($o in $cdn[$t].Origins) { Write-Host "        • $o" }
    Write-Host ""
    foreach ($p in $cdn[$t].Policies) {
      if ($p.Name -eq $blockedClassPolicy) {
        print_kv "Geblockte Klassifizierungen" "($($p.Name))"
        $items = @(Get-ListItem $p.Value)
        if ($items.Count) { foreach ($i in $items) { Write-Host "        • $i" } } else { Write-Host "        keine" }
      } else { print_kv $p.Name $p.Value }
    }
  }

  print_header "Zusammenfassung"
  Write-Host " ${BOLD}Brand Center:${RESET} $bcUrl"
  Write-Host " ${BOLD}Markenfarben:${RESET} $($colors.Count)"
  Write-Host " ${BOLD}Schriften:${RESET} $(@($fonts).Count) Datei(en)"
  Write-Host " ${BOLD}Themes:${RESET} $($themeRows.Count)"
  Write-Host " ${BOLD}Bilder:${RESET} $imageTotal Datei(en) in $($imgRows.Count) Library(en)"
  Write-Host " ${BOLD}Office-Schriften:${RESET} $ofFamilyCount Familie(n), $ofFileCount Datei(en) in $(@($ofLibs).Count) Library(en)"
  Write-Host " ${BOLD}Office-Vorlagen:${RESET} $tplFileCount in $(@($tplLibs).Count) Library(en)"
  Write-Host " ${BOLD}Brand Kits (Clipchamp):${RESET} $bkKitCount in $(@($bkLibs).Count) Library(en)"
  Write-Host " ${BOLD}OAL (gesamt):${RESET} $($libRows.Count) Library(en)"
  foreach ($t in @("Public", "Private")) {
    $c = switch ($cdn[$t].Status) { "on" { $GREEN } "off" { $YELLOW } default { $RED } }
    Write-Host " $t CDN: $c$(Get-StatusLabel $cdn[$t].Status)$RESET"
  }
  if ($errCount) {
    Write-Host ""
    Write-Host " $RED$errCount Abruf(e) fehlgeschlagen:$RESET"
    foreach ($e in $script:Errors) { Write-Host "   - $e" }
  }
  Write-Host ""
  Write-Host " ${DIM}Bericht erstellt am: $now${RESET}"
  Write-Host ""
}

# ============================================================================
# HTML-Ausgabe
# ============================================================================

function Esc { param($s) ConvertTo-HtmlEscape $s }

function render_html {
  $sb = [System.Text.StringBuilder]::new()
  function A { param([string]$s) [void]$sb.AppendLine($s) }

  # Hervorgehobener OAL-Eintrag (Brand-Center-Library) in derselben Darstellung wie die OAL-Tabelle
  function Add-OalEntry {
    param($Row)
    if ($null -eq $Row) { return }
    A "<div class=`"oal`"><div class=`"oal-title`">OAL-Eintrag: $(Esc $Row.Name)</div>"
    A "<div class=`"kv`"><b>Typ</b> $(Esc $Row.Type)</div>"
    if ($Row.PSObject.Properties['Copilot']) { A "<div class=`"kv`"><b>Copilot-Suche</b> $(Esc $Row.Copilot)</div>" }
    if ($Row.PSObject.Properties['Files']) { A "<div class=`"kv`"><b>Dateien</b> $(Esc $Row.Files)</div>" }
    A "<div class=`"kv`"><b>URL</b> <span class=`"val`">$(Esc $Row.Url)</span></div>"
    A "<div class=`"kv`"><b>List-ID</b> $(Esc $Row.ListId)</div>"
    A "<div class=`"kv`"><b>CDN-Typ</b> $(Esc $Row.CdnType)</div></div>"
  }

  A @'
<!DOCTYPE html>
<html lang="de"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Brand Center – Statusbericht</title><style>
:root{ --bg:#f4f5f7; --card:#ffffff; --ink:#1f2937; --muted:#6b7280; --line:#eef0f3; }
*{ box-sizing:border-box; }
body{ margin:0; background:var(--bg); color:var(--ink); font-family:system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif; line-height:1.55; }
header{ background:linear-gradient(135deg,#1f2937,#111827); color:#fff; padding:40px 32px; }
header h1{ margin:0 0 6px; font-size:1.9rem; }
header .meta{ color:#cbd5e1; font-size:.9rem; }
main{ max-width:1040px; margin:0 auto; padding:24px 16px 48px; }
.card{ background:var(--card); border-radius:12px; box-shadow:0 1px 3px rgba(0,0,0,.12); overflow:hidden; margin-bottom:16px; }
.card h2{ margin:0; padding:14px 18px; font-size:1.05rem; border-bottom:1px solid var(--line); background:#fafbfc; }
.card .body{ padding:16px 18px; }
table{ width:100%; border-collapse:collapse; }
th,td{ text-align:left; padding:8px 10px; border-bottom:1px solid var(--line); font-size:.87rem; vertical-align:top; }
th{ color:var(--muted); font-weight:600; }
td.val{ word-break:break-all; }
.kv{ margin:6px 0; }
.kv.sub{ margin:2px 0 2px 10px; }
.kv b{ display:inline-block; min-width:130px; color:var(--muted); }
.kv.sub b{ min-width:300px; }
.origins{ margin:4px 0 8px; padding-left:24px; font-size:.87rem; word-break:break-all; }
.badge{ display:inline-block; padding:2px 10px; border-radius:999px; font-size:.75rem; font-weight:600; color:#fff; }
.b-ok{ background:#16a34a; }
.b-no{ background:#dc2626; }
.b-warn{ background:#d97706; }
.chip{ display:inline-block; width:20px; height:20px; border-radius:6px; vertical-align:middle; border:1px solid rgba(0,0,0,.15); margin-right:8px; }
.sec{ margin:20px 0 6px; padding-top:12px; border-top:1px solid var(--line); font-size:.95rem; }
.oal{ background:#f3f7ff; border-left:3px solid #2563eb; border-radius:6px; padding:8px 12px; margin:10px 0 14px; }
.oal-title{ font-weight:600; font-size:.87rem; margin-bottom:2px; }
.oal .kv{ margin:2px 0; font-size:.87rem; }
.oal .val{ word-break:break-all; }
.empty{ color:var(--muted); font-style:italic; }
.errors{ border-left:4px solid #dc2626; }
footer{ text-align:center; color:var(--muted); font-size:.82rem; padding:24px 16px; }
</style></head><body>
'@
  A "<header><h1>Brand Center – Statusbericht</h1><div class=`"meta`">Organisation: $(Esc $orgName) &nbsp;–&nbsp; Domain: $(Esc $domain) &nbsp;–&nbsp; Erstellt: $(Esc $now)</div><div class=`"meta`">Erstellt mit: $(Esc $toolInfo)</div></header><main>"

  if ($errCount) {
    A "<div class=`"card errors`"><h2>$errCount Abruf(e) fehlgeschlagen</h2><div class=`"body`"><ul>"
    foreach ($e in $script:Errors) { A "<li>$(Esc $e)</li>" }
    A "</ul></div></div>"
  }

  A "<div class=`"card`"><h2>Tenant-Grunddaten</h2><div class=`"body`">"
  A "<div class=`"kv`"><b>Organisation</b> $(Esc $orgName)</div>"
  A "<div class=`"kv`"><b>Tenant-Domäne</b> $(Esc $domain)</div>"
  A "<div class=`"kv`"><b>Tenant-ID</b> $(Esc $tenantId)</div></div></div>"

  # Brand Center
  A "<div class=`"card`"><h2>Brand Center</h2><div class=`"body`">"
  if ($bcOk) {
    $link = if ($bcUrl -match '^https?://') { "<a href=`"$(Esc $bcUrl)`">$(Esc $bcUrl)</a>" } else { Esc $bcUrl }
    A "<div class=`"kv`"><b>Site-URL</b> $link</div>"
    A "<div class=`"kv`"><b>Feature aktiv</b> $(Get-YesNo $bcEnabled)</div>"
  } else { A "<div class=`"empty`">Brand-Center-Konfiguration nicht abrufbar.</div>" }

  A "<h3 class=`"sec`">Markenfarben ($($colors.Count))</h3>"
  Add-OalEntry $bcOalColors
  if ($colors.Count) {
    A "<table><thead><tr><th>Name</th><th>Farbe</th><th>Sichtbar</th></tr></thead><tbody>"
    foreach ($c in $colors) {
      if (Test-Hex $c.Code) {
        A "<tr><td>$(Esc $c.Title)</td><td><span class=`"chip`" style=`"background:#$($c.Code)`"></span>#$(Esc $c.Code)</td><td>$(Get-YesNo $c.Visible)</td></tr>"
      } else {
        A "<tr><td>$(Esc $c.Title)</td><td class=`"empty`">$(Esc $c.Code)</td><td>$(Get-YesNo $c.Visible)</td></tr>"
      }
    }
    A "</tbody></table>"
  } elseif ($bcOk) { A "<div class=`"empty`">Keine Markenfarben definiert.</div>" }

  # Schriften SharePoint/Viva (Unterabschnitt des Brand Centers)
  A "<h3 class=`"sec`">Schriften – SharePoint / Viva Connections ($(Get-Plural (@($fonts).Count) Datei Dateien))</h3>"
  Add-OalEntry $bcOalFonts
  if ($fonts.Count) {
    A "<table><thead><tr><th>Datei</th><th>Familie</th><th>Schnitte</th><th>Sichtbar</th></tr></thead><tbody>"
    foreach ($f in $fonts) { A "<tr><td class=`"val`">$(Esc $f.File)</td><td>$(Esc $f.Family)</td><td>$(Esc $f.Faces)</td><td>$(Esc $f.Visible)</td></tr>" }
    A "</tbody></table>"
  } elseif ($fontsFetched) { A "<div class=`"empty`">Keine Schriften hochgeladen.</div>" }
  else { A "<div class=`"empty`">Schriften nicht abrufbar.</div>" }

  # Themes (Unterabschnitt des Brand Centers)
  A "<h3 class=`"sec`">Themes ($($themeRows.Count))</h3>"
  if ($themeRows.Count) {
    A "<table><thead><tr><th>Name</th><th>Primärfarbe</th><th>Invertiert</th><th>Sichtbar</th></tr></thead><tbody>"
    foreach ($th in $themeRows) {
      if (Test-Hex $th.Primary) {
        A "<tr><td>$(Esc $th.Name)</td><td><span class=`"chip`" style=`"background:#$($th.Primary)`"></span>$(Esc $th.Primary)</td><td>$(Esc $th.Inverted)</td><td>$(Esc $th.Visible)</td></tr>"
      } else {
        A "<tr><td>$(Esc $th.Name)</td><td class=`"empty`">$(Esc $th.Primary)</td><td>$(Esc $th.Inverted)</td><td>$(Esc $th.Visible)</td></tr>"
      }
    }
    A "</tbody></table>"
  } elseif ($themesOk) { A "<div class=`"empty`">Keine Themes vorhanden.</div>" }
  else { A "<div class=`"empty`">Themes nicht abrufbar.</div>" }
  A "</div></div>"

  # Office-Schriften
  A "<div class=`"card`"><h2>Office-Schriften – OfficeFontLibrary ($(Get-Plural (@($ofLibs).Count) Library Libraries), $(Get-Plural $ofFamilyCount Familie Familien), $(Get-Plural $ofFileCount Datei Dateien))</h2><div class=`"body`">"
  if (@($ofLibs).Count) {
    foreach ($lib in $ofLibs) {
      Add-OalEntry $lib.Oal
      if (-not $lib.FilesOk) { A "<div class=`"empty`">Dateien der Library nicht abrufbar.</div>"; continue }
      if ($lib.CatalogPath) { A "<div class=`"kv`"><b>Katalogdatei</b> <span class=`"val`">$(Esc $lib.CatalogPath)</span></div>" }
      else { A "<div class=`"empty`">Kein Font-Katalog gefunden.</div>" }
      if (@($lib.Files).Count) {
        A "<table><thead><tr><th>Datei</th><th>Familie</th><th>Größe (MB)</th><th>Geändert</th></tr></thead><tbody>"
        foreach ($f in $lib.Files) { A "<tr><td class=`"val`">$(Esc $f.File)</td><td>$(Esc $f.Family)</td><td>$(Esc $f.SizeMb)</td><td>$(Esc $f.Modified)</td></tr>" }
        A "</tbody></table>"
      } else { A "<div class=`"empty`">Keine Schriftdateien vorhanden.</div>" }
    }
  } else { A "<div class=`"empty`">Keine Library vom Typ OfficeFontLibrary gefunden oder Brand Center nicht abrufbar.</div>" }
  A "</div></div>"

  # Office-Vorlagen
  A "<div class=`"card`"><h2>Office-Vorlagen – OfficeTemplateLibrary ($(Get-Plural (@($tplLibs).Count) Library Libraries), $(Get-Plural $tplFileCount Vorlage Vorlagen))</h2><div class=`"body`">"
  if (@($tplLibs).Count) {
    foreach ($lib in $tplLibs) {
      Add-OalEntry $lib.Oal
      if (-not $lib.FilesOk) { A "<div class=`"empty`">Dateien der Library nicht abrufbar.</div>"; continue }
      if (@($lib.Files).Count) {
        A "<table><thead><tr><th>Vorlage</th><th>Größe (MB)</th><th>Geändert</th></tr></thead><tbody>"
        foreach ($f in $lib.Files) { A "<tr><td class=`"val`">$(Esc $f.File)</td><td>$(Esc $f.SizeMb)</td><td>$(Esc $f.Modified)</td></tr>" }
        A "</tbody></table>"
      } else { A "<div class=`"empty`">Keine Vorlagen in der Library vorhanden.</div>" }
    }
  } else { A "<div class=`"empty`">Keine Library vom Typ OfficeTemplateLibrary gefunden.</div>" }
  A "</div></div>"

  # Brand Kits (Clipchamp)
  A "<div class=`"card`"><h2>Brand Kits (Clipchamp) – BrandKitLibrary ($(Get-Plural (@($bkLibs).Count) Library Libraries), $(Get-Plural $bkKitCount "Brand Kit" "Brand Kits"))</h2><div class=`"body`">"
  if (@($bkLibs).Count) {
    foreach ($lib in $bkLibs) {
      Add-OalEntry $lib.Oal
      if (-not $lib.FilesOk) { A "<div class=`"empty`">Dateien der Library nicht abrufbar.</div>"; continue }
      if (@($lib.Kits).Count) {
        A "<table><thead><tr><th>Brand Kit</th></tr></thead><tbody>"
        foreach ($k in $lib.Kits) { A "<tr><td>$(Esc $k)</td></tr>" }
        A "</tbody></table>"
      } else { A "<div class=`"empty`">Keine Brand Kits in der Library vorhanden.</div>" }
    }
  } else { A "<div class=`"empty`">Keine Library vom Typ BrandKitLibrary gefunden.</div>" }
  A "</div></div>"

  # OAL
  A "<div class=`"card`"><h2>Bilder – ImageDocumentLibrary ($(Get-Plural (@($imgRows).Count) Library Libraries), $(Get-Plural $imageTotal Datei Dateien))</h2><div class=`"body`">"
  if ($imgRows.Count) {
    foreach ($r in $imgRows) { Add-OalEntry $r }
  } else { A "<div class=`"empty`">Keine Bilder-Libraries (ImageDocumentLibrary) registriert.</div>" }
  A "</div></div>"

  # CDN
  foreach ($t in @("Public", "Private")) {
    $cls = switch ($cdn[$t].Status) { "on" { "b-ok" } "off" { "b-no" } default { "b-warn" } }
    A "<div class=`"card`"><h2>CDN – $t</h2><div class=`"body`">"
    A "<div class=`"kv`"><b>Status</b> <span class=`"badge $cls`">$(Esc (Get-StatusLabel $cdn[$t].Status))</span></div>"
    A "<div class=`"kv`"><b>Origins</b> $($cdn[$t].Origins.Count)</div>"
    if ($cdn[$t].Origins.Count) {
      A "<ul class=`"origins`">"
      foreach ($o in $cdn[$t].Origins) { A "<li>$(Esc $o)</li>" }
      A "</ul>"
    }
    A "<div class=`"kv`"><b>Policies</b></div>"
    if ($cdn[$t].Policies.Count) {
      foreach ($p in $cdn[$t].Policies) {
        if ($p.Name -eq $blockedClassPolicy) {
          $items = @(Get-ListItem $p.Value)
          $inner = if ($items.Count) { "<ul class=`"origins`">" + (($items | ForEach-Object { "<li>$(Esc $_)</li>" }) -join "") + "</ul>" } else { " keine" }
          A "<div class=`"kv sub`"><b>Geblockte Klassifizierungen</b> <span class=`"empty`">($(Esc $p.Name))</span>$inner</div>"
        } else {
          A "<div class=`"kv sub`"><b>$(Esc $p.Name)</b> $(Esc $p.Value)</div>"
        }
      }
    } else { A "<div class=`"empty`">Keine Policies abrufbar.</div>" }
    A "</div></div>"
  }

  A "<footer>Bericht abgeschlossen – erstellt am $(Esc $now)</footer></main></body></html>"
  return $sb.ToString()
}

# --- Ausgabe ---------------------------------------------------------------------------

if ($Format -eq "html") {
  $tmp = "$Output.tmp.$PID"
  try {
    render_html | Out-File -FilePath $tmp -Encoding utf8
    Move-Item -Path $tmp -Destination $Output -Force
    Write-Host "${GREEN}HTML-Export erfolgreich:${RESET} ${BOLD}${Output}${RESET}"
  } catch {
    if (Test-Path $tmp) { Remove-Item $tmp -Force }
    Exit-Script "HTML-Export nach '$Output' fehlgeschlagen: $($_.Exception.Message)"
  }
  if ($errCount) { [Console]::Error.WriteLine("${YELLOW}Hinweis:${RESET} $errCount Abruf(e) fehlgeschlagen, siehe Bericht.") }
} else {
  render_text
}

if ($errCount) { exit 1 }
