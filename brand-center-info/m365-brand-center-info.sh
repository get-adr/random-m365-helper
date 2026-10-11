#!/usr/bin/env bash
# ============================================================================
# m365-brand-center-info.sh – SharePoint Online Brand-Center-Bericht
# ============================================================================
# Zeigt mit der CLI for Microsoft 365 (m365):
#   • Tenant-Grunddaten
#   • Brand Center (Site-URL, Feature-Status, Markenfarben, Schriften für SharePoint/Viva, Themes)
#   • Office-Schriften (OfficeFontLibrary) und Office-Vorlagen (OfficeTemplateLibrary)
#   • OAL (Org Asset Library): Bilder-Libraries, Asset-Typ, CDN-Typ, Dateianzahl
#   • CDN Config (Public & Private): Status, Origins, Policies
#
# Ablauf: Voraussetzungen prüfen -> alle Daten einmal parallel abrufen ->
#         auswerten -> als Text oder HTML ausgeben.
#
# Voraussetzungen: m365 (angemeldet, SharePoint-Admin-Rechte) und jq.
#   npm install -g @pnp/m365-cli   |   m365 login --authType browser (sonst fragt das Script beim Start danach)
#
# Exit-Codes: 0 = vollständig, 1 = mindestens ein Abruf fehlgeschlagen,
#             2 = Aufruf-/Voraussetzungsfehler.
# ============================================================================

# jq-Programme stehen bewusst in einfachen Anführungszeichen (kein Shell-Expand)
# shellcheck disable=SC2016
set -eu
set -o pipefail
# Bash >= 5.2 deutet "&" in Ersetzungsstrings sonst als Treffer (html_escape).
shopt -u patsub_replacement 2>/dev/null || true

# --- Farben (nur bei Terminal und ohne NO_COLOR) --------------------------------

if [ -t 1 ] && [ -z "${NO_COLOR:-}" ]; then
  RED=$'\033[0;31m'; GREEN=$'\033[0;32m'; YELLOW=$'\033[1;33m'
  BLUE=$'\033[0;34m'; CYAN=$'\033[0;36m'; BOLD=$'\033[1m'
  DIM=$'\033[2m'; RESET=$'\033[0m'
else
  RED=""; GREEN=""; YELLOW=""; BLUE=""; CYAN=""; BOLD=""; DIM=""; RESET=""
fi

die() { printf '%s\n' "${RED}FEHLER:${RESET} $*" >&2; exit 2; }

usage() {
  cat <<'EOF'
Aufruf: m365-brand-center-info.sh [-f text|html] [-o DATEI]

  -f, --format  Ausgabeformat: text (Standard) oder html
  -o, --output  Zieldatei für den HTML-Bericht (impliziert --format html);
                Standard: brand-center-report-<Datum>_<Uhrzeit>.html im
                Skriptverzeichnis
  -h, --help    Diese Hilfe

Umgebung: NO_COLOR=1 schaltet Farben ab, M365_SEQUENTIAL=1 ruft m365
nacheinander statt parallel auf.
EOF
}

# --- Argumente ------------------------------------------------------------------

FORMAT=""
OUTPUT=""
while [ $# -gt 0 ]; do
  case "$1" in
    -f|--format) [ $# -ge 2 ] || die "Option $1 benötigt einen Wert."; FORMAT="$2"; shift 2;;
    -o|--output) [ $# -ge 2 ] || die "Option $1 benötigt einen Wert."; OUTPUT="$2"; shift 2;;
    -h|--help)   usage; exit 0;;
    *)           die "Unbekannte Option: $1 (siehe --help)";;
  esac
done
if [ -z "$FORMAT" ]; then
  if [ -n "$OUTPUT" ]; then FORMAT="html"; else FORMAT="text"; fi
fi
case "$FORMAT" in
  text|html) ;;
  *) die "Unbekanntes Ausgabeformat: '$FORMAT'. Erlaubt: text, html.";;
esac
if [ "$FORMAT" = "text" ] && [ -n "$OUTPUT" ]; then
  die "--output wird nur mit --format html unterstützt."
fi

if [ "$FORMAT" = "html" ] && [ -z "$OUTPUT" ]; then
  SCRIPT_DIR=$(cd "$(dirname "$0")" 2>/dev/null && pwd) || SCRIPT_DIR="."
  OUTPUT="${SCRIPT_DIR}/brand-center-report-$(date '+%Y-%m-%d_%H-%M-%S').html"
fi

# --- Voraussetzungen (für beide Formate) ----------------------------------------

command -v jq >/dev/null 2>&1 || die "jq ist nicht installiert (z. B. brew install jq)."
command -v m365 >/dev/null 2>&1 || die "CLI for Microsoft 365 (m365) ist nicht installiert: npm install -g @pnp/m365-cli"

# Anmeldung sicherstellen: bei bestehender Sitzung passiert nichts, sonst folgt die Anmeldung im Browser
# (der Gerätecode ist in vielen Tenants gesperrt). Das geschieht vor den parallelen Abrufen, damit nicht
# mehrere Prozesse gleichzeitig eine Anmeldung anfordern.
m365 login --ensure --authType browser >&2 || die "Anmeldung bei Microsoft 365 fehlgeschlagen. Bitte manuell anmelden: m365 login --authType browser"

STATUS_JSON=$(m365 status --output json 2>/dev/null || true)
if ! printf '%s' "$STATUS_JSON" | jq -e '.connectedAs // .connectionName' >/dev/null 2>&1; then
  die "Nicht bei Microsoft 365 angemeldet. Bitte zuerst anmelden: m365 login --authType browser"
fi
CONNECTED_AS=$(printf '%s' "$STATUS_JSON" | jq -r '.connectedAs // "?"')

# Werkzeuge, mit denen der Bericht erstellt wird (für Kopfzeile bzw. Textkopf)
M365_VERSION=$(m365 version --output json 2>/dev/null | jq -r 'ltrimstr("v")' 2>/dev/null || true)
JQ_VERSION=$(jq --version 2>/dev/null || true)
TOOL_INFO="CLI for Microsoft 365 ${M365_VERSION:-unbekannt} (Bash ${BASH_VERSION%%(*}, jq ${JQ_VERSION#jq-})"

TMPD=$(mktemp -d)
trap 'rm -rf "$TMPD"' EXIT
: >"$TMPD/errors.txt"

# --- Datenabruf (einmalig, parallel) ----------------------------------------------

fetch() { # name m365-Argumente...
  local name="$1" rc=0
  shift
  m365 "$@" --output json >"$TMPD/$name.out" 2>"$TMPD/$name.err" || rc=$?
  printf '%s' "$rc" >"$TMPD/$name.rc"
}
run_fetch() {
  if [ "${M365_SEQUENTIAL:-0}" = 1 ]; then fetch "$@"; else fetch "$@" & fi
}
f_ok()  { [ "$(cat "$TMPD/$1.rc" 2>/dev/null || echo 1)" = 0 ]; }
f_err() { sed -n '1p' "$TMPD/$1.err" 2>/dev/null | cut -c1-200; }

# Fehler dateibasiert sammeln, damit sie auch aus Subshells erhalten bleiben.
add_error() { printf '%s\n' "$1" >>"$TMPD/errors.txt"; }
fetch_failed() { # label name
  local msg
  msg=$(f_err "$2")
  add_error "$1: ${msg:-Abruf fehlgeschlagen}"
}

# jq auf einer Antwortdatei; bei ungültiger Antwort Fehler vermerken, nichts ausgeben.
jq_file() { # label name jq-Argumente...
  local label="$1" name="$2" out
  shift 2
  if out=$(jq "$@" "$TMPD/$name.out" 2>/dev/null); then
    printf '%s' "$out"
  else
    add_error "$label: Antwort nicht auswertbar"
  fi
}
count_lines() { if [ -z "$1" ]; then echo 0; else printf '%s\n' "$1" | wc -l | tr -d ' '; fi; }

if [ -t 2 ]; then printf 'Lade Daten von Microsoft 365 ...' >&2; fi
run_fetch tenant  tenant info get
run_fetch oal     spo orgassetslibrary list
run_fetch bc_settings spo brandcenter settings list
run_fetch bc_colors   spo brandcenter colors list
for T in Public Private; do
  run_fetch "orig_$T" spo cdn origin list --type "$T"
  run_fetch "cdn_$T"  spo cdn get --type "$T"
  run_fetch "pol_$T"  spo cdn policy list --cdnType "$T"
done
wait
if [ -t 2 ]; then printf '\r\033[K' >&2; fi

# --- Auswertung -----------------------------------------------------------------------

# Tenant
ORG_NAME="N/A"; DOMAIN="N/A"; TENANT_ID="N/A"; TENANT_OK=false
if f_ok tenant; then
  TENANT_TSV=$(jq_file "Tenant-Info" tenant -r '[(.displayName // "N/A"), (.defaultDomainName // "N/A"), (.tenantId // "N/A")] | @tsv')
  if [ -n "$TENANT_TSV" ]; then
    IFS=$'\t' read -r ORG_NAME DOMAIN TENANT_ID <<<"$TENANT_TSV" || true
    TENANT_OK=true
  fi
else
  fetch_failed "Tenant-Info" tenant
fi

# Brand Center
BC_URL="-"; BC_ENABLED="false"; BC_FONTS_URL="-"; BC_OK=false
if f_ok bc_settings; then
  BC_TSV=$(jq_file "Brand Center" bc_settings -r '
    def nz: if . == null or . == "" then "-" else tostring end;
    def u: if type == "object" then .DecodedUrl else . end;
    [ (.SiteUrl | nz),
      (.IsBrandCenterSiteFeatureEnabled // false | tostring),
      (.BrandColorsListUrl | u | nz),
      (.BrandFontLibraryUrl | u | nz) ] | @tsv')
  if [ -n "$BC_TSV" ]; then
    IFS=$'\t' read -r BC_URL BC_ENABLED _ BC_FONTS_URL <<<"$BC_TSV" || true
    BC_OK=true
  fi
else
  fetch_failed "Brand Center" bc_settings
fi

# jq-Definition: OrgAssetType/-Flags -> Bezeichnung (OAL-Tabelle und Brand-Center-Abschnitt)
TNAME_DEF='
# Enum OrgAssetType (PnP.PowerShell / Microsoft.SharePoint.Administration) und OrgAssetFlags
def tname($e):
  if $e == null then "-"
  else (({"1":"Bilder (ImageDocumentLibrary)","2":"Office-Vorlagen (OfficeTemplateLibrary)","4":"Office-Schriften (OfficeFontLibrary)","8":"SharePoint-Schriften (BrandFontsLibrary)","16":"Brand Kit (BrandKitLibrary)","32":"Markenfarben (BrandColorsList)","64":"Skills (OrgSkillsLibrary)"}[($e.t | tostring)]) // "Typ \($e.t)")
       + (if ($e.f // 0) > 1 then " (Flags: \($e.f))" else "" end)
  end;
'

# gemeinsame jq-Definitionen für OAL-Zeilen (Namen, Pfad-Normalisierung, CDN-Abdeckung)
JQ_COMMON_DEFS='
  def nz: if . == null or . == "" then "-" else tostring end;
  def norm: ascii_downcase | sub("^https?://[^/]+"; "") | gsub("^/+|/+$"; "");
  def org: ascii_downcase | sub("^\\*/"; "") | gsub("^/+|/+$"; "");
  def guid: tostring | ((capture("Guid\\((?<g>[^)]*)\\)") | .g)? // .);
  def covered($u; $list):
    $u != "" and ([$list[] | org | select(. != "" and ($u == . or ($u | startswith(. + "/"))))] | length > 0);
'

# ListId -> Asset-Typ/Flags aus der Brand-Center-Konfiguration (für die OAL-Tabelle)
BC_TYPEMAP="{}"
if f_ok bc_settings; then
  BC_TYPEMAP=$(jq_file "Brand Center (Asset-Typen)" bc_settings -c '
    [ .OrgAssets.OrgAssetsLibraries.OrgAssetsLibraries[]?
      | { key: (.ListId | tostring | ascii_downcase), value: { t: .OrgAssetType, f: .OrgAssetFlags } } ]
    | from_entries')
  BC_TYPEMAP="${BC_TYPEMAP:-{\}}"
fi

BC_COLOR_ROWS=""
if f_ok bc_colors; then
  BC_COLOR_ROWS=$(jq_file "Brand Center (Farben)" bc_colors -r '
    def nz: if . == null or . == "" then "-" else tostring end;
    (if type == "array" then . else [] end)[]
    | [ (.Title | nz), (.ColorCode // "" | tostring | ltrimstr("#") | nz), (.IsVisible // false | tostring) ]
    | @tsv')
else
  fetch_failed "Brand Center (Farben)" bc_colors
fi
BC_COLOR_COUNT=$(count_lines "$BC_COLOR_ROWS")

# Schriften (SharePoint/Viva): hängt von der Brand-Center-URL ab, daher zweiter Abruf
FONT_ROWS=""; FONT_FETCHED=false
if [ "$BC_OK" = true ] && [ "$BC_FONTS_URL" != "-" ]; then
  FONT_LIST_PATH="/${BC_FONTS_URL#*://*/}"
  FONT_FETCHED=true
  fetch fonts spo listitem list --webUrl "$BC_URL" --listUrl "$FONT_LIST_PATH" \
    --fields "FileLeafRef,OData__SPFontFamilyName,OData__SPFontFaces,OData__SPFontVisible"
  if f_ok fonts; then
    # Eine Zeile je Eintrag der Schriften-Library, Werte unverändert aus den Spalten
    FONT_ROWS=$(jq_file "Brand Center (Schriften)" fonts -r '
      def nz: if . == null or . == "" then "-" else tostring end;
      (if type == "array" then . else [] end)
      | map(select(.FileLeafRef != null))
      | sort_by(.FileLeafRef)[]
      | [ (.FileLeafRef | nz),
          (.OData__SPFontFamilyName | nz),
          (.OData__SPFontFaces | if type == "array" then map(tostring) | join(", ") else . end | nz),
          (.OData__SPFontVisible | if . == null then "-" else tostring end) ]
      | @tsv')
  else
    fetch_failed "Brand Center (Schriften)" fonts
  fi
fi
FONT_COUNT=$(count_lines "$FONT_ROWS")

# jq-Definitionen: Größe in MB (zwei Nachkommastellen, Komma) und Datum deutsch in lokaler Zeit
JQ_FMT_DEFS='
  def mb: (. * 100 / 1048576 | round) as $c
    | "\($c / 100 | floor),\($c % 100 | tostring | if length < 2 then "0" + . else . end)";
  def de: sub("\\.[0-9]+"; "") | try (fromdateiso8601 | strflocaltime("%d.%m.%Y %H:%M")) catch .;
'

# Office-Schriften (OfficeFontLibrary, Typ 4): je Library Dateien + Familien aus dem Font-Katalog.
# Es kann mehrere Libraries dieses Typs geben; die Werte liegen in Arrays je Library.
OF_LIB_COUNT=0; OF_FILE_COUNT=0; OF_FAMILY_COUNT=0
OF_FILES_OK=(); OF_FILE_ROWS=(); OF_CATALOG_PATHS=(); OF_FAMILY_ROWS=()
if [ "$BC_OK" = true ]; then
  OF_LIB_LIST=$(jq_file "Office-Schriften" bc_settings -r '
    .OrgAssets.OrgAssetsLibraries.OrgAssetsLibraries[]? | select(.OrgAssetType == 4)
    | .LibraryUrl | if type == "object" then .DecodedUrl else . end')
  while IFS= read -r of_url; do
    [ -n "$of_url" ] || continue
    i=$OF_LIB_COUNT
    OF_LIB_COUNT=$((OF_LIB_COUNT + 1))
    OF_FILES_OK[i]=false; OF_FILE_ROWS[i]=""; OF_CATALOG_PATHS[i]=""; OF_FAMILY_ROWS[i]=""
    fetch "of_files_$i" spo file list --webUrl "$BC_URL" --folderUrl "/${of_url#/}" --recursive
    if ! f_ok "of_files_$i"; then fetch_failed "Office-Schriften (Dateien in '$of_url')" "of_files_$i"; continue; fi
    OF_FILES_OK[i]=true
    OF_FILES_JSON=$(jq_file "Office-Schriften (Dateien in '$of_url')" "of_files_$i" -c '
      [ (if type == "array" then . else [] end)[]
        | select(.ServerRelativeUrl | test("/Forms/") | not)
        | select(.Name | test("\\.fontconfig$"; "i") | not)
        | { name: .Name, size: .Length, mod: .TimeLastModified } ]')
    OF_FILES_JSON="${OF_FILES_JSON:-[]}"
    echo '{}' >"$TMPD/fontcatalog_$i.json"
    # Gelesen wird die Katalogdatei im Ordner mit der höchsten Nummer (Annahme, nicht belegt)
    OF_CATALOG_PATHS[i]=$(jq_file "Office-Schriften (Katalog in '$of_url')" "of_files_$i" -r '
      [ (if type == "array" then . else [] end)[]
        | select(.ServerRelativeUrl | test("/[0-9]+/fontcatalog\\.fontconfig$"))
        | .ServerRelativeUrl ]
      | sort_by(capture("/(?<n>[0-9]+)/fontcatalog").n | tonumber) | last // empty')
    if [ -n "${OF_CATALOG_PATHS[i]}" ]; then
      fetch "of_catalog_$i" spo file get --webUrl "$BC_URL" --url "${OF_CATALOG_PATHS[i]}" --asFile --path "$TMPD/fontcatalog_$i.json"
      if f_ok "of_catalog_$i"; then
        # Familien laut Katalog (nur zum Zählen)
        if ! OF_FAMILY_ROWS[i]=$(jq -r '(.Fonts // [])[] | (.f // "-")' "$TMPD/fontcatalog_$i.json" 2>/dev/null); then
          OF_FAMILY_ROWS[i]=""; add_error "Office-Schriften (Katalog in '$of_url'): Antwort nicht auswertbar"
        fi
      else
        echo '{}' >"$TMPD/fontcatalog_$i.json"
        fetch_failed "Office-Schriften (Katalog in '$of_url')" "of_catalog_$i"
      fi
    fi
    # Eine Zeile je Datei: Familie nur, wenn der Katalog einen Schnitt mit gleicher ID (Dateiname ohne Endung)
    # und gleicher Dateigröße enthält; Größe in MB, Datum deutsch im lokalen Zeitraum
    if ! OF_FILE_ROWS[i]=$(jq -r --argjson files "$OF_FILES_JSON" "$JQ_FMT_DEFS"'
        def nz: if . == null or . == "" then "-" else tostring end;
        [ (.Fonts // [])[] | .f as $fam | (.sf // [])[] | { id: .id, fs: .fs, fam: $fam } ] as $sfs
        | $files | sort_by(.name)[] as $f
        | ($f.name | sub("\\.[^.]+$"; "")) as $b
        | ([ $sfs[] | select(.id == $b and (.fs | tostring) == ($f.size | tostring)) ][0].fam) as $fam
        | [ ($f.name | nz),
            ($fam | nz),
            (try ($f.size | tonumber | mb) catch "-"),
            (($f.mod // "-") | de) ]
        | @tsv' "$TMPD/fontcatalog_$i.json" 2>/dev/null); then
      OF_FILE_ROWS[i]=""; add_error "Office-Schriften (Dateien in '$of_url'): Antwort nicht auswertbar"
    fi
    OF_FILE_COUNT=$((OF_FILE_COUNT + $(count_lines "${OF_FILE_ROWS[i]}")))
    OF_FAMILY_COUNT=$((OF_FAMILY_COUNT + $(count_lines "${OF_FAMILY_ROWS[i]}")))
  done <<<"$OF_LIB_LIST"
fi

# Dateien in Libraries vom Typ 2 (OfficeTemplateLibrary: Auflistung) und
# Typ 1 (ImageDocumentLibrary: nur Anzahl)
TPL_LIB_COUNT=0; TPL_COUNT=0; TPL_LIB_OK=(); TPL_FILE_ROWS=(); IMG_ROWS=""; ASSET_LIB_ROWS=""
if [ "$BC_OK" = true ]; then
  ASSET_LIB_ROWS=$(jq_file "Asset-Libraries" bc_settings -r '
    .OrgAssets.OrgAssetsLibraries.OrgAssetsLibraries[]?
    | select(.OrgAssetType == 1 or .OrgAssetType == 2)
    | [ (.OrgAssetType | tostring),
        (if (.DisplayName // "") == "" then "-" else .DisplayName end),
        (.LibraryUrl | if type == "object" then .DecodedUrl else . end | if . == null or . == "" then "-" else . end) ]
    | @tsv')
fi
if [ -n "$ASSET_LIB_ROWS" ]; then
  i=0
  while IFS=$'\t' read -r atype aname aurl; do
    i=$((i + 1))
    if [ "$aurl" != "-" ]; then
      run_fetch "al_$i" spo file list --webUrl "$BC_URL" --folderUrl "/${aurl#/}" --recursive
    fi
  done <<<"$ASSET_LIB_ROWS"
  wait
  i=0
  while IFS=$'\t' read -r atype aname aurl; do
    i=$((i + 1))
    # Vorlagen-Libraries: je Library ein Eintrag (Reihenfolge wie in den Brand-Center-Einstellungen)
    if [ "$atype" = 2 ]; then
      ti=$TPL_LIB_COUNT
      TPL_LIB_COUNT=$((TPL_LIB_COUNT + 1))
      TPL_LIB_OK[ti]=false; TPL_FILE_ROWS[ti]=""
    fi
    [ "$aurl" != "-" ] || continue
    if ! f_ok "al_$i"; then fetch_failed "Dateien in '$aname'" "al_$i"; continue; fi
    if [ "$atype" = 2 ]; then
      TPL_LIB_OK[ti]=true
      TPL_FILE_ROWS[ti]=$(jq_file "Dateien in '$aname'" "al_$i" -r --arg base "/${aurl#/}/" "$JQ_FMT_DEFS"'
        def nz: if . == null or . == "" then "-" else tostring end;
        (if type == "array" then . else [] end)[]
        | select(.ServerRelativeUrl | test("/Forms/") | not)
        | [ (.ServerRelativeUrl | ltrimstr($base) | nz),
            (try (.Length | tonumber | mb) catch "-"),
            ((.TimeLastModified // "-") | de) ] | @tsv')
      TPL_COUNT=$((TPL_COUNT + $(count_lines "${TPL_FILE_ROWS[ti]}")))
    else
      cnt=$(jq_file "Dateien in '$aname'" "al_$i" -r '
        [ (if type == "array" then . else [] end)[] | select(.ServerRelativeUrl | test("/Forms/") | not) ] | length')
      if [ -n "$cnt" ]; then
        IMG_ROWS="${IMG_ROWS:+$IMG_ROWS
}${aurl}"$'\t'"${cnt}"
      fi
    fi
  done <<<"$ASSET_LIB_ROWS"
fi
IMG_TOTAL=0
if [ -n "$IMG_ROWS" ]; then
  IMG_TOTAL=$(printf '%s\n' "$IMG_ROWS" | awk -F'\t' '{ n += $2 } END { print n + 0 }')
fi
IMG_COUNT_MAP="{}"
if [ -n "$IMG_ROWS" ]; then
  IMG_COUNT_MAP=$(printf '%s\n' "$IMG_ROWS" | jq -Rnc '
    [ inputs | split("\t") | select(length == 2)
      | { key: (.[0] | ascii_downcase | ltrimstr("/")), value: (.[1] | tonumber) } ] | from_entries')
fi

# CDN: Status, Origins, Policies je Typ
POLICY_NAMES='["IncludeFileExtensions","ExcludeRestrictedSiteClassifications","ExcludeIfNoScriptDisabled"]'
# Nur anzeigen, wenn der Tenant sie liefert
POLICY_OPTIONAL='["ExcludeRestrictedSiteClassificationsFileExtensions"]'
BLOCKED_CLASS_POLICY="ExcludeRestrictedSiteClassifications"
for T in Public Private; do
  status="error"
  if f_ok "cdn_$T"; then
    v=$(tr -d '[:space:]"' <"$TMPD/cdn_$T.out")
    case "$v" in
      true)       status="on";;
      false|"")   status="off";;
      *)          add_error "CDN $T: unerwartete Antwort '$v'";;
    esac
  else
    fetch_failed "CDN $T (Status)" "cdn_$T"
  fi
  printf -v "CDN_STATUS_$T" '%s' "$status"

  origins=""
  if f_ok "orig_$T"; then
    origins=$(jq_file "CDN $T (Origins)" "orig_$T" -r 'if type == "array" then .[] | tostring else empty end')
  else
    fetch_failed "CDN $T (Origins)" "orig_$T"
  fi
  printf -v "CDN_ORIGINS_$T" '%s' "$origins"
  printf -v "CDN_OCOUNT_$T" '%s' "$(count_lines "$origins")"

  policies=""
  if f_ok "pol_$T"; then
    policies=$(jq_file "CDN $T (Policies)" "pol_$T" -r --argjson names "$POLICY_NAMES" --argjson optional "$POLICY_OPTIONAL" '
      (if type == "array" then . else [] end) as $a
      | ($names + [$optional[] | select(. as $o | any($a[]; .Policy == $o))])[] as $n
      | [$n, ([$a[] | select(.Policy == $n) | .Value][0] // "" | tostring | if . == "" then "(leer)" else . end)]
      | @tsv')
  else
    fetch_failed "CDN $T (Policies)" "pol_$T"
  fi
  printf -v "CDN_POLICIES_$T" '%s' "$policies"
done

# OAL inkl. CDN-Typ-Zuordnung (Origin deckt Library-Pfad exakt oder als Elternpfad ab)
LIB_ROWS=""
# CDN_ORIGINS_* werden weiter oben per printf -v dynamisch gesetzt
# shellcheck disable=SC2154
ORIG_PUB=$(printf '%s\n' "$CDN_ORIGINS_Public"  | jq -Rsc 'split("\n") | map(select(. != ""))')
# shellcheck disable=SC2154
ORIG_PRIV=$(printf '%s\n' "$CDN_ORIGINS_Private" | jq -Rsc 'split("\n") | map(select(. != ""))')
if f_ok oal; then
  LIB_ROWS=$(jq_file "Org Asset Library" oal -r --argjson pub "$ORIG_PUB" --argjson priv "$ORIG_PRIV" --argjson tm "$BC_TYPEMAP" --argjson fc "$IMG_COUNT_MAP" "$JQ_COMMON_DEFS$TNAME_DEF"'
    (if type == "object" then .Libraries // [] else [] end)[] as $l
    | ($l.LibraryUrl // "" | norm) as $u
    | [ (if covered($u; $pub)  then "Public"  else empty end),
        (if covered($u; $priv) then "Private" else empty end) ] as $types
    | [ ($l.DisplayName | nz),
        ($l.LibraryUrl | nz),
        ($l.ListId // "" | guid | nz),
        (if ($types | length) == 0 then "-" else $types | join(", ") end),
        tname($tm[($l.ListId // "" | guid | ascii_downcase)]),
        ($fc[($l.LibraryUrl // "" | ascii_downcase | ltrimstr("/"))] | if . == null then "-" else tostring end),
        # OrgAssetFlags: Bit 1 = CopilotSearchable
        ($tm[($l.ListId // "" | guid | ascii_downcase)] | if . == null then "-" elif ((.f // 0) % 2) == 1 then "Ja" else "Nein" end) ]
    | @tsv')
elif grep -qi 'no libraries' "$TMPD/oal.err" 2>/dev/null; then
  :  # leere OAL ist kein Fehler
else
  fetch_failed "Org Asset Library" oal
fi
LIB_COUNT=$(count_lines "$LIB_ROWS")
# Nur Bilder-Libraries (ImageDocumentLibrary): Office-Schriften und -Vorlagen haben eigene Abschnitte
IMG_LIB_ROWS=$(printf '%s\n' "$LIB_ROWS" | awk -F'\t' '$5 ~ /ImageDocumentLibrary/')
IMG_LIB_COUNT=$(count_lines "$IMG_LIB_ROWS")

# Die vom Brand Center automatisch angelegten OAL-Einträge (BrandColorsList = 32, BrandFontsLibrary = 8, OfficeFontLibrary = 4)
bc_oal_row() { # Typ-Zahl [Index innerhalb des Typs]
  jq_file "Brand Center (OAL-Eintrag)" bc_settings -r --argjson pub "$ORIG_PUB" --argjson priv "$ORIG_PRIV" \
    --argjson tm "$BC_TYPEMAP" --argjson t "$1" --argjson n "${2:-0}" "$JQ_COMMON_DEFS$TNAME_DEF"'
    [ (.OrgAssets.OrgAssetsLibraries.OrgAssetsLibraries // [])[] | select(.OrgAssetType == $t) ][$n] as $l
    | select($l != null)
    | ($l.LibraryUrl | if type == "object" then .DecodedUrl else . end) as $url
    | ($url // "" | norm) as $u
    | [ (if covered($u; $pub)  then "Public"  else empty end),
        (if covered($u; $priv) then "Private" else empty end) ] as $types
    | [ ($l.DisplayName | nz),
        ($url | nz),
        ($l.ListId // "" | guid | nz),
        (if ($types | length) == 0 then "-" else $types | join(", ") end),
        tname($tm[($l.ListId // "" | guid | ascii_downcase)]) ]
    | @tsv'
}
BC_OAL_COLORS_ROW=""; BC_OAL_FONTS_ROW=""
if [ "$BC_OK" = true ]; then
  BC_OAL_COLORS_ROW=$(bc_oal_row 32)
  BC_OAL_FONTS_ROW=$(bc_oal_row 8)
  OF_OAL_ROWS=()
  for ((i = 0; i < OF_LIB_COUNT; i++)); do OF_OAL_ROWS[i]=$(bc_oal_row 4 "$i"); done
  TPL_OAL_ROWS=()
  for ((i = 0; i < TPL_LIB_COUNT; i++)); do TPL_OAL_ROWS[i]=$(bc_oal_row 2 "$i"); done
fi

# Themes: Brand-Center-Endpunkt (liefert die im Brand Center sichtbaren Themes; "spo theme list" bzw. Get-PnPTenantTheme
# liefern dagegen nur die klassischen Tenant-Themes und waren im Tenant leer)
THEME_ROWS=""
if [ "$BC_OK" = true ]; then
  fetch themes request --url "${BC_URL%/}/_api/brandcenter/GetTenantThemes" --method get
  if f_ok themes; then
    THEME_ROWS=$(jq_file "Themes" themes -r '
      def nz: if . == null or . == "" then "-" else tostring end;
      (.themeData // [])[]
      | (((.themeJson // "{}") | if type == "string" then fromjson? else . end) // {}
         | if type == "object" then . else {} end) as $t
      | [ ((.name // $t.name) | nz),
          ($t.isInverted | if . == null then "-" else tostring end),
          ($t.palette.themePrimary // "" | tostring | ltrimstr("#") | nz),
          (.isVisible | if . == null then "-" else tostring end) ]
      | @tsv')
  else
    fetch_failed "Themes" themes
  fi
else
  add_error "Themes: Brand Center nicht abrufbar"
fi
THEME_COUNT=$(count_lines "$THEME_ROWS")

ERRORS=$(cat "$TMPD/errors.txt")
ERROR_COUNT=$(count_lines "$ERRORS")
NOW=$(date '+%Y-%m-%d %H:%M:%S')

status_label() { case "$1" in on) echo "Aktiviert";; off) echo "Nicht aktiviert";; *) echo "Fehler beim Abruf";; esac; }
yes_no() { case "$1" in true) echo "Ja";; false) echo "Nein";; *) echo "-";; esac; }
is_hex() { [[ "$1" =~ ^[0-9A-Fa-f]{3,8}$ ]]; }
cdn_var() { local n="$1_$2"; printf '%s' "${!n}"; } # cdn_var CDN_STATUS Public

# ============================================================================
# Textausgabe
# ============================================================================

print_rule() { # Titel Farbe Einrückung
  local rule
  rule=$(printf '%*s' $(( ${#1} + 2 )) '' | tr ' ' '-')
  printf '%s%s%s%s\n' "$3" "$2" "$rule" "$RESET"
}
plural() { # Anzahl Singular Plural -> "1 Library" / "2 Libraries"
  if [ "$1" = "1" ]; then printf '%s %s' "$1" "$2"; else printf '%s %s' "$1" "$3"; fi
}
print_header() {
  echo ""
  print_rule "$1" "$BLUE" ""
  printf '%s%s %s %s\n' "$BLUE" "$BOLD" "$1" "$RESET"
  print_rule "$1" "$BLUE" ""
  echo ""
}
print_subheader() {
  echo ""
  print_rule "$1" "$CYAN" " "
  printf ' %s%s %s %s\n' "$CYAN" "$BOLD" "$1" "$RESET"
  print_rule "$1" "$CYAN" " "
}
print_ok()    { printf '  %sOK:%s      %s\n' "$GREEN" "$RESET" "$*"; }
print_warn()  { printf '  %sWARN:%s    %s\n' "$YELLOW" "$RESET" "$*"; }
print_error() { printf '  %sFEHLER:%s  %s\n' "$RED" "$RESET" "$*"; }
# Farbfeld im Terminal (24-Bit-Hintergrund), nur bei aktivierten Farben und 6-stelligem Hex
color_swatch() {
  local hex="$1"
  if [ -n "$RESET" ] && [ "${#hex}" -eq 6 ]; then
    printf '\033[48;2;%d;%d;%dm    %s' "$((16#${hex:0:2}))" "$((16#${hex:2:2}))" "$((16#${hex:4:2}))" "$RESET"
  else
    printf '    '
  fi
}
# Hervorgehobener OAL-Eintrag (Brand-Center-Library) wie im HTML-Bericht
print_oal_entry() { # Zeile: Name, URL, List-ID, CDN-Typ, Typ (Tab-getrennt)
  [ -n "$1" ] || return 0
  local name url list_id cdntype atype fcount copilot
  IFS=$'\t' read -r name url list_id cdntype atype fcount copilot <<<"$1" || true
  echo ""
  printf '     %s%sOAL-Eintrag: %s%s\n' "$CYAN" "$BOLD" "$name" "$RESET"
  print_kv "Typ" "$atype"
  if [ -n "$copilot" ]; then print_kv "Copilot-Suche" "$copilot"; fi
  if [ -n "$fcount" ]; then print_kv "Dateien" "$fcount"; fi
  print_kv "URL" "$url"
  print_kv "List-ID" "$list_id"
  print_kv "CDN-Typ" "$cdntype"
}
# Tabelle im Terminal: Spaltenbreiten aus dem Inhalt, Kopfzeile fett
pt_text() { # Zelle -> PT_SW (Farbwert), PT_TXT (Text), PT_LEN (sichtbare Breite); "@RRGGBB|Text" ergibt ein Farbfeld
  local cell="$1"
  if [[ "$cell" == @*\|* ]]; then
    PT_SW="${cell%%|*}"; PT_SW="${PT_SW#@}"; PT_TXT="${cell#*|}"; PT_LEN=$(( ${#PT_TXT} + 5 ))
  else
    PT_SW=""; PT_TXT="$cell"; PT_LEN=${#cell}
  fi
}
pt_row() { # Stil Zeile (Tab-getrennt); nutzt w und right des Aufrufers
  local style="$1" line="$2" i n pad out="" shown
  local -a pt_c
  IFS=$'\t' read -ra pt_c <<<"$line" || true
  n=${#pt_c[@]}
  for ((i = 0; i < n; i++)); do
    pt_text "${pt_c[i]}"
    shown="$PT_TXT"
    if [ -n "$PT_SW" ]; then shown="$(color_swatch "$PT_SW") $PT_TXT"; fi
    pad=$(( ${w[i]:-0} - PT_LEN ))
    if [[ " $right " == *" $i "* ]]; then
      out+="$(printf '%*s' "$pad" '')${shown}"
    elif [ "$i" -lt $((n - 1)) ]; then
      out+="${shown}$(printf '%*s' "$pad" '')"
    else
      out+="$shown"
    fi
    if [ "$i" -lt $((n - 1)) ]; then out+="  "; fi
  done
  printf '     %s%s%s\n' "$style" "$out" "$RESET"
}
print_table() { # Kopfzeile Zeilen [rechtsbündige Spalten, nullbasiert, z. B. "2"]
  local header="$1" rows="$2" right="${3:-}" i n total=0 rule line
  local -a cells w
  w=()
  while IFS=$'\t' read -ra cells; do
    n=${#cells[@]}
    for ((i = 0; i < n; i++)); do
      pt_text "${cells[i]}"
      if [ "$PT_LEN" -gt "${w[i]:-0}" ]; then w[i]=$PT_LEN; fi
    done
  done <<<"$header"$'\n'"$rows"
  for ((i = 0; i < ${#w[@]}; i++)); do total=$((total + w[i] + 2)); done
  rule=$(printf '%*s' $((total - 2)) '' | tr ' ' '-')
  pt_row "$BOLD" "$header"
  printf '     %s%s%s\n' "$DIM" "$rule" "$RESET"
  while IFS= read -r line; do
    [ -n "$line" ] || continue
    pt_row "" "$line"
  done <<<"$rows"
}
print_kv()    { printf '     %s%s:%s %s\n' "$BOLD" "$1" "$RESET" "$2"; }

render_text() {
  echo ""
  printf '%s\n' "${BOLD}  Brand Center – Statusbericht${RESET}"
  print_ok "Verbunden als ${CONNECTED_AS}."
  print_kv "Erstellt mit" "$TOOL_INFO"

  print_header "Tenant-Grunddaten"
  if [ "$TENANT_OK" = true ]; then
    print_kv "Organisation" "$ORG_NAME"
    print_kv "Tenant-Domäne" "$DOMAIN"
    print_kv "Tenant-ID" "$TENANT_ID"
  else
    print_error "Tenant-Konfiguration konnte nicht abgerufen werden."
  fi

  print_header "Brand Center"
  if [ "$BC_OK" = true ]; then
    print_kv "Site-URL" "$BC_URL"
    print_kv "Feature aktiv" "$(yes_no "$BC_ENABLED")"
  else
    print_error "Brand-Center-Konfiguration konnte nicht abgerufen werden."
  fi

  print_subheader "Markenfarben ($BC_COLOR_COUNT)"
  print_oal_entry "$BC_OAL_COLORS_ROW"
  if [ "$BC_COLOR_COUNT" -gt 0 ]; then
    echo ""
    local cname ccode cvis crows="" cdisp
    while IFS=$'\t' read -r cname ccode cvis; do
      if is_hex "$ccode"; then cdisp="@${ccode}|#${ccode}"; else cdisp="$ccode"; fi
      crows+="${crows:+$'\n'}${cname}"$'\t'"${cdisp}"$'\t'"$(yes_no "$cvis")"
    done <<<"$BC_COLOR_ROWS"
    print_table $'Name\tFarbe\tSichtbar' "$crows"
  elif f_ok bc_colors; then
    print_warn "Keine Markenfarben definiert."
  fi

  print_subheader "Schriften – SharePoint / Viva Connections ($(plural "$FONT_COUNT" Datei Dateien))"
  print_oal_entry "$BC_OAL_FONTS_ROW"
  if [ "$FONT_COUNT" -gt 0 ]; then
    echo ""
    local ffile ffam faces fvis frows=""
    while IFS=$'\t' read -r ffile ffam faces fvis; do
      frows+="${frows:+$'\n'}${ffile}"$'\t'"${ffam}"$'\t'"${faces}"$'\t'"$(yes_no "$fvis")"
    done <<<"$FONT_ROWS"
    print_table $'Datei\tFamilie\tSchnitte\tSichtbar' "$frows"
  elif [ "$FONT_FETCHED" = true ] && f_ok fonts; then
    print_warn "Keine Schriften hochgeladen."
  else
    print_error "Schriften konnten nicht abgerufen werden."
  fi

  print_subheader "Themes ($THEME_COUNT)"
  if [ "$THEME_COUNT" -gt 0 ]; then
    echo ""
    local tname inv color tvis trows="" tdisp
    while IFS=$'\t' read -r tname inv color tvis; do
      if is_hex "$color"; then tdisp="@${color}|${color}"; else tdisp="$color"; fi
      trows+="${trows:+$'\n'}${tname}"$'\t'"${tdisp}"$'\t'"$(yes_no "$inv")"$'\t'"$(yes_no "$tvis")"
    done <<<"$THEME_ROWS"
    print_table $'Name\tPrimärfarbe\tInvertiert\tSichtbar' "$trows"
  elif f_ok themes; then
    print_warn "Keine Themes vorhanden."
  else
    print_error "Themes konnten nicht abgerufen werden."
  fi

  print_header "Office-Schriften – OfficeFontLibrary ($(plural "$OF_LIB_COUNT" Library Libraries), $(plural "$OF_FAMILY_COUNT" Familie Familien), $(plural "$OF_FILE_COUNT" Datei Dateien))"
  if [ "$OF_LIB_COUNT" -gt 0 ]; then
    local oi
    for ((oi = 0; oi < OF_LIB_COUNT; oi++)); do
      print_oal_entry "${OF_OAL_ROWS[oi]}"
      if ! ${OF_FILES_OK[oi]}; then print_error "Dateien der Library konnten nicht abgerufen werden."; continue; fi
      if [ -n "${OF_CATALOG_PATHS[oi]}" ]; then print_kv "Katalogdatei" "${OF_CATALOG_PATHS[oi]}"
      else print_warn "Kein Font-Katalog gefunden."; fi
      if [ -n "${OF_FILE_ROWS[oi]}" ]; then
        echo ""
        print_table $'Datei\tFamilie\tGröße (MB)\tGeändert' "${OF_FILE_ROWS[oi]}" "2"
      else
        print_warn "Keine Schriftdateien vorhanden."
      fi
    done
  else
    print_warn "Keine Library vom Typ OfficeFontLibrary gefunden oder Brand Center nicht abrufbar."
  fi

  print_header "Office-Vorlagen – OfficeTemplateLibrary ($(plural "$TPL_LIB_COUNT" Library Libraries), $(plural "$TPL_COUNT" Vorlage Vorlagen))"
  if [ "$TPL_LIB_COUNT" -gt 0 ]; then
    local ti
    for ((ti = 0; ti < TPL_LIB_COUNT; ti++)); do
      print_oal_entry "${TPL_OAL_ROWS[ti]}"
      if ! ${TPL_LIB_OK[ti]}; then print_error "Dateien der Library konnten nicht abgerufen werden."; continue; fi
      if [ -n "${TPL_FILE_ROWS[ti]}" ]; then
        echo ""
        print_table $'Vorlage\tGröße (MB)\tGeändert' "${TPL_FILE_ROWS[ti]}" "1"
      else
        print_warn "Keine Vorlagen in der Library vorhanden."
      fi
    done
  else
    print_warn "Keine Library vom Typ OfficeTemplateLibrary gefunden."
  fi

  print_header "Bilder – ImageDocumentLibrary ($(plural "$IMG_LIB_COUNT" Library Libraries), $(plural "$IMG_TOTAL" Datei Dateien))"
  if [ "$IMG_LIB_COUNT" -gt 0 ]; then
    local irow
    while IFS= read -r irow; do
      print_oal_entry "$irow"
    done <<<"$IMG_LIB_ROWS"
  elif f_ok oal || grep -qi 'no libraries' "$TMPD/oal.err" 2>/dev/null; then
    print_warn "Keine Bilder-Libraries (ImageDocumentLibrary) registriert."
  else
    print_error "Org Asset Library konnte nicht abgerufen werden."
  fi

  print_header "CDN Configuration"
  local T st origin pname pval item
  for T in Public Private; do
    print_subheader "$T CDN"
    st=$(cdn_var CDN_STATUS "$T")
    case "$st" in
      on)  print_ok "$T CDN ist aktiviert.";;
      off) print_warn "$T CDN ist deaktiviert.";;
      *)   print_error "$T CDN: Status konnte nicht abgerufen werden.";;
    esac
    echo ""
    print_kv "Anzahl Origins" "$(cdn_var CDN_OCOUNT "$T")"
    if [ "$(cdn_var CDN_OCOUNT "$T")" -gt 0 ]; then
      while IFS= read -r origin; do
        printf '        • %s\n' "$origin"
      done <<<"$(cdn_var CDN_ORIGINS "$T")"
    fi
    echo ""
    if [ -n "$(cdn_var CDN_POLICIES "$T")" ]; then
      while IFS=$'\t' read -r pname pval; do
        if [ "$pname" = "$BLOCKED_CLASS_POLICY" ]; then
          print_kv "Geblockte Klassifizierungen" "($pname)"
          if [ "$pval" = "(leer)" ]; then
            printf '        keine\n'
          else
            while IFS= read -r item; do printf '        • %s\n' "$item"; done <<<"$(printf '%s' "$pval" | tr ',' '\n' | sed 's/^ *//; s/ *$//')"
          fi
        else
          print_kv "$pname" "$pval"
        fi
      done <<<"$(cdn_var CDN_POLICIES "$T")"
    fi
  done

  print_header "Zusammenfassung"
  printf ' %sBrand Center:%s %s\n' "$BOLD" "$RESET" "$BC_URL"
  printf ' %sMarkenfarben:%s %s\n' "$BOLD" "$RESET" "$BC_COLOR_COUNT"
  printf ' %sSchriften:%s %s Datei(en)\n' "$BOLD" "$RESET" "$FONT_COUNT"
  printf ' %sThemes:%s %s\n' "$BOLD" "$RESET" "$THEME_COUNT"
  printf ' %sBilder:%s %s Datei(en) in %s Library(en)\n' "$BOLD" "$RESET" "$IMG_TOTAL" "$IMG_LIB_COUNT"
  printf ' %sOffice-Schriften:%s %s Familie(n), %s Datei(en) in %s Library(en)\n' "$BOLD" "$RESET" "$OF_FAMILY_COUNT" "$OF_FILE_COUNT" "$OF_LIB_COUNT"
  printf ' %sOffice-Vorlagen:%s %s in %s Library(en)\n' "$BOLD" "$RESET" "$TPL_COUNT" "$TPL_LIB_COUNT"
  printf ' %sOAL (gesamt):%s %s Library(en)\n' "$BOLD" "$RESET" "$LIB_COUNT"
  for T in Public Private; do
    st=$(cdn_var CDN_STATUS "$T")
    case "$st" in on) c="$GREEN";; off) c="$YELLOW";; *) c="$RED";; esac
    printf ' %s CDN: %s%s%s\n' "$T" "$c" "$(status_label "$st")" "$RESET"
  done
  if [ "$ERROR_COUNT" -gt 0 ]; then
    echo ""
    printf ' %s%s Abruf(e) fehlgeschlagen:%s\n' "$RED" "$ERROR_COUNT" "$RESET"
    printf '%s\n' "$ERRORS" | sed 's/^/   - /'
  fi
  echo ""
  printf ' %sBericht erstellt am: %s%s\n\n' "$DIM" "$NOW" "$RESET"
}
print_info_line() { printf '  Hinweis: %s\n' "$*"; }

# ============================================================================
# HTML-Ausgabe
# ============================================================================

html_escape() {
  local s="$1"
  s="${s//&/&amp;}"
  s="${s//</&lt;}"
  s="${s//>/&gt;}"
  s="${s//\"/&quot;}"
  s="${s//\'/&#39;}"
  printf '%s' "$s"
}
h() { html_escape "$1"; }

# Hervorgehobener OAL-Eintrag (Brand-Center-Library) in derselben Darstellung wie die OAL-Tabelle
html_oal_entry() { # Zeile: Name, URL, List-ID, CDN-Typ, Typ (Tab-getrennt)
  [ -n "$1" ] || return 0
  local name url list_id cdntype atype fcount copilot
  IFS=$'\t' read -r name url list_id cdntype atype fcount copilot <<<"$1" || true
  printf '<div class="oal"><div class="oal-title">OAL-Eintrag: %s</div>' "$(h "$name")"
  printf '<div class="kv"><b>Typ</b> %s</div>' "$(h "$atype")"
  if [ -n "$copilot" ]; then printf '<div class="kv"><b>Copilot-Suche</b> %s</div>' "$(h "$copilot")"; fi
  if [ -n "$fcount" ]; then printf '<div class="kv"><b>Dateien</b> %s</div>' "$(h "$fcount")"; fi
  printf '<div class="kv"><b>URL</b> <span class="val">%s</span></div>' "$(h "$url")"
  printf '<div class="kv"><b>List-ID</b> %s</div>' "$(h "$list_id")"
  printf '<div class="kv"><b>CDN-Typ</b> %s</div></div>' "$(h "$cdntype")"
}

html_cdn_card() { # Typ
  local T="$1" st cls pname pval item
  st=$(cdn_var CDN_STATUS "$T")
  case "$st" in on) cls="b-ok";; off) cls="b-no";; *) cls="b-warn";; esac
  printf '<div class="card"><h2>CDN – %s</h2><div class="body">' "$T"
  printf '<div class="kv"><b>Status</b> <span class="badge %s">%s</span></div>' "$cls" "$(h "$(status_label "$st")")"
  printf '<div class="kv"><b>Origins</b> %s</div>' "$(cdn_var CDN_OCOUNT "$T")"
  if [ "$(cdn_var CDN_OCOUNT "$T")" -gt 0 ]; then
    printf '<ul class="origins">'
    while IFS= read -r origin; do
      printf '<li>%s</li>' "$(h "$origin")"
    done <<<"$(cdn_var CDN_ORIGINS "$T")"
    printf '</ul>'
  fi
  printf '<div class="kv"><b>Policies</b></div>'
  if [ -n "$(cdn_var CDN_POLICIES "$T")" ]; then
    while IFS=$'\t' read -r pname pval; do
      if [ "$pname" = "$BLOCKED_CLASS_POLICY" ]; then
        printf '<div class="kv sub"><b>Geblockte Klassifizierungen</b> <span class="empty">(%s)</span>' "$(h "$pname")"
        if [ "$pval" = "(leer)" ]; then
          printf ' keine</div>'
        else
          printf '<ul class="origins">'
          while IFS= read -r item; do printf '<li>%s</li>' "$(h "$item")"; done <<<"$(printf '%s' "$pval" | tr ',' '\n' | sed 's/^ *//; s/ *$//')"
          printf '</ul></div>'
        fi
      else
        printf '<div class="kv sub"><b>%s</b> %s</div>' "$(h "$pname")" "$(h "$pval")"
      fi
    done <<<"$(cdn_var CDN_POLICIES "$T")"
  else
    printf '<div class="empty">Keine Policies abrufbar.</div>'
  fi
  printf '</div></div>\n'
}

render_html() {
  cat <<'EOF'
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
EOF
  printf '<header><h1>Brand Center – Statusbericht</h1><div class="meta">Organisation: %s &nbsp;–&nbsp; Domain: %s &nbsp;–&nbsp; Erstellt: %s</div><div class="meta">Erstellt mit: %s</div></header><main>\n' \
    "$(h "$ORG_NAME")" "$(h "$DOMAIN")" "$(h "$NOW")" "$(h "$TOOL_INFO")"

  if [ "$ERROR_COUNT" -gt 0 ]; then
    printf '<div class="card errors"><h2>%s Abruf(e) fehlgeschlagen</h2><div class="body"><ul>' "$ERROR_COUNT"
    while IFS= read -r line; do printf '<li>%s</li>' "$(h "$line")"; done <<<"$ERRORS"
    printf '</ul></div></div>\n'
  fi

  printf '<div class="card"><h2>Tenant-Grunddaten</h2><div class="body">'
  printf '<div class="kv"><b>Organisation</b> %s</div>' "$(h "$ORG_NAME")"
  printf '<div class="kv"><b>Tenant-Domäne</b> %s</div>' "$(h "$DOMAIN")"
  printf '<div class="kv"><b>Tenant-ID</b> %s</div>' "$(h "$TENANT_ID")"
  printf '</div></div>\n'

  printf '<div class="card"><h2>Brand Center</h2><div class="body">'
  if [ "$BC_OK" = true ]; then
    if [[ "$BC_URL" =~ ^https?:// ]]; then
      printf '<div class="kv"><b>Site-URL</b> <a href="%s">%s</a></div>' "$(h "$BC_URL")" "$(h "$BC_URL")"
    else
      printf '<div class="kv"><b>Site-URL</b> %s</div>' "$(h "$BC_URL")"
    fi
    printf '<div class="kv"><b>Feature aktiv</b> %s</div>' "$(yes_no "$BC_ENABLED")"
  else
    printf '<div class="empty">Brand-Center-Konfiguration nicht abrufbar.</div>'
  fi

  printf '<h3 class="sec">Markenfarben (%s)</h3>' "$BC_COLOR_COUNT"
  html_oal_entry "$BC_OAL_COLORS_ROW"
  if [ "$BC_COLOR_COUNT" -gt 0 ]; then
    printf '<table><thead><tr><th>Name</th><th>Farbe</th><th>Sichtbar</th></tr></thead><tbody>'
    local cname ccode cvis
    while IFS=$'\t' read -r cname ccode cvis; do
      if is_hex "$ccode"; then
        printf '<tr><td>%s</td><td><span class="chip" style="background:#%s"></span>#%s</td><td>%s</td></tr>' "$(h "$cname")" "$ccode" "$(h "$ccode")" "$(yes_no "$cvis")"
      else
        printf '<tr><td>%s</td><td class="empty">%s</td><td>%s</td></tr>' "$(h "$cname")" "$(h "$ccode")" "$(yes_no "$cvis")"
      fi
    done <<<"$BC_COLOR_ROWS"
    printf '</tbody></table>'
  elif [ "$BC_OK" = true ]; then
    printf '<div class="empty">Keine Markenfarben definiert.</div>'
  fi

  printf '<h3 class="sec">Schriften – SharePoint / Viva Connections (%s)</h3>' "$(plural "$FONT_COUNT" Datei Dateien)"
  html_oal_entry "$BC_OAL_FONTS_ROW"
  if [ "$FONT_COUNT" -gt 0 ]; then
    printf '<table><thead><tr><th>Datei</th><th>Familie</th><th>Schnitte</th><th>Sichtbar</th></tr></thead><tbody>'
    local ffile ffam faces fvis
    while IFS=$'\t' read -r ffile ffam faces fvis; do
      printf '<tr><td class="val">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>' \
        "$(h "$ffile")" "$(h "$ffam")" "$(h "$faces")" "$(yes_no "$fvis")"
    done <<<"$FONT_ROWS"
    printf '</tbody></table>'
  elif [ "$FONT_FETCHED" = true ] && f_ok fonts; then
    printf '<div class="empty">Keine Schriften hochgeladen.</div>'
  else
    printf '<div class="empty">Schriften nicht abrufbar.</div>'
  fi

  printf '<h3 class="sec">Themes (%s)</h3>' "$THEME_COUNT"
  if [ "$THEME_COUNT" -gt 0 ]; then
    printf '<table><thead><tr><th>Name</th><th>Primärfarbe</th><th>Invertiert</th><th>Sichtbar</th></tr></thead><tbody>'
    local tname inv color tvis
    while IFS=$'\t' read -r tname inv color tvis; do
      if is_hex "$color"; then
        printf '<tr><td>%s</td><td><span class="chip" style="background:#%s"></span>%s</td><td>%s</td><td>%s</td></tr>' \
          "$(h "$tname")" "$color" "$(h "$color")" "$(yes_no "$inv")" "$(yes_no "$tvis")"
      else
        printf '<tr><td>%s</td><td class="empty">%s</td><td>%s</td><td>%s</td></tr>' \
          "$(h "$tname")" "$(h "$color")" "$(yes_no "$inv")" "$(yes_no "$tvis")"
      fi
    done <<<"$THEME_ROWS"
    printf '</tbody></table>'
  elif f_ok themes; then
    printf '<div class="empty">Keine Themes vorhanden.</div>'
  else
    printf '<div class="empty">Themes nicht abrufbar.</div>'
  fi
  printf '</div></div>\n'

  printf '<div class="card"><h2>Office-Schriften – OfficeFontLibrary (%s, %s, %s)</h2><div class="body">' "$(plural "$OF_LIB_COUNT" Library Libraries)" "$(plural "$OF_FAMILY_COUNT" Familie Familien)" "$(plural "$OF_FILE_COUNT" Datei Dateien)"
  if [ "$OF_LIB_COUNT" -gt 0 ]; then
    local oi ofam oname osize omod
    for ((oi = 0; oi < OF_LIB_COUNT; oi++)); do
      html_oal_entry "${OF_OAL_ROWS[oi]}"
      if ! ${OF_FILES_OK[oi]}; then printf '<div class="empty">Dateien der Library nicht abrufbar.</div>'; continue; fi
      if [ -n "${OF_CATALOG_PATHS[oi]}" ]; then printf '<div class="kv"><b>Katalogdatei</b> <span class="val">%s</span></div>' "$(h "${OF_CATALOG_PATHS[oi]}")"
      else printf '<div class="empty">Kein Font-Katalog gefunden.</div>'; fi
      if [ -n "${OF_FILE_ROWS[oi]}" ]; then
        printf '<table><thead><tr><th>Datei</th><th>Familie</th><th>Größe (MB)</th><th>Geändert</th></tr></thead><tbody>'
        while IFS=$'\t' read -r oname ofam osize omod; do
          printf '<tr><td class="val">%s</td><td>%s</td><td>%s</td><td>%s</td></tr>' "$(h "$oname")" "$(h "$ofam")" "$(h "$osize")" "$(h "$omod")"
        done <<<"${OF_FILE_ROWS[oi]}"
        printf '</tbody></table>'
      else
        printf '<div class="empty">Keine Schriftdateien vorhanden.</div>'
      fi
    done
  else
    printf '<div class="empty">Keine Library vom Typ OfficeFontLibrary gefunden oder Brand Center nicht abrufbar.</div>'
  fi
  printf '</div></div>\n'

  printf '<div class="card"><h2>Office-Vorlagen – OfficeTemplateLibrary (%s, %s)</h2><div class="body">' "$(plural "$TPL_LIB_COUNT" Library Libraries)" "$(plural "$TPL_COUNT" Vorlage Vorlagen)"
  if [ "$TPL_LIB_COUNT" -gt 0 ]; then
    local ti tfile tsize tmod
    for ((ti = 0; ti < TPL_LIB_COUNT; ti++)); do
      html_oal_entry "${TPL_OAL_ROWS[ti]}"
      if ! ${TPL_LIB_OK[ti]}; then printf '<div class="empty">Dateien der Library nicht abrufbar.</div>'; continue; fi
      if [ -n "${TPL_FILE_ROWS[ti]}" ]; then
        printf '<table><thead><tr><th>Vorlage</th><th>Größe (MB)</th><th>Geändert</th></tr></thead><tbody>'
        while IFS=$'\t' read -r tfile tsize tmod; do
          printf '<tr><td class="val">%s</td><td>%s</td><td>%s</td></tr>' "$(h "$tfile")" "$(h "$tsize")" "$(h "$tmod")"
        done <<<"${TPL_FILE_ROWS[ti]}"
        printf '</tbody></table>'
      else
        printf '<div class="empty">Keine Vorlagen in der Library vorhanden.</div>'
      fi
    done
  else
    printf '<div class="empty">Keine Library vom Typ OfficeTemplateLibrary gefunden.</div>'
  fi
  printf '</div></div>\n'

  # Office-Schriften und -Vorlagen haben eigene Blöcke; hier nur die Bilder-Libraries (ImageDocumentLibrary)
  printf '<div class="card"><h2>Bilder – ImageDocumentLibrary (%s, %s)</h2><div class="body">' "$(plural "$IMG_LIB_COUNT" Library Libraries)" "$(plural "$IMG_TOTAL" Datei Dateien)"
  if [ "$IMG_LIB_COUNT" -gt 0 ]; then
    local irow
    while IFS= read -r irow; do
      html_oal_entry "$irow"
    done <<<"$IMG_LIB_ROWS"
  else
    printf '<div class="empty">Keine Bilder-Libraries (ImageDocumentLibrary) registriert.</div>'
  fi
  printf '</div></div>\n'

  html_cdn_card Public
  html_cdn_card Private

  printf '<footer>Bericht abgeschlossen – erstellt am %s</footer></main></body></html>\n' "$(h "$NOW")"
}

# --- Ausgabe -----------------------------------------------------------------------

if [ "$FORMAT" = "html" ]; then
  TMP_OUT="${OUTPUT}.tmp.$$"
  if render_html >"$TMP_OUT" && mv "$TMP_OUT" "$OUTPUT"; then
    printf '%s\n' "${GREEN}HTML-Export erfolgreich:${RESET} ${BOLD}${OUTPUT}${RESET}"
  else
    rm -f "$TMP_OUT"
    die "HTML-Export nach '$OUTPUT' fehlgeschlagen."
  fi
  if [ "$ERROR_COUNT" -gt 0 ]; then
    printf '%s\n' "${YELLOW}Hinweis:${RESET} $ERROR_COUNT Abruf(e) fehlgeschlagen, siehe Bericht." >&2
  fi
else
  render_text
fi

[ "$ERROR_COUNT" -eq 0 ] || exit 1
