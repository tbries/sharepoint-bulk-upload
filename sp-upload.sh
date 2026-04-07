#!/usr/bin/env bash
# sp-upload.sh — Bulk-upload a local folder tree to a SharePoint document library
# Wraps the Microsoft 365 CLI (https://pnp.github.io/cli-microsoft365/)
set -euo pipefail

###############################################################################
# Defaults
###############################################################################
SOURCE=""
SITE_URL=""
LIBRARY=""
REMOTE_PATH=""
DRY_RUN=false
LOG_FILE=""
LEDGER_FILE=""

###############################################################################
# Helpers
###############################################################################
readonly RED=$'\033[0;31m'
readonly GREEN=$'\033[0;32m'
readonly YELLOW=$'\033[0;33m'
readonly CYAN=$'\033[0;36m'
readonly RESET=$'\033[0m'

ts() { date "+%Y-%m-%d %H:%M:%S"; }

log()  { local msg; msg="[$(ts)] $*"; echo "${CYAN}${msg}${RESET}"; [ -n "$LOG_FILE" ] && echo "$msg" >> "$LOG_FILE"; }
ok()   { local msg; msg="[$(ts)] ✓ $*"; echo "${GREEN}${msg}${RESET}"; [ -n "$LOG_FILE" ] && echo "$msg" >> "$LOG_FILE"; }
warn() { local msg; msg="[$(ts)] ⚠ $*"; echo "${YELLOW}${msg}${RESET}"; [ -n "$LOG_FILE" ] && echo "$msg" >> "$LOG_FILE"; }
err()  { local msg; msg="[$(ts)] ✗ $*"; echo "${RED}${msg}${RESET}" >&2; [ -n "$LOG_FILE" ] && echo "$msg" >> "$LOG_FILE"; }

usage() {
  cat <<EOF
Usage: $(basename "$0") [options]

Upload a local folder tree to a SharePoint document library.

Required:
  --source <path>       Local folder to upload
  --site-url <url>      SharePoint site URL
  --library <name>      Target document library (e.g. "Shared Documents")

Optional:
  --remote-path <path>  Sub-path inside the library (default: library root)
  --dry-run             Preview operations without executing
  --log <path>          Log file (default: ./sp-upload.log)
  --ledger <path>       Ledger file for resume tracking
                        (default: <source>/.sp-upload-ledger)
  -h, --help            Show this help message

Prerequisites:
  • Microsoft 365 CLI installed (npm i -g @pnp/cli-microsoft365)
  • Logged in via: m365 login
EOF
  exit 0
}

die() { err "$@"; exit 1; }

###############################################################################
# Argument parsing
###############################################################################
parse_args() {
  while [[ $# -gt 0 ]]; do
    case "$1" in
      --source)      SOURCE="$2";      shift 2 ;;
      --site-url)    SITE_URL="$2";    shift 2 ;;
      --library)     LIBRARY="$2";     shift 2 ;;
      --remote-path) REMOTE_PATH="$2"; shift 2 ;;
      --dry-run)     DRY_RUN=true;     shift   ;;
      --log)         LOG_FILE="$2";    shift 2 ;;
      --ledger)      LEDGER_FILE="$2"; shift 2 ;;
      -h|--help)     usage ;;
      *) die "Unknown option: $1" ;;
    esac
  done

  [[ -z "$SOURCE" ]]   && die "Missing required option: --source"
  [[ -z "$SITE_URL" ]] && die "Missing required option: --site-url"
  [[ -z "$LIBRARY" ]]  && die "Missing required option: --library"

  # Resolve source to absolute path
  SOURCE="$(cd "$SOURCE" && pwd)"

  # Defaults that depend on parsed values
  [[ -z "$LOG_FILE" ]]    && LOG_FILE="./sp-upload.log"
  [[ -z "$LEDGER_FILE" ]] && LEDGER_FILE="${SOURCE}/.sp-upload-ledger"
}

###############################################################################
# Pre-flight checks
###############################################################################
preflight() {
  command -v m365 >/dev/null 2>&1 || die "m365 CLI not found. Install with: npm i -g @pnp/cli-microsoft365"

  if ! $DRY_RUN; then
    log "Checking m365 login status..."
    local status
    status="$(m365 status -o text 2>&1)" || true
    if echo "$status" | grep -qi "logged out"; then
      die "Not logged in to m365. Run: m365 login"
    fi
    ok "Authenticated to Microsoft 365"
  else
    warn "Dry-run mode — skipping authentication check"
  fi

  [[ -d "$SOURCE" ]] || die "Source directory does not exist: $SOURCE"
  ok "Source directory: $SOURCE"
}

###############################################################################
# Ledger helpers — SHA-256 based resume tracking
###############################################################################
ledger_hash() {
  # Returns the sha256 hash of a file
  shasum -a 256 "$1" | awk '{print $1}'
}

ledger_lookup() {
  # Check if a relative path + hash already exists in the ledger
  local rel_path="$1" hash="$2"
  [[ -f "$LEDGER_FILE" ]] || return 1
  grep -qF "${rel_path}	${hash}" "$LEDGER_FILE" 2>/dev/null
}

ledger_record() {
  # Record a successful upload
  local rel_path="$1" hash="$2"
  # Remove any previous entry for this path (content may have changed)
  if [[ -f "$LEDGER_FILE" ]]; then
    local tmp="${LEDGER_FILE}.tmp"
    grep -vF "${rel_path}	" "$LEDGER_FILE" > "$tmp" 2>/dev/null || true
    mv "$tmp" "$LEDGER_FILE"
  fi
  echo -e "${rel_path}\t${hash}" >> "$LEDGER_FILE"
}

###############################################################################
# Build the remote base folder path
###############################################################################
remote_base() {
  local base="$LIBRARY"
  if [[ -n "$REMOTE_PATH" ]]; then
    base="${base}/${REMOTE_PATH}"
  fi
  # Strip trailing slashes
  echo "${base%/}"
}

###############################################################################
# Upload workflow
###############################################################################
upload_all() {
  local base
  base="$(remote_base)"

  local total=0 uploaded=0 skipped=0 failed=0 dirs_created=0

  # Collect unique directories (relative to SOURCE), sorted shortest-first
  local dirs=()
  while IFS= read -r -d '' dir; do
    local rel="${dir#"$SOURCE"}"
    rel="${rel#/}"
    [[ -z "$rel" ]] && continue
    dirs+=("$rel")
  done < <(find "$SOURCE" -type d -print0 | sort -z)

  # Create remote folders
  for rel_dir in "${dirs[@]}"; do
    local remote_folder="${base}/${rel_dir}"
    if $DRY_RUN; then
      log "[DRY-RUN] Would create folder: ${remote_folder}"
    else
      log "Creating folder: ${remote_folder}"
      # Extract parent and leaf
      local parent leaf
      parent="$(dirname "$remote_folder")"
      leaf="$(basename "$remote_folder")"
      if m365 spo folder add \
            --webUrl "$SITE_URL" \
            --parentFolderUrl "$parent" \
            --name "$leaf" \
            --ensureParentFolders \
            -o none 2>&1; then
        ok "Folder ready: ${remote_folder}"
        (( dirs_created++ ))
      else
        warn "Folder may already exist (continuing): ${remote_folder}"
      fi
    fi
  done

  # Upload files
  while IFS= read -r -d '' file; do
    local rel="${file#"$SOURCE/"}"
    (( total++ ))

    local hash
    hash="$(ledger_hash "$file")"

    if ledger_lookup "$rel" "$hash"; then
      (( skipped++ ))
      log "Skipping (already uploaded): ${rel}"
      continue
    fi

    local remote_folder="${base}"
    local file_dir
    file_dir="$(dirname "$rel")"
    [[ "$file_dir" != "." ]] && remote_folder="${base}/${file_dir}"

    if $DRY_RUN; then
      log "[DRY-RUN] Would upload: ${rel} → ${remote_folder}/"
      continue
    fi

    log "Uploading: ${rel} → ${remote_folder}/"
    if m365 spo file add \
          --webUrl "$SITE_URL" \
          --folder "$remote_folder" \
          --path "$file" \
          -o none 2>&1; then
      ok "Uploaded: ${rel}"
      ledger_record "$rel" "$hash"
      (( uploaded++ ))
    else
      err "Failed: ${rel}"
      (( failed++ ))
    fi
  done < <(find "$SOURCE" -type f ! -name ".sp-upload-ledger" -print0 | sort -z)

  # Summary
  echo ""
  log "═══════════════════════════════════════"
  log "Upload complete"
  log "  Directories created : ${dirs_created}"
  log "  Files total         : ${total}"
  log "  Uploaded            : ${uploaded}"
  log "  Skipped (unchanged) : ${skipped}"
  log "  Failed              : ${failed}"
  log "═══════════════════════════════════════"

  if $DRY_RUN; then
    warn "Dry-run mode — no changes were made."
  fi

  [[ $failed -gt 0 ]] && return 1
  return 0
}

###############################################################################
# Main
###############################################################################
main() {
  parse_args "$@"
  preflight

  log "═══════════════════════════════════════"
  log "SharePoint Bulk Upload"
  log "  Source      : $SOURCE"
  log "  Site URL    : $SITE_URL"
  log "  Library     : $LIBRARY"
  log "  Remote path : ${REMOTE_PATH:-(root)}"
  log "  Ledger      : $LEDGER_FILE"
  log "  Dry-run     : $DRY_RUN"
  log "═══════════════════════════════════════"
  echo ""

  upload_all
}

main "$@"
