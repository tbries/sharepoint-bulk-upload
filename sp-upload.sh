#!/usr/bin/env bash
# sp-upload.sh — Bulk-upload a local folder tree to a SharePoint document library
# Uses Azure CLI for auth + Microsoft Graph API (no app registration required)
set -euo pipefail

###############################################################################
# Constants
###############################################################################
readonly GRAPH_BASE="https://graph.microsoft.com/v1.0"
readonly SMALL_FILE_LIMIT=$((4 * 1024 * 1024))  # 4 MB threshold
readonly TOKEN_TTL=2700                          # Refresh token every 45 min
readonly MIN_CHUNK_SIZE=$((320 * 1024))          # 320 KiB — Graph API minimum
readonly MAX_CHUNK_SIZE=$((60 * 1024 * 1024))    # 60 MiB  — Graph API maximum
readonly MAX_RETRIES=5                            # Max retry attempts on throttling
readonly INITIAL_BACKOFF=5                        # Initial retry backoff (seconds)
readonly MAX_BACKOFF=120                          # Max retry backoff (seconds)

###############################################################################
# Global state
###############################################################################
SOURCE=""
SITE_URL=""
LIBRARY=""
REMOTE_PATH=""
CHUNK_SIZE=$((10 * 1024 * 1024))  # default 10 MiB per chunk
DRY_RUN=false
LOG_FILE=""
LEDGER_FILE=""
ACCESS_TOKEN=""
TOKEN_ACQUIRED_AT=0
SITE_ID=""
DRIVE_ID=""
_HTTP_BODY=""
_HTTP_CODE=""

###############################################################################
# Helpers
###############################################################################
readonly RED=$'\033[0;31m'
readonly GREEN=$'\033[0;32m'
readonly YELLOW=$'\033[0;33m'
readonly CYAN=$'\033[0;36m'
readonly RESET=$'\033[0m'

ts() { date "+%Y-%m-%d %H:%M:%S"; }

log()  { local msg; msg="[$(ts)] $*"; echo "${CYAN}${msg}${RESET}"; [ -n "${LOG_FILE:-}" ] && echo "$msg" >> "$LOG_FILE"; }
ok()   { local msg; msg="[$(ts)] ✓ $*"; echo "${GREEN}${msg}${RESET}"; [ -n "${LOG_FILE:-}" ] && echo "$msg" >> "$LOG_FILE"; }
warn() { local msg; msg="[$(ts)] ⚠ $*"; echo "${YELLOW}${msg}${RESET}"; [ -n "${LOG_FILE:-}" ] && echo "$msg" >> "$LOG_FILE"; }
err()  { local msg; msg="[$(ts)] ✗ $*"; echo "${RED}${msg}${RESET}" >&2; [ -n "${LOG_FILE:-}" ] && echo "$msg" >> "$LOG_FILE"; }

die() { err "$@"; exit 1; }

usage() {
  cat <<EOF
Usage: $(basename "$0") [options]

Upload a local folder tree to a SharePoint document library.
Uses Azure CLI for authentication (no app registration required).

Required:
  --source <path>       Local folder to upload
  --site-url <url>      SharePoint site URL
  --library <name>      Target document library (e.g. "Shared Documents")

Optional:
  --remote-path <path>  Sub-path inside the library (default: library root)
  --chunk-size <bytes>  Upload chunk size in bytes (default: 10485760 / 10 MiB)
                        Min: 327680 (320 KiB), Max: 62914560 (60 MiB)
                        Must be a multiple of 327680 (320 KiB)
  --dry-run             Preview operations without executing
  --log <path>          Log file (default: ./sp-upload.log)
  --ledger <path>       Ledger file for resume tracking
                        (default: <source>/.sp-upload-ledger)
  -h, --help            Show this help message

Prerequisites:
  • Azure CLI installed (https://aka.ms/installazurecli)
  • Logged in via: az login
  • jq installed (brew install jq / apt install jq)
  • curl installed
EOF
  exit 0
}

human_size() {
  local bytes="$1"
  if (( bytes >= 1073741824 )); then
    echo "$(( bytes / 1073741824 )) GB"
  elif (( bytes >= 1048576 )); then
    echo "$(( bytes / 1048576 )) MB"
  elif (( bytes >= 1024 )); then
    echo "$(( bytes / 1024 )) KB"
  else
    echo "${bytes} B"
  fi
}

human_time() {
  local secs="$1"
  if (( secs <= 0 )); then
    echo "< 1s"
    return
  fi
  local h=$(( secs / 3600 ))
  local m=$(( (secs % 3600) / 60 ))
  local s=$(( secs % 60 ))
  local out=""
  (( h > 0 )) && out="${h}h "
  (( m > 0 )) && out="${out}${m}m "
  out="${out}${s}s"
  echo "$out"
}

get_file_size() {
  stat -f%z "$1" 2>/dev/null || stat -c%s "$1" 2>/dev/null
}

# Extract error details from a Graph API JSON response body.
# Outputs a one-liner like: code=accessDenied, message=Access denied
_graph_error_detail() {
  local body="${1:-$_HTTP_BODY}"
  local code msg
  code="$(printf '%s' "$body" | jq -r '.error.code // empty' 2>/dev/null)"
  msg="$(printf '%s' "$body" | jq -r '.error.message // empty' 2>/dev/null)"
  if [[ -n "$code" || -n "$msg" ]]; then
    echo "code=${code:-unknown}, message=${msg:-<none>}"
  else
    # Truncate raw body to avoid flooding logs
    echo "${body:0:500}"
  fi
}

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
      --chunk-size)  CHUNK_SIZE="$2";  shift 2 ;;
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

  SOURCE="$(cd "$SOURCE" && pwd)"
  [[ -z "$LOG_FILE" ]]    && LOG_FILE="./sp-upload.log"
  [[ -z "$LEDGER_FILE" ]] && LEDGER_FILE="${SOURCE}/.sp-upload-ledger"

  # Validate chunk size
  if ! [[ "$CHUNK_SIZE" =~ ^[0-9]+$ ]]; then
    die "--chunk-size must be a positive integer (bytes), got: $CHUNK_SIZE"
  fi
  if (( CHUNK_SIZE < MIN_CHUNK_SIZE )); then
    die "--chunk-size must be at least $MIN_CHUNK_SIZE bytes (320 KiB), got: $CHUNK_SIZE"
  fi
  if (( CHUNK_SIZE > MAX_CHUNK_SIZE )); then
    die "--chunk-size must be at most $MAX_CHUNK_SIZE bytes (60 MiB), got: $CHUNK_SIZE"
  fi
  if (( CHUNK_SIZE % MIN_CHUNK_SIZE != 0 )); then
    die "--chunk-size must be a multiple of $MIN_CHUNK_SIZE bytes (320 KiB), got: $CHUNK_SIZE"
  fi
}

###############################################################################
# Pre-flight checks
###############################################################################
preflight() {
  command -v az   >/dev/null 2>&1 || die "Azure CLI not found. Install: https://aka.ms/installazurecli"
  command -v curl >/dev/null 2>&1 || die "curl not found."
  command -v jq   >/dev/null 2>&1 || die "jq not found. Install: brew install jq"

  if ! $DRY_RUN; then
    log "Checking Azure CLI login status..."
    if ! az account show -o none 2>/dev/null; then
      die "Not logged in. Run: az login"
    fi
    ok "Authenticated via Azure CLI"
  else
    warn "Dry-run mode — skipping authentication check"
  fi

  [[ -d "$SOURCE" ]] || die "Source directory does not exist: $SOURCE"
  ok "Source directory: $SOURCE"
}

###############################################################################
# Token management — refreshes every 45 min (tokens valid ~1 hour)
###############################################################################
refresh_token() {
  local now
  now="$(date +%s)"
  if (( now - TOKEN_ACQUIRED_AT < TOKEN_TTL )) && [[ -n "$ACCESS_TOKEN" ]]; then
    return 0
  fi

  ACCESS_TOKEN="$(az account get-access-token \
    --resource https://graph.microsoft.com \
    --query accessToken -o tsv 2>/dev/null)" || true
  [[ -z "$ACCESS_TOKEN" ]] && die "Failed to acquire access token for Microsoft Graph"
  TOKEN_ACQUIRED_AT="$now"
}

###############################################################################
# URL-encode a path (preserving forward slashes)
###############################################################################
urlencode_path() {
  local path="$1"
  local encoded=""
  local i c hex
  for (( i = 0; i < ${#path}; i++ )); do
    c="${path:$i:1}"
    case "$c" in
      [a-zA-Z0-9.~_/-]) encoded+="$c" ;;
      ' ') encoded+="%20" ;;
      *) printf -v hex '%%%02X' "'$c"; encoded+="$hex" ;;
    esac
  done
  echo "$encoded"
}

###############################################################################
# Retry-aware HTTP client — backs off on 429 / 503, then aborts
# Sets globals: _HTTP_BODY, _HTTP_CODE
# Usage: graph_curl [curl-args...]  (omit -sS / -w / -o / -D)
###############################################################################
graph_curl() {
  local attempt=0 backoff=$INITIAL_BACKOFF
  local raw_response header_file
  header_file="$(mktemp)"
  trap "rm -f '${header_file}'" RETURN

  while true; do
    raw_response="$(curl -sS -D "$header_file" -w "\n%{http_code}" "$@")"
    _HTTP_CODE="$(echo "$raw_response" | tail -1)"
    _HTTP_BODY="$(echo "$raw_response" | sed '$d')"

    case "$_HTTP_CODE" in
      429|503)
        (( attempt++ )) || true
        if (( attempt > MAX_RETRIES )); then
          local detail; detail="$(_graph_error_detail)"
          die "Persistent throttling (HTTP $_HTTP_CODE) after ${MAX_RETRIES} retries — aborting. Detail: ${detail}. Re-run to resume."
        fi

        local retry_after=""
        retry_after="$(grep -i '^Retry-After:' "$header_file" 2>/dev/null \
          | head -1 | awk '{print $2}' | tr -d '\r\n')"

        local wait_time="$backoff"
        if [[ -n "$retry_after" ]] && [[ "$retry_after" =~ ^[0-9]+$ ]]; then
          wait_time="$retry_after"
        fi

        warn "Throttled (HTTP $_HTTP_CODE), waiting ${wait_time}s before retry ${attempt}/${MAX_RETRIES}..."
        sleep "$wait_time"

        backoff=$(( backoff * 2 ))
        (( backoff > MAX_BACKOFF )) && backoff=$MAX_BACKOFF
        ;;
      *)
        return 0
        ;;
    esac
  done
}

###############################################################################
# Resolve SharePoint site ID and drive (library) ID via Graph API
###############################################################################
resolve_site_and_drive() {
  refresh_token

  # Parse site URL → hostname + site path
  local url="$SITE_URL"
  url="${url#https://}"; url="${url#http://}"; url="${url%/}"
  local hostname="${url%%/*}"
  local site_path="/${url#*/}"
  [[ "$hostname" == "$url" ]] && site_path=""

  # Resolve site
  log "Resolving SharePoint site..."
  local endpoint
  if [[ -n "$site_path" ]]; then
    endpoint="/sites/${hostname}:${site_path}"
  else
    endpoint="/sites/${hostname}"
  fi

  graph_curl -H "Authorization: Bearer $ACCESS_TOKEN" "${GRAPH_BASE}${endpoint}"
  SITE_ID="$(echo "$_HTTP_BODY" | jq -r '.id // empty')"
  if [[ -z "$SITE_ID" ]]; then
    local detail; detail="$(_graph_error_detail)"
    die "Could not resolve site (HTTP $_HTTP_CODE): $SITE_URL — ${detail}"
  fi
  ok "Site ID: ${SITE_ID}"

  # Resolve drive (library)
  log "Resolving document library '${LIBRARY}'..."
  graph_curl -H "Authorization: Bearer $ACCESS_TOKEN" \
    "${GRAPH_BASE}/sites/${SITE_ID}/drives?\$select=id,name"
  DRIVE_ID="$(echo "$_HTTP_BODY" | jq -r --arg lib "$LIBRARY" \
    '.value[] | select(.name == $lib) | .id // empty')"
  if [[ -z "$DRIVE_ID" ]]; then
    if [[ "$_HTTP_CODE" != "200" ]]; then
      local detail; detail="$(_graph_error_detail)"
      die "Failed to list drives (HTTP $_HTTP_CODE): ${detail}"
    fi
    local available
    available="$(echo "$_HTTP_BODY" | jq -r '.value[].name // empty' 2>/dev/null | paste -sd', ' -)"
    die "Library '${LIBRARY}' not found. Available: ${available:-none}"
  fi
  ok "Drive ID: ${DRIVE_ID}"
}

###############################################################################
# Create a remote folder via Graph API
###############################################################################
create_remote_folder() {
  local folder_path="$1"
  local parent_path leaf
  parent_path="$(dirname "$folder_path")"
  leaf="$(basename "$folder_path")"

  local endpoint
  if [[ "$parent_path" == "." || -z "$parent_path" ]]; then
    endpoint="/drives/${DRIVE_ID}/root/children"
  else
    local encoded
    encoded="$(urlencode_path "$parent_path")"
    endpoint="/drives/${DRIVE_ID}/root:/${encoded}:/children"
  fi

  local body
  body="$(jq -n --arg name "$leaf" \
    '{"name": $name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}')"

  graph_curl -X POST "${GRAPH_BASE}${endpoint}" \
    -H "Authorization: Bearer $ACCESS_TOKEN" \
    -H "Content-Type: application/json" \
    -d "$body"

  case "$_HTTP_CODE" in
    200|201) return 0 ;;
    409)     return 0 ;;  # Already exists
    *)
      local detail; detail="$(_graph_error_detail)"
      err "Failed to create folder (HTTP $_HTTP_CODE): $folder_path — ${detail}"
      return 1
      ;;
  esac
}

###############################################################################
# Upload — small files (< 4 MB): simple PUT
###############################################################################
upload_small_file() {
  local local_path="$1" remote_path="$2"
  local encoded
  encoded="$(urlencode_path "$remote_path")"

  graph_curl -X PUT \
    "${GRAPH_BASE}/drives/${DRIVE_ID}/root:/${encoded}:/content" \
    -H "Authorization: Bearer $ACCESS_TOKEN" \
    -H "Content-Type: application/octet-stream" \
    --data-binary "@${local_path}"

  case "$_HTTP_CODE" in
    200|201) return 0 ;;
    *)
      local detail; detail="$(_graph_error_detail)"
      err "Upload failed (HTTP $_HTTP_CODE): $remote_path — ${detail}"
      return 1
      ;;
  esac
}

###############################################################################
# Upload — large files (≥ 4 MB): chunked upload session
###############################################################################
upload_large_file() {
  local local_path="$1" remote_path="$2"
  local file_size
  file_size="$(get_file_size "$local_path")"

  local encoded
  encoded="$(urlencode_path "$remote_path")"

  # Create upload session (with retries + exponential backoff)
  local upload_url="" session_attempt=0 session_backoff=$INITIAL_BACKOFF
  while true; do
    graph_curl -X POST \
      "${GRAPH_BASE}/drives/${DRIVE_ID}/root:/${encoded}:/createUploadSession" \
      -H "Authorization: Bearer $ACCESS_TOKEN" \
      -H "Content-Type: application/json" \
      -d '{"item":{"@microsoft.graph.conflictBehavior":"replace"}}'

    upload_url="$(echo "$_HTTP_BODY" | jq -r '.uploadUrl // empty')"
    if [[ -n "$upload_url" ]]; then
      break
    fi

    (( session_attempt++ )) || true
    if (( session_attempt > MAX_RETRIES )); then
      local detail; detail="$(_graph_error_detail)"
      err "Failed to create upload session (HTTP $_HTTP_CODE) after ${MAX_RETRIES} retries: $remote_path — ${detail}"
      return 1
    fi

    local detail; detail="$(_graph_error_detail)"
    warn "Upload session creation failed (HTTP $_HTTP_CODE, attempt ${session_attempt}/${MAX_RETRIES}): $remote_path — ${detail}"
    warn "Retrying in ${session_backoff}s..."
    sleep "$session_backoff"
    session_backoff=$(( session_backoff * 2 ))
    (( session_backoff > MAX_BACKOFF )) && session_backoff=$MAX_BACKOFF
    refresh_token
  done

  # Upload in 10 MiB chunks
  local offset=0 chunk_idx=0
  local total_chunks=$(( (file_size + CHUNK_SIZE - 1) / CHUNK_SIZE ))

  while (( offset < file_size )); do
    local remaining=$(( file_size - offset ))
    local this_chunk=$(( remaining < CHUNK_SIZE ? remaining : CHUNK_SIZE ))
    local end=$(( offset + this_chunk - 1 ))
    (( chunk_idx++ )) || true

    log "  Chunk ${chunk_idx}/${total_chunks}: bytes ${offset}-${end}/${file_size}"

    local http_code="" chunk_attempt=0 chunk_backoff=$INITIAL_BACKOFF
    while true; do
      http_code="$(dd if="$local_path" bs="$CHUNK_SIZE" skip=$(( chunk_idx - 1 )) count=1 2>/dev/null | \
        curl -sS -o /dev/null -w "%{http_code}" -X PUT "$upload_url" \
          -H "Content-Length: ${this_chunk}" \
          -H "Content-Range: bytes ${offset}-${end}/${file_size}" \
          -H "Content-Type: application/octet-stream" \
          --data-binary @-)"

      case "$http_code" in
        200|201|202) break ;;
        429|503)
          (( chunk_attempt++ )) || true
          if (( chunk_attempt > MAX_RETRIES )); then
            curl -sS -X DELETE "$upload_url" >/dev/null 2>&1 || true
            die "Chunk throttled (HTTP $http_code) after ${MAX_RETRIES} retries at offset $offset of $remote_path — aborting. Re-run to resume."
          fi
          warn "Chunk throttled (HTTP $http_code), waiting ${chunk_backoff}s before retry ${chunk_attempt}/${MAX_RETRIES}..."
          sleep "$chunk_backoff"
          chunk_backoff=$(( chunk_backoff * 2 ))
          (( chunk_backoff > MAX_BACKOFF )) && chunk_backoff=$MAX_BACKOFF
          ;;
        *)
          err "Chunk upload failed (HTTP $http_code) at offset $offset of $remote_path"
          curl -sS -X DELETE "$upload_url" >/dev/null 2>&1 || true
          return 1
          ;;
      esac
    done

    offset=$(( offset + this_chunk ))
  done

  return 0
}

###############################################################################
# Ledger helpers — path-based resume tracking
###############################################################################
ledger_lookup() {
  local rel_path="$1"
  [[ -f "$LEDGER_FILE" ]] || return 1
  grep -qFx "$rel_path" "$LEDGER_FILE" 2>/dev/null
}

ledger_record() {
  local rel_path="$1"
  echo "$rel_path" >> "$LEDGER_FILE"
}

###############################################################################
# Progress display — called after each file is processed
###############################################################################
_print_progress() {
  local files_done="$1" files_total="$2"
  local bytes_done="$3" bytes_total="$4"
  local upload_bytes="$5" upload_secs="$6"

  local files_left=$(( files_total - files_done ))
  local bytes_left=$(( bytes_total - bytes_done ))

  log "  Progress : ${files_done}/${files_total} files, $(human_size "$bytes_done")/$(human_size "$bytes_total")"
  log "  Remaining: ${files_left} files, $(human_size "$bytes_left")"

  if (( upload_secs > 0 && upload_bytes > 0 )); then
    local throughput=$(( upload_bytes / upload_secs ))
    if (( throughput > 0 && bytes_left > 0 )); then
      local eta_secs=$(( bytes_left / throughput ))
      log "  ETA      : ~$(human_time "$eta_secs")"
    elif (( bytes_left == 0 )); then
      log "  ETA      : done"
    fi
  else
    if (( bytes_left > 0 )); then
      log "  ETA      : estimating..."
    else
      log "  ETA      : done"
    fi
  fi
}

###############################################################################
# Upload workflow
###############################################################################
upload_all() {
  local total=0 uploaded=0 skipped=0 failed=0 dirs_created=0
  local base_path=""
  [[ -n "$REMOTE_PATH" ]] && base_path="$REMOTE_PATH"

  # Create --remote-path directories first
  if [[ -n "$base_path" ]] && ! $DRY_RUN; then
    refresh_token
    local parts cumulative=""
    IFS='/' read -ra parts <<< "$base_path"
    for part in "${parts[@]}"; do
      [[ -z "$part" ]] && continue
      if [[ -z "$cumulative" ]]; then
        cumulative="$part"
      else
        cumulative="${cumulative}/${part}"
      fi
      log "Ensuring base folder: ${cumulative}"
      create_remote_folder "$cumulative" && (( dirs_created++ )) || true
    done
  elif [[ -n "$base_path" ]] && $DRY_RUN; then
    local parts cumulative=""
    IFS='/' read -ra parts <<< "$base_path"
    for part in "${parts[@]}"; do
      [[ -z "$part" ]] && continue
      [[ -z "$cumulative" ]] && cumulative="$part" || cumulative="${cumulative}/${part}"
      log "[DRY-RUN] Would create base folder: ${cumulative}"
    done
  fi

  # Collect subdirectories from source, sorted shallowest-first
  local dirs=()
  while IFS= read -r -d '' dir; do
    local rel="${dir#"$SOURCE"}"
    rel="${rel#/}"
    [[ -z "$rel" ]] && continue
    dirs+=("$rel")
  done < <(find "$SOURCE" -type d -print0 | sort -z)

  # Create remote subdirectories
  for rel_dir in "${dirs[@]}"; do
    local remote_folder
    [[ -n "$base_path" ]] && remote_folder="${base_path}/${rel_dir}" || remote_folder="$rel_dir"

    if $DRY_RUN; then
      log "[DRY-RUN] Would create folder: ${remote_folder}"
    else
      refresh_token
      log "Creating folder: ${remote_folder}"
      if create_remote_folder "$remote_folder"; then
        ok "Folder ready: ${remote_folder}"
        (( dirs_created++ ))
      else
        warn "Folder may already exist: ${remote_folder}"
      fi
    fi
  done

  # Pre-scan: streaming pass to compute totals only (no arrays)
  local total=0 total_bytes=0 largest_bytes=0 largest_name=""
  while IFS= read -r -d '' file; do
    local fsize
    fsize="$(get_file_size "$file")"
    (( total++ )) || true
    (( total_bytes += fsize )) || true
    if (( fsize > largest_bytes )); then
      largest_bytes=$fsize
      largest_name="${file#"$SOURCE/"}"
    fi
  done < <(find "$SOURCE" -type f ! -name ".sp-upload-ledger" -print0 | sort -z)

  echo ""
  log "───────────────────────────────────────"
  log "Files to process  : ${total}"
  log "Total size        : $(human_size "$total_bytes")"
  if (( total > 0 )); then
    log "Largest file      : $(human_size "$largest_bytes") (${largest_name})"
  fi
  log "───────────────────────────────────────"
  echo ""

  # Upload files with progress tracking — stream directly from find
  local bytes_processed=0 files_processed=0
  local upload_bytes_done=0 upload_time_elapsed=0

  while IFS= read -r -d '' file; do
    local fsize
    fsize="$(get_file_size "$file")"
    local rel="${file#"$SOURCE/"}"

    if ledger_lookup "$rel"; then
      (( skipped++ )) || true
      log "Skipping (already uploaded): ${rel}"
      (( files_processed++ )) || true
      (( bytes_processed += fsize )) || true
      _print_progress "$files_processed" "$total" "$bytes_processed" "$total_bytes" \
        "$upload_bytes_done" "$upload_time_elapsed"
      continue
    fi

    local remote_file_path
    [[ -n "$base_path" ]] && remote_file_path="${base_path}/${rel}" || remote_file_path="$rel"

    if $DRY_RUN; then
      local method="simple"
      (( fsize >= SMALL_FILE_LIMIT )) && method="chunked"
      log "[DRY-RUN] Would upload (${method}, $(human_size "$fsize")): ${rel}"
      (( files_processed++ )) || true
      (( bytes_processed += fsize )) || true
      continue
    fi

    refresh_token
    log "Uploading: ${rel} ($(human_size "$fsize"))"

    local file_start file_end
    file_start="$(date +%s)"

    local upload_ok=false
    if (( fsize < SMALL_FILE_LIMIT )); then
      upload_small_file "$file" "$remote_file_path" && upload_ok=true
    else
      upload_large_file "$file" "$remote_file_path" && upload_ok=true
    fi

    file_end="$(date +%s)"

    if $upload_ok; then
      ok "Uploaded: ${rel}"
      ledger_record "$rel"
      (( uploaded++ )) || true
      (( upload_bytes_done += fsize )) || true
      (( upload_time_elapsed += file_end - file_start )) || true
    else
      err "Failed: ${rel}"
      (( failed++ )) || true
    fi

    (( files_processed++ )) || true
    (( bytes_processed += fsize )) || true
    _print_progress "$files_processed" "$total" "$bytes_processed" "$total_bytes" \
      "$upload_bytes_done" "$upload_time_elapsed"
  done < <(find "$SOURCE" -type f ! -name ".sp-upload-ledger" -print0 | sort -z)

  # Summary
  echo ""
  log "═══════════════════════════════════════"
  log "Upload complete"
  log "  Directories created : ${dirs_created}"
  log "  Files total         : ${total}"
  log "  Uploaded            : ${uploaded}"
  log "  Skipped (resuming)  : ${skipped}"
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

  if ! $DRY_RUN; then
    resolve_site_and_drive
  fi

  log "═══════════════════════════════════════"
  log "SharePoint Bulk Upload"
  log "  Source      : $SOURCE"
  log "  Site URL    : $SITE_URL"
  log "  Library     : $LIBRARY"
  log "  Remote path : ${REMOTE_PATH:-(root)}"
  log "  Ledger      : $LEDGER_FILE"
  log "  Chunk size  : $(human_size $CHUNK_SIZE)"
  log "  Dry-run     : $DRY_RUN"
  if [[ -n "${DRIVE_ID:-}" ]]; then
    log "  Drive ID    : $DRIVE_ID"
  fi
  log "═══════════════════════════════════════"
  echo ""

  upload_all
}

main "$@"
