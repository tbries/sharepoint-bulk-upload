#!/usr/bin/env bash
# generate-test-data.sh — Fill test-data/ with randomly sized files in varied folder structures
# Usage: ./generate-test-data.sh <total_size>
#   <total_size>  Target aggregate size with unit, e.g. "1TB", "500 GB", "100GB"
#                 Supported units: MB, GB, TB
#
# Files range from 50 MB to 20 GB each, with realistic names and nested directories.
#
# Examples:
#   ./generate-test-data.sh 1TB
#   ./generate-test-data.sh "500 GB"
#   ./generate-test-data.sh 100GB

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TARGET_DIR="${SCRIPT_DIR}/test-data"

# ---------------------------------------------------------------------------
# Size parsing (reuses pattern from generate-test-file.sh)
# ---------------------------------------------------------------------------
parse_size() {
  local input
  input="$(echo "$1" | tr -d ' ' | tr '[:lower:]' '[:upper:]')"

  if [[ "$input" =~ ^([0-9]+\.?[0-9]*)([KMGT]?B)$ ]]; then
    local number="${BASH_REMATCH[1]}"
    local unit="${BASH_REMATCH[2]}"
  else
    echo "Error: Invalid size format '$1'. Use a number followed by MB, GB, or TB." >&2
    exit 1
  fi

  local multiplier
  case "$unit" in
    B)  multiplier=1 ;;
    KB) multiplier=1024 ;;
    MB) multiplier=$((1024 * 1024)) ;;
    GB) multiplier=$((1024 * 1024 * 1024)) ;;
    TB) multiplier=$((1024 * 1024 * 1024 * 1024)) ;;
  esac

  awk "BEGIN { printf \"%.0f\", $number * $multiplier }"
}

# ---------------------------------------------------------------------------
# Human-readable size
# ---------------------------------------------------------------------------
human_size() {
  local bytes=$1
  awk "BEGIN {
    b = $bytes
    if      (b >= 1024^4) printf \"%.2f TB\", b/1024^4
    else if (b >= 1024^3) printf \"%.2f GB\", b/1024^3
    else if (b >= 1024^2) printf \"%.2f MB\", b/1024^2
    else if (b >= 1024)   printf \"%.2f KB\", b/1024
    else                  printf \"%d B\", b
  }"
}

# ---------------------------------------------------------------------------
# Random integer in [min, max] (bytes). Uses $RANDOM (0-32767).
# For large ranges we combine multiple $RANDOM calls.
# ---------------------------------------------------------------------------
rand_between() {
  local min=$1 max=$2
  local range=$((max - min + 1))
  # Combine two $RANDOM values for ~30 bits of entropy, use modulo
  local r=$(( (RANDOM * 32768 + RANDOM) % range + min ))
  echo "$r"
}

# ---------------------------------------------------------------------------
# Configuration: folder structures and file-name components
# ---------------------------------------------------------------------------
FOLDERS=(
  "reports/quarterly"
  "reports/annual"
  "reports/monthly/2025"
  "reports/monthly/2026"
  "backups/database"
  "backups/file-server"
  "backups/vm-snapshots"
  "backups/config"
  "projects/alpha/data"
  "projects/alpha/exports"
  "projects/beta/raw"
  "projects/beta/processed"
  "projects/gamma/archives"
  "projects/gamma/deliverables"
  "media/video/raw-footage"
  "media/video/transcoded"
  "media/images/hi-res"
  "media/audio/recordings"
  "archives/2023"
  "archives/2024"
  "archives/2025"
  "archives/2026"
  "logs/application"
  "logs/audit"
  "logs/access"
  "datasets/training"
  "datasets/validation"
  "datasets/raw"
  "datasets/exports"
  "compliance/legal"
  "compliance/financial"
  "docs/contracts"
  "docs/specifications"
  "images/scans"
  "images/screenshots"
  "large_files"
  "uploads/pending"
  "uploads/processed"
  "staging/ingest"
  "staging/transform"
)

PREFIXES=(
  "report" "backup" "export" "snapshot" "dump" "dataset" "archive"
  "recording" "scan" "capture" "extract" "migration" "audit-log"
  "transaction-log" "telemetry" "metrics" "payload" "bundle"
  "package" "release" "build-artifact" "vm-image" "disk-image"
  "footage" "render" "output" "results" "analysis" "model"
  "training-data" "checkpoint" "inventory" "ledger" "manifest"
)

EXTENSIONS=(
  ".bin" ".dat" ".bak" ".log" ".csv" ".zip" ".tar.gz" ".iso"
  ".img" ".dump" ".sql" ".parquet" ".avro" ".jsonl" ".xml"
  ".xlsx" ".pdf" ".mp4" ".mov" ".wav" ".raw" ".tiff" ".psd"
  ".vmdk" ".vhd" ".qcow2" ".ova" ".dmg" ".pkg" ".deb"
)

# Minimum 50 MB, maximum 20 GB (in bytes)
MIN_FILE_SIZE=$((50 * 1024 * 1024))
MAX_FILE_SIZE=$((20 * 1024 * 1024 * 1024))

# ---------------------------------------------------------------------------
# Generate a single file using dd (fast, /dev/urandom)
# ---------------------------------------------------------------------------
generate_file() {
  local filepath=$1
  local size_bytes=$2
  local block_size=$((1024 * 1024))  # 1 MB blocks
  local full_blocks=$((size_bytes / block_size))
  local remainder=$((size_bytes % block_size))

  mkdir -p "$(dirname "$filepath")"

  if [[ $full_blocks -gt 0 ]]; then
    dd if=/dev/urandom of="$filepath" bs=$block_size count=$full_blocks status=none 2>/dev/null
  fi

  if [[ $remainder -gt 0 ]]; then
    if [[ $full_blocks -gt 0 ]]; then
      dd if=/dev/urandom of="$filepath" bs=1 count=$remainder oflag=append conv=notrunc status=none 2>/dev/null
    else
      dd if=/dev/urandom of="$filepath" bs=1 count=$remainder status=none 2>/dev/null
    fi
  fi
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
usage() {
  echo "Usage: $(basename "$0") <total_size>"
  echo "  <total_size>  Target total size (e.g. \"1TB\", \"500 GB\", \"100GB\")"
  echo ""
  echo "Generates random files (50 MB – 20 GB each) in test-data/ until the"
  echo "target total size is reached."
  exit 1
}

[[ $# -lt 1 ]] && usage

TOTAL_TARGET=$(parse_size "$1")
echo "=== Generate Test Data ==="
echo "Target total size : $(human_size "$TOTAL_TARGET")"
echo "Output directory  : ${TARGET_DIR}"
echo ""

mkdir -p "$TARGET_DIR"

generated_bytes=0
file_count=0

while (( generated_bytes < TOTAL_TARGET )); do
  remaining=$((TOTAL_TARGET - generated_bytes))

  # Stop if remaining is less than minimum file size
  if (( remaining < MIN_FILE_SIZE )); then
    break
  fi

  # Cap max for this file at remaining bytes or MAX_FILE_SIZE
  cap=$MAX_FILE_SIZE
  if (( remaining < cap )); then
    cap=$remaining
  fi

  # Pick a random size between MIN and cap.
  # For sizes larger than ~1 GB we pick in MB granularity to keep it simple.
  min_mb=$((MIN_FILE_SIZE / (1024 * 1024)))         # 50
  cap_mb=$((cap / (1024 * 1024)))                    # up to 20480
  file_size_mb=$(rand_between "$min_mb" "$cap_mb")
  file_size_bytes=$((file_size_mb * 1024 * 1024))

  # Pick random folder and name
  folder=${FOLDERS[$(rand_between 0 $((${#FOLDERS[@]} - 1)))]}
  prefix=${PREFIXES[$(rand_between 0 $((${#PREFIXES[@]} - 1)))]}
  ext=${EXTENSIONS[$(rand_between 0 $((${#EXTENSIONS[@]} - 1)))]}

  # Build a unique, realistic filename
  timestamp="$(date +%Y%m%d)-$(printf '%04d' $file_count)"
  filename="${prefix}_${timestamp}${ext}"
  filepath="${TARGET_DIR}/${folder}/${filename}"

  echo "[$(( file_count + 1 ))] $(human_size $file_size_bytes)  ${folder}/${filename}"

  generate_file "$filepath" "$file_size_bytes"

  generated_bytes=$((generated_bytes + file_size_bytes))
  file_count=$((file_count + 1))

  # Progress update every 10 files
  if (( file_count % 10 == 0 )); then
    pct=$(awk "BEGIN { printf \"%.1f\", ($generated_bytes / $TOTAL_TARGET) * 100 }")
    echo "--- Progress: $(human_size $generated_bytes) / $(human_size "$TOTAL_TARGET") (${pct}%) — ${file_count} files ---"
  fi
done

echo ""
echo "=== Complete ==="
echo "Files created : ${file_count}"
echo "Total size    : $(human_size $generated_bytes)"
echo "Location      : ${TARGET_DIR}"
