#!/usr/bin/env bash
# generate-test-file.sh — Generate a file of a specified target size
# Usage: ./generate-test-file.sh <size> [output_file]
#   <size>       Target size with unit, e.g. "500 MB", "2GB", "10 KB"
#                Supported units: B, KB, MB, GB, TB
#   [output_file] Optional output path (default: testfile-<size>.bin)
#
# Examples:
#   ./generate-test-file.sh "10 GB"
#   ./generate-test-file.sh 500MB my-large-file.bin
#   ./generate-test-file.sh "1.5 GB" uploads/payload.bin

set -euo pipefail

usage() {
  echo "Usage: $(basename "$0") <size> [output_file]"
  echo "  <size>  Target size with unit (B, KB, MB, GB, TB). Examples: \"10 GB\", 500MB, \"1.5 GB\""
  echo "  [output_file]  Optional output path (default: testfile-<normalized>.bin)"
  exit 1
}

parse_size() {
  local input
  input="$(echo "$1" | tr -d ' ' | tr '[:lower:]' '[:upper:]')"

  if [[ "$input" =~ ^([0-9]+\.?[0-9]*)([KMGT]?B)$ ]]; then
    local number="${BASH_REMATCH[1]}"
    local unit="${BASH_REMATCH[2]}"
  else
    echo "Error: Invalid size format '$1'. Use a number followed by B, KB, MB, GB, or TB." >&2
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

  # Use awk for floating-point arithmetic and emit an integer byte count
  TOTAL_BYTES=$(awk "BEGIN { printf \"%.0f\", $number * $multiplier }")
}

[[ $# -lt 1 ]] && usage

parse_size "$1"

# Default output filename derived from the size argument
if [[ $# -ge 2 ]]; then
  OUTPUT="$2"
else
  sanitised=$(echo "$1" | tr -d ' ' | tr '[:upper:]' '[:lower:]')
  OUTPUT="testfile-${sanitised}.bin"
fi

# Pick a block size and count to keep dd efficient
BLOCK_SIZE=$((1024 * 1024))  # 1 MB blocks
FULL_BLOCKS=$((TOTAL_BYTES / BLOCK_SIZE))
REMAINDER=$((TOTAL_BYTES % BLOCK_SIZE))

echo "Generating ${OUTPUT} (${TOTAL_BYTES} bytes) …"

# Write full 1 MB blocks
if [[ $FULL_BLOCKS -gt 0 ]]; then
  dd if=/dev/urandom of="$OUTPUT" bs=$BLOCK_SIZE count=$FULL_BLOCKS status=progress 2>&1
fi

# Append any remaining bytes
if [[ $REMAINDER -gt 0 ]]; then
  if [[ $FULL_BLOCKS -gt 0 ]]; then
    dd if=/dev/urandom of="$OUTPUT" bs=1 count=$REMAINDER oflag=append conv=notrunc status=none 2>&1
  else
    dd if=/dev/urandom of="$OUTPUT" bs=1 count=$REMAINDER status=none 2>&1
  fi
fi

# Verify
ACTUAL=$(wc -c < "$OUTPUT" | tr -d ' ')
echo "Done. ${OUTPUT} is ${ACTUAL} bytes."

if [[ "$ACTUAL" -ne "$TOTAL_BYTES" ]]; then
  echo "Warning: expected ${TOTAL_BYTES} bytes but got ${ACTUAL}." >&2
  exit 1
fi
