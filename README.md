# sharepoint-bulk-upload

Lightweight Bash wrapper that bulk-uploads a local folder tree (including large
files) to a SharePoint document library using the
[Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview).

**No app registration required** — authenticates via the Azure CLI using your
own identity (browser sign-in, MFA supported).

## Features

- Mirrors local directory structure on SharePoint automatically
- **Large-file support** — files ≥ 4 MB use chunked upload sessions (10 MiB chunks)
- **Resume/retry** — SHA-256 ledger tracks uploads; re-running skips unchanged
  files and retries failures
- **Dry-run mode** — preview what would happen without touching SharePoint
- Zero dependencies beyond Bash, Azure CLI, curl, and jq

## Prerequisites

| Requirement | Install |
|---|---|
| Bash 3.2+ | Pre-installed on macOS/Linux |
| Azure CLI | `brew install azure-cli` or [aka.ms/installazurecli](https://aka.ms/installazurecli) |
| jq | `brew install jq` or `apt install jq` |
| curl | Pre-installed on macOS/Linux |

## Getting Started

```bash
# 1. Sign in (opens browser, supports MFA)
az login

# 2. Run the upload
./sp-upload.sh \
  --source ./my-local-folder \
  --site-url https://contoso.sharepoint.com/sites/team-x \
  --library "Shared Documents"
```

## Usage

```bash
./sp-upload.sh \
  --source <path> \
  --site-url <url> \
  --library <name> \
  [--remote-path <path>] \
  [--dry-run] \
  [--log <path>] \
  [--ledger <path>]
```

### Options

| Flag | Required | Description |
|------|----------|-------------|
| `--source <path>` | ✅ | Local folder to upload |
| `--site-url <url>` | ✅ | SharePoint site URL |
| `--library <name>` | ✅ | Target document library (e.g. `Shared Documents`) |
| `--remote-path <path>` | | Sub-path inside the library (default: root) |
| `--dry-run` | | Preview operations without executing |
| `--log <path>` | | Log file path (default: `./sp-upload.log`) |
| `--ledger <path>` | | Ledger file for resume tracking (default: `<source>/.sp-upload-ledger`) |
| `-h`, `--help` | | Show help |

### Examples

Upload a project folder to the default library:

```bash
./sp-upload.sh \
  --source ./project-files \
  --site-url https://contoso.sharepoint.com/sites/engineering \
  --library "Shared Documents"
```

Upload into a sub-folder with dry-run:

```bash
./sp-upload.sh \
  --source ./reports \
  --site-url https://contoso.sharepoint.com/sites/finance \
  --library "Documents" \
  --remote-path "2026/Q1" \
  --dry-run
```

## How It Works

1. Authenticates via `az account get-access-token` (delegated user identity)
2. Resolves the SharePoint site ID and document library (drive) ID via Graph API
3. Walks the local folder tree and mirrors it on SharePoint
4. Uploads each file:
   - **< 4 MB**: simple `PUT` to Graph API
   - **≥ 4 MB**: creates an upload session, then uploads in 10 MiB chunks
5. Records each successful upload in a ledger file

## Resume / Retry Behavior

The script maintains a **ledger file** (`.sp-upload-ledger`) that records the
relative path and SHA-256 hash of each successfully uploaded file.

On re-run:
- **Unchanged files** are skipped (matching path + hash)
- **Changed files** are re-uploaded (path exists but hash differs)
- **Failed files** from a previous run are retried (not in ledger)

Delete the ledger file to force a full re-upload.

## Troubleshooting

| Problem | Solution |
|---|---|
| `Azure CLI not found` | Install with `brew install azure-cli` |
| `Not logged in` | Run `az login` — opens browser for sign-in |
| `jq not found` | Install with `brew install jq` |
| `Library not found` | Check the exact library name (case-sensitive). The error message lists available libraries. |
| Large file upload fails | Ensure stable network; the script uses 10 MiB chunks and cancels the session on failure — re-run to retry |
| Permission denied (403) | Your account needs write access to the target document library |
| Token expired during upload | The script refreshes tokens automatically every 45 minutes |

## License

MIT