# sharepoint-bulk-upload

Lightweight Bash wrapper around the
[Microsoft 365 CLI](https://pnp.github.io/cli-microsoft365/) to bulk-upload a
local folder tree (including large files) to a SharePoint document library.

## Features

- Mirrors local directory structure on SharePoint automatically
- Handles large files via the m365 CLI's built-in chunked upload
- **Resume/retry** — tracks uploads in a ledger file; re-running skips unchanged
  files and retries failures
- **Dry-run mode** — preview what would happen without touching SharePoint
- Zero dependencies beyond Bash and the m365 CLI

## Prerequisites

| Requirement | Install |
|---|---|
| Bash 3.2+ | Pre-installed on macOS/Linux |
| Microsoft 365 CLI | `npm i -g @pnp/cli-microsoft365` |
| Active m365 session | `m365 login` |

## Usage

```bash
./sp-upload.sh \
  --source ./my-local-folder \
  --site-url https://contoso.sharepoint.com/sites/team-x \
  --library "Shared Documents" \
  [--remote-path "Optional/Sub/Path"] \
  [--dry-run] \
  [--log upload.log] \
  [--ledger .my-ledger]
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
| `m365 CLI not found` | Install with `npm i -g @pnp/cli-microsoft365` |
| `Not logged in` | Run `m365 login` and follow the prompts |
| Upload fails for large files | Ensure you have a stable network connection; the m365 CLI chunks files > 250 MB automatically |
| Permission denied | Verify your account has `AllSites.Write` permission on the target site |

## License

MIT