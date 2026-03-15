# OneDrive Tools

Command-line and interactive tool for managing OneDrive files via Microsoft Graph API. Designed for auditing permissions, creating sharing links, tracking file versions, and managing access control.

## Features

- **List & search** files across OneDrive folders
- **Download** files and specific file versions
- **Audit permissions** — see every sharing link, who has access, password protection, inherited permissions
- **Create view-only links** — with optional password protection
- **Revoke permissions** — remove edit links, replace edit with view-only
- **Version history** — list and download previous versions
- **Activity tracking** — see recent file actions (edits, views, shares)
- **File hashes** — extract SHA256/quickXor hashes for integrity verification

## Setup

### 1. Install Python dependencies

```bash
pip3 install -r requirements.txt
```

### 2. Run the tool

**Interactive mode** (double-click or run without arguments):
```bash
python3 onedrive_tools.py
```
Or double-click `OneDrive Tools.command` on macOS.

**CLI mode** (for scripts and Claude Code skill):
```bash
python3 onedrive_tools.py <command> [args]
```

### 3. Authenticate

On first run, the tool uses device-code flow:
1. A URL and code are displayed in the terminal
2. Open the URL in a browser and enter the code
3. Sign in with your Microsoft account
4. The token is cached at `~/.onedrive_token_cache.json` for future runs

No Azure app registration is required — the tool uses Microsoft's public Graph Command Line Tools client ID.

## CLI Commands

| Command | Usage | Description |
|---------|-------|-------------|
| `list` | `list "/path" [--no-recursive]` | List files with metadata |
| `search` | `search "query" ["/path"]` | Full-text search |
| `download` | `download "/path/to/file" ["./dir"]` | Download a file |
| `metadata` | `metadata "/path/to/file"` | Get file metadata (JSON) |
| `hashes` | `hashes "/path" [--no-recursive]` | Extract file hashes to CSV |
| `audit` | `audit "/path" [--no-recursive]` | Audit all permissions |
| `links` | `links "/path" [--no-recursive] [--password "pw"]` | Create view-only sharing links |
| `revoke` | `revoke "/path/to/file" "perm_id"` | Revoke a specific permission |
| `revoke-edit-links` | `revoke-edit-links "/path" [--no-recursive]` | Revoke all direct edit links |
| `replace-edit-view` | `replace-edit-view "/folder"` | Replace edit link with view-only |
| `versions` | `versions "/path/to/file"` | List file versions |
| `download-version` | `download-version "/path" "ver_id" ["./dir"]` | Download specific version |
| `activity` | `activity ["/path"]` | Get recent activity |

## Examples

```bash
# List all files in a folder
python3 onedrive_tools.py list "/Documents"

# Audit permissions on legal files
python3 onedrive_tools.py audit "/Legal Files"

# Create password-protected sharing links
python3 onedrive_tools.py links "/Legal Files" --password "SecurePass123"

# Search for a document
python3 onedrive_tools.py search "custody report"

# Check file version history
python3 onedrive_tools.py versions "/Documents/contract.pdf"

# Download a specific version
python3 onedrive_tools.py download-version "/Documents/contract.pdf" "1.0" ./downloads

# Revoke all edit links in a folder
python3 onedrive_tools.py revoke-edit-links "/Shared Documents"

# View recent activity
python3 onedrive_tools.py activity "/Documents"
```

## Output

All output files are saved to `./output/YYYY-MM-DD/` with descriptive filenames:
- `audit_Documents_143022.csv` / `.md` — permission audit
- `links_Legal-Files_091502.csv` / `.md` — sharing links
- `files_Documents_120530.csv` — file listings
- `hashes_Documents_143022.csv` — file hashes
- `activity_Documents_143022.csv` — activity log

## Claude Code Skill

A Claude Code skill is included at `skill/SKILL.md`. To install it globally:

```bash
mkdir -p ~/.claude/skills/onedrive-tools
cp skill/SKILL.md ~/.claude/skills/onedrive-tools/SKILL.md
```

This lets Claude automatically use OneDrive Tools when you mention OneDrive files, sharing links, or permissions in any conversation. You can also invoke it manually with `/onedrive-tools`.

## Authentication Details

- **Client ID:** Microsoft Graph Command Line Tools (public client — no registration needed)
- **Authority:** `https://login.microsoftonline.com/consumers` (personal Microsoft accounts)
- **Scopes:** `Files.ReadWrite.All`
- **Token cache:** `~/.onedrive_token_cache.json`

## Requirements

- Python 3.7+
- macOS, Linux, or Windows
- Personal Microsoft account with OneDrive
