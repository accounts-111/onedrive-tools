---
name: onedrive-tools
description: Use when you need to interact with OneDrive files — list files, search files, download files, get file metadata, extract file hashes, audit file permissions, create view-only sharing links, create password-protected sharing links, check sharing links, revoke sharing permissions, revoke edit links, replace edit links with view-only links, list file versions, download file versions, or view recent file activity. Trigger when the user mentions OneDrive, asks about their files in the cloud, wants sharing links, needs to audit or check document permissions, or references files stored in OneDrive.
user-invocable: true
allowed-tools: Bash(python3 *onedrive_tools.py*), Read
---

# OneDrive Tools

Manage OneDrive files via Microsoft Graph API using the CLI at:
```
/Users/rob-mac/Claude Local Files/OneDrive Tools/onedrive_tools.py
```

Authentication is automatic (cached token at `~/.onedrive_token_cache.json`). If the token is expired, the tool will print a device-code URL — tell the user to open it and sign in.

## CLI Reference

All commands use this base invocation:
```bash
python3 "/Users/rob-mac/Claude Local Files/OneDrive Tools/onedrive_tools.py" <command> [args]
```

### File Operations (always safe, read-only)

**List files and folders:**
```bash
python3 "...onedrive_tools.py" list "/path" [--no-recursive]
```
Lists all files with metadata (name, size, modified date, mimeType). Outputs CSV to output dir.

**Search for files by name or content:**
```bash
python3 "...onedrive_tools.py" search "query" ["/optional-folder-scope"]
```
Full-text search across OneDrive. Optionally scope to a folder.

**Download a file:**
```bash
python3 "...onedrive_tools.py" download "/path/to/file" ["./local-dir"]
```
Downloads file to specified local directory (default: current dir).

**Get detailed file metadata:**
```bash
python3 "...onedrive_tools.py" metadata "/path/to/file"
```
Returns full JSON metadata including size, dates, hashes, webUrl.

**Extract file hashes (SHA256, quickXorHash):**
```bash
python3 "...onedrive_tools.py" hashes "/path" [--no-recursive]
```
Exports hashes for all files to CSV. Useful for integrity verification.

### Version History (read-only)

**List all versions of a file:**
```bash
python3 "...onedrive_tools.py" versions "/path/to/file"
```
Shows version ID, modified date, size, and modified-by for each version.

**Download a specific version:**
```bash
python3 "...onedrive_tools.py" download-version "/path/to/file" "version-id" ["./local-dir"]
```
Downloads a historical version. Get version IDs from the `versions` command first.

### Activity Tracking (read-only)

**View recent file activity:**
```bash
python3 "...onedrive_tools.py" activity ["/optional-folder-path"]
```
Shows recent actions (edits, views, shares, moves) with timestamps and actors. Outputs CSV.

### Permission Auditing (read-only)

**Audit all permissions on files:**
```bash
python3 "...onedrive_tools.py" audit "/path" [--no-recursive]
```
Lists every permission on every file: view links, edit links, password-protected links, inherited permissions, owner. Outputs both CSV and markdown. The markdown format shows:
- `view | scope=anonymous` — open view link
- `view | scope=anonymous | password-protected` — password-protected view link
- `edit | scope=anonymous` — open edit link (security risk!)
- `view | scope=anonymous | inherited from /path` — inherited from parent folder
- `owner` — file owner

### Sharing Link Management (modifies permissions — confirm with user first)

**Create view-only sharing links:**
```bash
python3 "...onedrive_tools.py" links "/path" [--no-recursive] [--password "SecurePass123"]
```
Creates anonymous view-only links for all files. Idempotent — reuses existing view links. With `--password`, creates password-protected links (only reuses existing password-protected links, won't skip files that only have unprotected links). Outputs CSV and markdown with all links.

**Revoke a specific permission:**
```bash
python3 "...onedrive_tools.py" revoke "/path/to/file" "permission-id"
```
Revokes a single permission by ID. Get permission IDs from the `audit` command.

**Revoke all direct edit links in a folder:**
```bash
python3 "...onedrive_tools.py" revoke-edit-links "/path" [--no-recursive]
```
Finds and revokes all direct (non-inherited) edit links. Prompts for confirmation. Outputs CSV of revoked links.

**Replace folder-level edit link with view-only:**
```bash
python3 "...onedrive_tools.py" replace-edit-view "/folder-path"
```
For a folder with an edit sharing link: revokes the edit link and creates a view-only link. All files inheriting from that folder will switch from edit to view access.

## Output

- Output files: `/Users/rob-mac/Claude Local Files/OneDrive Tools/output/YYYY-MM-DD/`
- Filenames include the folder name: `audit_Documents_143022.csv`, `links_Legal-File_091502.md`
- CSV + markdown generated where applicable
- **Always read the output files** after running a command to present results to the user

## Key folder paths

- `/Wright - Legal File` — Main legal case folder with subfolders (CPS, Custody, Divorce, Job, etc.)
- `/test-data` — Safe testing folder (use for any testing or experimentation)
- `/Documents` — General documents
- `/` — Root of OneDrive

## Important rules

1. **Never modify permissions on production folders without explicit user confirmation.** The `links` command is safe (idempotent, view-only). The `revoke`, `revoke-edit-links`, and `replace-edit-view` commands are destructive and irreversible.
2. **Use `--no-recursive`** when you only need the top level — avoids slow API calls on large folders.
3. **Read output files** after commands complete to present results to the user — use the Read tool on the CSV or MD file path printed by the command.
4. Paths are case-insensitive and should start with `/`.
5. A file can have multiple sharing links simultaneously (e.g., one open + one password-protected).
6. Inherited permissions come from parent folder sharing — they can't be revoked directly; revoke the parent folder's permission instead.

## Example workflows

**"Get me a link to the custody report":**
1. Search: `search "custody report"`
2. Find the file path, then create link: `links "/Wright - Legal File/Custody" --no-recursive`
3. Read the output markdown to get the link

**"Who has access to my legal files?":**
1. Audit: `audit "/Wright - Legal File"`
2. Read the output markdown and summarize: which files have edit access, which are shared, any password protection

**"Lock down my legal folder — view-only, no editing":**
1. First audit: `audit "/Wright - Legal File"` to see current state
2. Revoke edit links: `revoke-edit-links "/Wright - Legal File"`
3. Or replace at folder level: `replace-edit-view "/Wright - Legal File"`

**"What changed recently?":**
1. Activity: `activity "/Wright - Legal File"`
2. Read the output CSV for timestamps and actors

**"Create password-protected links for sensitive documents":**
1. Links with password: `links "/Wright - Legal File/Custody" --password "SecurePass123"`
2. Share the links + password separately with the recipient

**"Compare file versions":**
1. List versions: `versions "/Wright - Legal File/document.pdf"`
2. Download specific version: `download-version "/Wright - Legal File/document.pdf" "1.0" "./downloads"`
3. Download current: `download "/Wright - Legal File/document.pdf" "./downloads"`
