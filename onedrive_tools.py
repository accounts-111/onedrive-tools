#!/usr/bin/env python3
"""
OneDrive Tools — comprehensive OneDrive management library via Microsoft Graph API.

Provides file listing, search, download, hash extraction, permission auditing,
sharing link management, version history, and activity tracking.

Uses Microsoft Graph Command Line Tools client ID (device code flow).
No Azure app registration needed.

Usage:
    python3 onedrive_tools.py <command> [args]

Commands:
    list "/path"                          List files with metadata
    search "query" ["/path"]              Full-text search
    download "/path/to/file" ["./dir"]    Download a file
    metadata "/path/to/file"              Get file metadata
    hashes "/path"                        Extract file hashes to CSV
    audit "/path"                         Audit permissions
    links "/path"                         Create view-only links
    revoke "/path/to/file" "perm_id"      Revoke a specific permission
    revoke-edit-links "/path"             Revoke all direct edit links
    replace-edit-view "/folder"           Replace edit link with view link
    versions "/path/to/file"              List file versions
    download-version "/path" "ver_id" ["./dir"]  Download specific version
    activity ["/path"]                    Get recent activity
"""

import argparse
import csv
import json
import os
import sys
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path

import msal
import requests

CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
AUTHORITY = "https://login.microsoftonline.com/consumers"
SCOPES = ["Files.ReadWrite.All"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TOKEN_CACHE_PATH = Path.home() / ".onedrive_token_cache.json"
SCRIPT_DIR = Path(__file__).parent
MAX_WORKERS = 10


class OneDriveTools:
    """Comprehensive OneDrive management via Microsoft Graph API."""

    def __init__(self, output_dir=None):
        self.token = None
        self._cache = None
        self._app = None
        # Output goes to: <base>/output/YYYY-MM-DD/
        if output_dir:
            self.output_dir = Path(output_dir)
        else:
            today = date.today().isoformat()  # e.g. 2026-03-14
            self.output_dir = SCRIPT_DIR / "output" / today
        self.output_dir.mkdir(parents=True, exist_ok=True)

    # ─────────────────────────────────────────────
    # Auth
    # ─────────────────────────────────────────────

    def authenticate(self):
        """Authenticate via device code flow with token caching.

        Returns the access token string. Sets self.token for use by all methods.
        """
        self._cache = msal.SerializableTokenCache()
        if TOKEN_CACHE_PATH.exists():
            self._cache.deserialize(TOKEN_CACHE_PATH.read_text())

        self._app = msal.PublicClientApplication(
            CLIENT_ID, authority=AUTHORITY, token_cache=self._cache
        )

        # Try cached token first
        accounts = self._app.get_accounts()
        if accounts:
            result = self._app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                print("[auth] Authenticated using cached token")
                self._save_cache()
                self.token = result["access_token"]
                return self.token

        # Device code flow
        flow = self._app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise RuntimeError(
                f"Error initiating device flow: {flow.get('error_description', 'Unknown error')}"
            )

        print(f"\n{flow['message']}\n")
        result = self._app.acquire_token_by_device_flow(flow)

        if "access_token" not in result:
            raise RuntimeError(
                f"Authentication failed: {result.get('error_description', 'Unknown error')}"
            )

        print("[auth] Authenticated successfully")
        self._save_cache()
        self.token = result["access_token"]
        return self.token

    def _save_cache(self):
        if self._cache and self._cache.has_state_changed:
            TOKEN_CACHE_PATH.write_text(self._cache.serialize())

    def _require_token(self):
        if not self.token:
            raise RuntimeError("Not authenticated. Call authenticate() first.")

    def _output_path(self, prefix, ext="csv", folder_path=None):
        """Generate a timestamped output file path with optional folder context.

        e.g. output/2026-03-14/audit_test-data_153042.csv
             output/2026-03-14/files_Exhibit-C-CPS_153042.csv
        """
        timestamp = time.strftime("%H%M%S")
        if folder_path and folder_path != "/":
            # Use the last folder name, sanitized for filenames
            folder_name = folder_path.strip("/").split("/")[-1]
            folder_name = folder_name.replace(" ", "-")
            # Remove any characters not safe for filenames
            folder_name = "".join(c for c in folder_name if c.isalnum() or c in "-_.")
            if len(folder_name) > 40:
                folder_name = folder_name[:40]
            return self.output_dir / f"{prefix}_{folder_name}_{timestamp}.{ext}"
        return self.output_dir / f"{prefix}_{timestamp}.{ext}"

    # ─────────────────────────────────────────────
    # HTTP helpers (with retry and rate limiting)
    # ─────────────────────────────────────────────

    _last_request_time = 0
    _min_request_interval = 0.05  # 50ms between requests (max ~20/sec)

    def _headers(self):
        return {"Authorization": f"Bearer {self.token}"}

    def _throttle(self):
        """Enforce minimum interval between API requests."""
        now = time.time()
        elapsed = now - self._last_request_time
        if elapsed < self._min_request_interval:
            time.sleep(self._min_request_interval - elapsed)
        self._last_request_time = time.time()

    def _request_with_retry(self, method, url, max_retries=5, **kwargs):
        """Make an HTTP request with exponential backoff retry.

        Retries on 429 (throttled), 503 (service unavailable), 504 (gateway timeout),
        and connection errors.
        """
        for attempt in range(max_retries + 1):
            self._throttle()
            try:
                resp = method(url, **kwargs)
                if resp.status_code == 429:
                    retry_after = int(resp.headers.get("Retry-After", 2 ** attempt))
                    print(f"  [throttled] Retry-After {retry_after}s (attempt {attempt + 1}/{max_retries + 1})", flush=True)
                    time.sleep(retry_after)
                    continue
                if resp.status_code in (503, 504) and attempt < max_retries:
                    wait = 2 ** attempt
                    print(f"  [retry] {resp.status_code} — waiting {wait}s (attempt {attempt + 1}/{max_retries + 1})", flush=True)
                    time.sleep(wait)
                    continue
                resp.raise_for_status()
                return resp
            except (requests.exceptions.ConnectionError, ConnectionResetError) as e:
                if attempt < max_retries:
                    wait = 2 ** attempt
                    print(f"  [connection error] {e.__class__.__name__} — waiting {wait}s (attempt {attempt + 1}/{max_retries + 1})", flush=True)
                    time.sleep(wait)
                    continue
                raise
        resp.raise_for_status()
        return resp

    def _get(self, url, params=None):
        resp = self._request_with_retry(requests.get, url, headers=self._headers(), params=params)
        return resp.json()

    def _post(self, url, json_body):
        headers = {**self._headers(), "Content-Type": "application/json"}
        resp = self._request_with_retry(requests.post, url, headers=headers, json=json_body)
        return resp.json()

    def _delete(self, url):
        self._request_with_retry(requests.delete, url, headers=self._headers())

    def _item_url(self, file_path):
        """Build Graph API URL for an item by path."""
        path = file_path.strip("/")
        if not path:
            return f"{GRAPH_BASE}/me/drive/root"
        return f"{GRAPH_BASE}/me/drive/root:/{path}"

    def _children_url(self, folder_path):
        """Build Graph API URL for children of a folder by path."""
        path = folder_path.strip("/") if folder_path else ""
        if not path:
            return f"{GRAPH_BASE}/me/drive/root/children"
        return f"{GRAPH_BASE}/me/drive/root:/{path}:/children"

    # ─────────────────────────────────────────────
    # Progress helper
    # ─────────────────────────────────────────────

    @staticmethod
    def _progress(completed, total, start_time, every=50):
        """Print progress line every N items."""
        if completed % every == 0 or completed == total:
            elapsed = time.time() - start_time
            rate = completed / elapsed if elapsed > 0 else 0
            eta = (total - completed) / rate if rate > 0 else 0
            print(
                f"  [{completed}/{total}] {rate:.1f} items/sec, ~{eta:.0f}s remaining",
                flush=True,
            )

    # ─────────────────────────────────────────────
    # File Operations
    # ─────────────────────────────────────────────

    def list_folder_contents(self, folder_path):
        """List immediate children (folders and files) of a folder.

        Returns:
            Tuple of (folders, files) where each is a list of dicts.
            Folders have: name, path, childCount.
            Files have: name, path, size, lastModified.
        """
        self._require_token()
        url = self._children_url(folder_path)
        folders = []
        files = []

        while url:
            data = self._get(url)
            for item in data.get("value", []):
                item_path = f"{(folder_path or '/').rstrip('/')}/{item['name']}"
                if "folder" in item:
                    folders.append({
                        "name": item["name"],
                        "path": item_path,
                        "childCount": item.get("folder", {}).get("childCount", 0),
                    })
                elif "file" in item:
                    files.append({
                        "name": item["name"],
                        "path": item_path,
                        "size": item.get("size", 0),
                        "lastModified": item.get("lastModifiedDateTime", ""),
                    })
            url = data.get("@odata.nextLink")

        return folders, files

    def list_files(self, folder_path, recursive=True):
        """List files with metadata in a OneDrive folder.

        Uses delta query for recursive listing (much faster — single paginated
        request instead of one API call per subfolder). Falls back to children
        API for non-recursive listing.

        Args:
            folder_path: OneDrive path (e.g. "/Documents"). Use "/" or None for root.
            recursive: If True, recurse into subfolders.

        Returns:
            List of dicts with keys: id, name, path, folder, size, lastModified,
            mimeType, webUrl.
        """
        self._require_token()
        folder_path = folder_path or "/"

        if not recursive:
            # Non-recursive: use children API (single folder)
            url = self._children_url(folder_path)
            all_files = []
            self._list_children(url, folder_path, all_files)
            return all_files

        # Recursive: use delta query for efficiency
        return self._list_via_delta(folder_path)

    def _list_via_delta(self, folder_path):
        """List all files recursively using delta query.

        Delta query returns all items in the drive in a flat list,
        paginated. We filter to the target folder prefix.
        """
        folder_path = folder_path or "/"
        folder_prefix = folder_path.rstrip("/") if folder_path != "/" else ""

        # Get the folder's item ID to scope the delta query
        if folder_path == "/":
            delta_url = f"{GRAPH_BASE}/me/drive/root/delta"
        else:
            path = folder_path.strip("/")
            item = self._get(f"{GRAPH_BASE}/me/drive/root:/{path}")
            item_id = item["id"]
            delta_url = f"{GRAPH_BASE}/me/drive/items/{item_id}/delta"

        all_files = []
        page = 0
        while delta_url:
            page += 1
            data = self._get(delta_url)
            for item in data.get("value", []):
                # Skip folders and deleted items
                if "folder" in item or item.get("deleted"):
                    continue
                if "file" not in item:
                    continue

                # Build path from parentReference
                parent = item.get("parentReference", {})
                parent_path = parent.get("path", "")
                # Strip /drive/root: prefix
                if ":/drive/root:" in parent_path:
                    parent_path = parent_path.split(":/drive/root:")[-1]
                elif parent_path.startswith("/drive/root:"):
                    parent_path = parent_path[len("/drive/root:"):]
                elif parent_path.startswith("/drive/root"):
                    parent_path = parent_path[len("/drive/root"):]
                if not parent_path:
                    parent_path = "/"

                item_path = f"{parent_path.rstrip('/')}/{item['name']}"

                all_files.append({
                    "id": item["id"],
                    "name": item["name"],
                    "path": item_path,
                    "folder": parent_path,
                    "size": item.get("size", 0),
                    "lastModified": item.get("lastModifiedDateTime", ""),
                    "mimeType": item.get("file", {}).get("mimeType", ""),
                    "webUrl": item.get("webUrl", ""),
                })

            if len(all_files) > 0 and page % 2 == 0:
                print(f"  ... {len(all_files)} files found ({page} pages fetched)", flush=True)

            # Follow pagination (nextLink), stop at deltaLink
            delta_url = data.get("@odata.nextLink")
            if not delta_url:
                break  # deltaLink means we've got everything

        print(f"  ... {len(all_files)} files total ({page} pages)", flush=True)
        return all_files

    def _list_children(self, url, current_path, all_files):
        """List immediate children (files only) with pagination."""
        while url:
            data = self._get(url)
            for item in data.get("value", []):
                if "file" in item:
                    item_path = f"{current_path.rstrip('/')}/{item['name']}"
                    all_files.append({
                        "id": item["id"],
                        "name": item["name"],
                        "path": item_path,
                        "folder": os.path.dirname(item_path),
                        "size": item.get("size", 0),
                        "lastModified": item.get("lastModifiedDateTime", ""),
                        "mimeType": item.get("file", {}).get("mimeType", ""),
                        "webUrl": item.get("webUrl", ""),
                    })
            url = data.get("@odata.nextLink")

    def search(self, query, folder_path=None):
        """Full-text search via the Graph API search endpoint.

        Args:
            query: Search query string.
            folder_path: Optional folder to scope the search. None = entire drive.

        Returns:
            List of dicts with keys: id, name, path, folder, size, lastModified,
            mimeType, webUrl.
        """
        self._require_token()
        if folder_path and folder_path != "/":
            path = folder_path.strip("/")
            url = f"{GRAPH_BASE}/me/drive/root:/{path}:/search(q='{query}')"
        else:
            url = f"{GRAPH_BASE}/me/drive/root/search(q='{query}')"

        results = []
        while url:
            data = self._get(url)
            for item in data.get("value", []):
                parent = item.get("parentReference", {})
                parent_path = parent.get("path", "")
                # Strip the /drive/root: prefix
                if ":/drive/root:" in parent_path:
                    parent_path = parent_path.split(":/drive/root:")[-1]
                elif parent_path.startswith("/drive/root:"):
                    parent_path = parent_path[len("/drive/root:"):]
                elif parent_path.startswith("/drive/root"):
                    parent_path = parent_path[len("/drive/root"):]
                if not parent_path:
                    parent_path = "/"

                item_path = f"{parent_path.rstrip('/')}/{item['name']}"
                results.append({
                    "id": item["id"],
                    "name": item["name"],
                    "path": item_path,
                    "folder": parent_path,
                    "size": item.get("size", 0),
                    "lastModified": item.get("lastModifiedDateTime", ""),
                    "mimeType": item.get("file", {}).get("mimeType", ""),
                    "webUrl": item.get("webUrl", ""),
                })
            url = data.get("@odata.nextLink")

        print(f"[search] Found {len(results)} results for '{query}'")
        return results

    def download(self, file_path, local_dir="."):
        """Download a file from OneDrive to local disk.

        Args:
            file_path: OneDrive path (e.g. "/Documents/report.pdf").
            local_dir: Local directory to save the file. Defaults to current dir.

        Returns:
            Dict with keys: local_path, size, name.
        """
        self._require_token()
        path = file_path.strip("/")
        url = f"{GRAPH_BASE}/me/drive/root:/{path}:/content"

        resp = requests.get(url, headers=self._headers(), stream=True, allow_redirects=True)
        resp.raise_for_status()

        local_dir = Path(local_dir)
        local_dir.mkdir(parents=True, exist_ok=True)
        filename = os.path.basename(path)
        local_path = local_dir / filename

        total = int(resp.headers.get("Content-Length", 0))
        downloaded = 0
        start = time.time()

        with open(local_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total > 0:
                        pct = downloaded / total * 100
                        elapsed = time.time() - start
                        speed = downloaded / elapsed / 1024 / 1024 if elapsed > 0 else 0
                        print(
                            f"\r  Downloading {filename}: {pct:.1f}% ({speed:.1f} MB/s)",
                            end="",
                            flush=True,
                        )

        print(f"\n[download] Saved to {local_path} ({downloaded:,} bytes)")
        return {"local_path": str(local_path), "size": downloaded, "name": filename}

    def get_metadata(self, file_path):
        """Get full metadata for a single file.

        Args:
            file_path: OneDrive path (e.g. "/Documents/report.pdf").

        Returns:
            Dict of the full Graph API item response.
        """
        self._require_token()
        url = self._item_url(file_path)
        data = self._get(url)
        return data

    # ─────────────────────────────────────────────
    # Hashes
    # ─────────────────────────────────────────────

    def _extract_hashes(self, file_info):
        """Fetch item metadata and extract hash values."""
        try:
            url = f"{GRAPH_BASE}/me/drive/items/{file_info['id']}"
            data = self._get(url)
            hashes = data.get("file", {}).get("hashes", {})
            return {
                "filename": file_info["name"],
                "folder": file_info["folder"],
                "path": file_info["path"],
                "sha1": hashes.get("sha1Hash", ""),
                "sha256": hashes.get("sha256Hash", ""),
                "quickXorHash": hashes.get("quickXorHash", ""),
                "size": file_info["size"],
            }
        except Exception as e:
            print(f"  [error] {file_info['path']}: {e}")
            return {
                "filename": file_info["name"],
                "folder": file_info["folder"],
                "path": file_info["path"],
                "sha1": "",
                "sha256": "",
                "quickXorHash": "",
                "size": file_info["size"],
            }

    def get_hashes(self, folder_path, recursive=True):
        """Get SHA1/SHA256/quickXorHash from file metadata (no download needed).

        Args:
            folder_path: OneDrive folder path.
            recursive: Recurse into subfolders.

        Returns:
            List of dicts with keys: filename, folder, path, sha1, sha256,
            quickXorHash, size. Also writes onedrive_hashes.csv.
        """
        self._require_token()
        print(f"[hashes] Listing files in '{folder_path}'...")
        files = self.list_files(folder_path, recursive=recursive)
        print(f"[hashes] Fetching hashes for {len(files)} files...")

        results = []
        completed = 0
        start_time = time.time()

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {
                executor.submit(self._extract_hashes, f): f for f in files
            }
            for future in as_completed(futures):
                completed += 1
                result = future.result()
                results.append(result)
                self._progress(completed, len(files), start_time)

        # Write CSV
        csv_path = self._output_path("hashes", folder_path=folder_path)
        fieldnames = ["filename", "folder", "path", "sha1", "sha256", "quickXorHash", "size"]
        with open(csv_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for r in sorted(results, key=lambda x: x["path"]):
                writer.writerow(r)

        print(f"[hashes] Wrote {len(results)} entries to {csv_path}")
        return results

    # ─────────────────────────────────────────────
    # Sharing & Permissions
    # ─────────────────────────────────────────────

    def _get_permissions(self, item_id):
        """Get all permissions for an item. Returns list of dicts."""
        url = f"{GRAPH_BASE}/me/drive/items/{item_id}/permissions"
        try:
            data = self._get(url)
        except requests.exceptions.HTTPError:
            return []

        permissions = []
        for perm in data.get("value", []):
            link = perm.get("link", {})
            inherited = perm.get("inheritedFrom")
            p = {
                "id": perm.get("id"),
                "type": link.get("type", perm.get("roles", ["unknown"])[0] if perm.get("roles") else "unknown"),
                "scope": link.get("scope", ""),
                "url": link.get("webUrl", ""),
                "inherited": bool(inherited),
                "inherited_from": inherited.get("path", "") if inherited else "",
                "has_password": bool(perm.get("hasPassword") or link.get("password")),
                "prevents_download": bool(link.get("preventsDownload")),
                "expiration": perm.get("expirationDateTime", ""),
                "granted_to": "",
            }
            granted = perm.get("grantedToV2") or perm.get("grantedTo")
            if granted:
                user = granted.get("user", {})
                p["granted_to"] = user.get("displayName", user.get("email", ""))
            permissions.append(p)
        return permissions

    def _audit_one(self, file_info):
        """Audit permissions for a single file."""
        perms = self._get_permissions(file_info["id"])
        return file_info, perms

    def audit_permissions(self, folder_path, recursive=True):
        """Audit sharing permissions for all files in a folder.

        Args:
            folder_path: OneDrive folder path.
            recursive: Recurse into subfolders.

        Returns:
            List of dicts with permission details. Also writes
            onedrive_audit.csv and onedrive_audit.md.
        """
        self._require_token()
        print(f"[audit] Listing files in '{folder_path}'...")
        files = self.list_files(folder_path, recursive=recursive)
        print(f"[audit] Auditing permissions for {len(files)} files...")

        results = []
        files_with_links = 0
        total_permissions = 0
        completed = 0
        start_time = time.time()

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(self._audit_one, f): f for f in files}
            for future in as_completed(futures):
                completed += 1
                file_info, perms = future.result()
                total_permissions += len(perms)
                if perms:
                    files_with_links += 1

                self._progress(completed, len(files), start_time)

                for p in perms:
                    results.append({
                        "filename": file_info["name"],
                        "folder": file_info["folder"],
                        "path": file_info["path"],
                        "permission_id": p["id"],
                        "permission_type": p["type"],
                        "scope": p["scope"],
                        "url": p["url"],
                        "inherited": "yes" if p["inherited"] else "no",
                        "inherited_from": p["inherited_from"],
                        "has_password": "yes" if p["has_password"] else "no",
                        "prevents_download": "yes" if p["prevents_download"] else "no",
                        "expiration": p["expiration"],
                        "granted_to": p["granted_to"],
                    })

                if not perms:
                    results.append({
                        "filename": file_info["name"],
                        "folder": file_info["folder"],
                        "path": file_info["path"],
                        "permission_id": "",
                        "permission_type": "(none)",
                        "scope": "",
                        "url": "",
                        "inherited": "",
                        "inherited_from": "",
                        "has_password": "",
                        "prevents_download": "",
                        "expiration": "",
                        "granted_to": "",
                    })

        # Write CSV
        csv_path = self._output_path("audit", folder_path=folder_path)
        fieldnames = [
            "filename", "folder", "path", "permission_id", "permission_type",
            "scope", "url", "inherited", "inherited_from", "has_password",
            "prevents_download", "expiration", "granted_to",
        ]
        with open(csv_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for r in sorted(results, key=lambda x: x["path"]):
                writer.writerow(r)

        # Write Markdown
        md_path = self._output_path("audit", "md", folder_path=folder_path)
        self._write_audit_markdown(results, md_path, folder_path)

        print(f"[audit] CSV: {csv_path}")
        print(f"[audit] Markdown: {md_path}")
        print(f"[audit] Files: {len(files)} | With sharing: {files_with_links} | Permissions: {total_permissions}")
        return results

    def _write_audit_markdown(self, results, output_path, base_folder):
        """Write audit results as markdown grouped by subfolder."""
        grouped = defaultdict(list)
        for r in results:
            folder = r["folder"]
            if base_folder and base_folder != "/":
                rel = folder.replace(base_folder, "", 1).strip("/")
            else:
                rel = folder.strip("/")
            grouped[rel or "(root)"].append(r)

        with open(output_path, "w") as f:
            f.write("# OneDrive Permissions Audit\n\n")
            for folder in sorted(grouped.keys()):
                f.write(f"## {folder}/\n\n")
                files_in_folder = defaultdict(list)
                for r in grouped[folder]:
                    files_in_folder[r["filename"]].append(r)

                for filename in sorted(files_in_folder.keys()):
                    perms = files_in_folder[filename]
                    if len(perms) == 1 and perms[0]["permission_type"] == "(none)":
                        f.write(f"- **{filename}** — no sharing\n")
                    else:
                        f.write(f"- **{filename}**\n")
                        for p in perms:
                            parts = [p["permission_type"]]
                            if p["scope"]:
                                parts.append(f"scope={p['scope']}")
                            if p["inherited"] == "yes":
                                parts.append(f"inherited from {p['inherited_from']}")
                            if p["has_password"] == "yes":
                                parts.append("password-protected")
                            if p["prevents_download"] == "yes":
                                parts.append("download-blocked")
                            if p["expiration"]:
                                parts.append(f"expires {p['expiration']}")
                            if p["granted_to"]:
                                parts.append(f"shared with {p['granted_to']}")
                            detail = " | ".join(parts)
                            if p["url"]:
                                f.write(f"  - [{detail}]({p['url']})\n")
                            else:
                                f.write(f"  - {detail}\n")
                f.write("\n")

    def create_view_links(self, folder_path, recursive=True, password=None):
        """Create view-only sharing links for all files, reusing existing ones.

        Args:
            folder_path: OneDrive folder path.
            recursive: Recurse into subfolders.
            password: Optional password to protect new links (OneDrive Personal only).

        Returns:
            List of dicts with keys: id, name, path, folder, size, link, reused.
            Also writes onedrive_links.csv and onedrive_links.md.
        """
        self._require_token()
        print(f"[links] Listing files in '{folder_path}'...")
        files = self.list_files(folder_path, recursive=recursive)
        print(f"[links] Processing {len(files)} files...")
        if password:
            print(f"[links] New links will be password-protected.")

        results = []
        reused = 0
        created = 0
        errors = 0
        completed = 0
        start_time = time.time()

        def _process_one(file_info):
            perms = self._get_permissions(file_info["id"])
            if not password:
                # Without password, reuse any existing direct view link
                for p in perms:
                    if p["type"] == "view" and not p["inherited"] and p["url"]:
                        return file_info, p["url"], True
            else:
                # With password, reuse only if a password-protected view link exists
                for p in perms:
                    if p["type"] == "view" and not p["inherited"] and p["url"] and p["has_password"]:
                        return file_info, p["url"], True
            # Create new link
            url = f"{GRAPH_BASE}/me/drive/items/{file_info['id']}/createLink"
            body = {"type": "view", "scope": "anonymous"}
            if password:
                body["password"] = password
            data = self._post(url, body)
            return file_info, data["link"]["webUrl"], False

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(_process_one, f): f for f in files}
            for future in as_completed(futures):
                completed += 1
                try:
                    file_info, link, was_reused = future.result()
                    if was_reused:
                        reused += 1
                    else:
                        created += 1
                    results.append({
                        "id": file_info["id"],
                        "name": file_info["name"],
                        "path": file_info["path"],
                        "folder": file_info["folder"],
                        "size": file_info["size"],
                        "link": link,
                        "reused": was_reused,
                    })
                except Exception as e:
                    errors += 1
                    f = futures[future]
                    print(f"  [error] {f['path']}: {e}")
                self._progress(completed, len(files), start_time)

        # Write CSV
        csv_path = self._output_path("links", folder_path=folder_path)
        with open(csv_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["filename", "folder", "path", "link"])
            writer.writeheader()
            for r in sorted(results, key=lambda x: x["path"]):
                writer.writerow({
                    "filename": r["name"],
                    "folder": r["folder"],
                    "path": r["path"],
                    "link": r["link"],
                })

        # Write Markdown
        md_path = self._output_path("links", "md", folder_path=folder_path)
        grouped = defaultdict(list)
        for r in results:
            folder = r["folder"]
            if folder_path and folder_path != "/":
                rel = folder.replace(folder_path, "", 1).strip("/")
            else:
                rel = folder.strip("/")
            grouped[rel or "(root)"].append(r)

        with open(md_path, "w") as f:
            f.write("# OneDrive Shared Links\n\n")
            for folder in sorted(grouped.keys()):
                f.write(f"## {folder}/\n\n")
                for r in sorted(grouped[folder], key=lambda x: x["name"]):
                    f.write(f"- [{r['name']}]({r['link']})\n")
                f.write("\n")

        print(f"[links] CSV: {csv_path}")
        print(f"[links] Markdown: {md_path}")
        print(f"[links] Reused: {reused} | Created: {created} | Errors: {errors}")
        return results

    def revoke_permission(self, file_path, permission_id):
        """Revoke a specific permission from a file.

        Args:
            file_path: OneDrive path to the file.
            permission_id: The permission ID to revoke.

        Returns:
            Dict with keys: file_path, permission_id, status.
        """
        self._require_token()
        path = file_path.strip("/")
        url = f"{GRAPH_BASE}/me/drive/root:/{path}:/permissions/{permission_id}"
        try:
            self._delete(url)
            print(f"[revoke] Revoked permission {permission_id} from /{path}")
            return {"file_path": file_path, "permission_id": permission_id, "status": "revoked"}
        except requests.exceptions.HTTPError as e:
            print(f"[revoke] Error revoking {permission_id} from /{path}: {e}")
            return {"file_path": file_path, "permission_id": permission_id, "status": f"error: {e}"}

    def revoke_edit_links(self, folder_path, recursive=True):
        """Find and revoke all direct (non-inherited) edit links in a folder.

        Prompts for confirmation before revoking. Only revokes links with
        type == "edit" that are NOT inherited.

        Args:
            folder_path: OneDrive folder path.
            recursive: Recurse into subfolders.

        Returns:
            List of dicts describing what was revoked.
        """
        self._require_token()
        print(f"[revoke-edit] Listing files in '{folder_path}'...")
        files = self.list_files(folder_path, recursive=recursive)
        print(f"[revoke-edit] Scanning permissions for {len(files)} files...")

        # Find all direct edit links
        edit_links = []
        completed = 0
        start_time = time.time()

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(self._audit_one, f): f for f in files}
            for future in as_completed(futures):
                completed += 1
                file_info, perms = future.result()
                for p in perms:
                    if p["type"] == "edit" and not p["inherited"]:
                        edit_links.append({
                            "file_id": file_info["id"],
                            "file_path": file_info["path"],
                            "file_name": file_info["name"],
                            "permission_id": p["id"],
                            "scope": p["scope"],
                            "url": p["url"],
                        })
                self._progress(completed, len(files), start_time)

        if not edit_links:
            print("[revoke-edit] No direct edit links found.")
            return []

        print(f"\n[revoke-edit] Found {len(edit_links)} direct edit link(s):")
        for el in edit_links:
            print(f"  - {el['file_path']} (scope={el['scope']}, id={el['permission_id']})")

        confirm = input(f"\nRevoke all {len(edit_links)} edit link(s)? [y/N]: ").strip().lower()
        if confirm != "y":
            print("[revoke-edit] Cancelled.")
            return []

        # Revoke them
        results = []
        for el in edit_links:
            url = f"{GRAPH_BASE}/me/drive/items/{el['file_id']}/permissions/{el['permission_id']}"
            try:
                self._delete(url)
                print(f"  Revoked: {el['file_path']}")
                results.append({**el, "status": "revoked"})
            except requests.exceptions.HTTPError as e:
                print(f"  Error: {el['file_path']}: {e}")
                results.append({**el, "status": f"error: {e}"})

        # Write CSV
        csv_path = self._output_path("revoked_edit_links", folder_path=folder_path)
        fieldnames = ["file_path", "file_name", "permission_id", "scope", "url", "status"]
        with open(csv_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            for r in results:
                writer.writerow(r)

        print(f"[revoke-edit] Results written to {csv_path}")
        return results

    def replace_edit_with_view(self, folder_path):
        """Revoke the direct edit link on a folder and create a view-only link.

        Operates on the folder item itself, not its children.

        Args:
            folder_path: OneDrive folder path.

        Returns:
            Dict with keys: folder_path, revoked_edit, new_view_link.
        """
        self._require_token()
        path = folder_path.strip("/")
        if not path:
            folder_url = f"{GRAPH_BASE}/me/drive/root"
        else:
            folder_url = f"{GRAPH_BASE}/me/drive/root:/{path}"

        # Get folder item to find its ID
        folder_item = self._get(folder_url)
        folder_id = folder_item["id"]

        # Get permissions on the folder
        perms = self._get_permissions(folder_id)
        revoked_count = 0

        for p in perms:
            if p["type"] == "edit" and not p["inherited"]:
                url = f"{GRAPH_BASE}/me/drive/items/{folder_id}/permissions/{p['id']}"
                try:
                    self._delete(url)
                    revoked_count += 1
                    print(f"[replace] Revoked edit link: {p['url']}")
                except requests.exceptions.HTTPError as e:
                    print(f"[replace] Error revoking edit link {p['id']}: {e}")

        # Create view-only link
        link_url = f"{GRAPH_BASE}/me/drive/items/{folder_id}/createLink"
        data = self._post(link_url, {"type": "view", "scope": "anonymous"})
        view_link = data["link"]["webUrl"]
        print(f"[replace] Created view-only link: {view_link}")

        return {
            "folder_path": folder_path,
            "revoked_edit": revoked_count,
            "new_view_link": view_link,
        }

    # ─────────────────────────────────────────────
    # Version History
    # ─────────────────────────────────────────────

    def list_versions(self, file_path):
        """List all versions of a file.

        Args:
            file_path: OneDrive path to the file.

        Returns:
            List of dicts with keys: id, lastModified, size.
        """
        self._require_token()
        path = file_path.strip("/")
        url = f"{GRAPH_BASE}/me/drive/root:/{path}:/versions"
        data = self._get(url)

        versions = []
        for v in data.get("value", []):
            versions.append({
                "id": v.get("id"),
                "lastModified": v.get("lastModifiedDateTime", ""),
                "size": v.get("size", 0),
            })

        print(f"[versions] {len(versions)} version(s) of /{path}:")
        for v in versions:
            size_kb = v["size"] / 1024 if v["size"] else 0
            print(f"  {v['id']}  {v['lastModified']}  {size_kb:.1f} KB")
        return versions

    def download_version(self, file_path, version_id, local_dir="."):
        """Download a specific version of a file.

        Args:
            file_path: OneDrive path to the file.
            version_id: The version ID to download (from list_versions).
            local_dir: Local directory to save the file.

        Returns:
            Dict with keys: local_path, size, name, version_id.
        """
        self._require_token()
        path = file_path.strip("/")
        # Resolve file to get item ID
        item = self._get(f"{GRAPH_BASE}/me/drive/root:/{path}")
        item_id = item["id"]

        # Get version metadata — it includes a direct download URL
        ver_data = self._get(f"{GRAPH_BASE}/me/drive/items/{item_id}/versions/{version_id}")
        download_url = ver_data.get("@microsoft.graph.downloadUrl")

        if download_url:
            # Use the pre-authenticated download URL from version metadata
            resp = requests.get(download_url, stream=True)
        else:
            # Fall back to content endpoint (works for non-current versions)
            url = f"{GRAPH_BASE}/me/drive/items/{item_id}/versions/{version_id}/content"
            resp = requests.get(url, headers=self._headers(), stream=True, allow_redirects=True)
        resp.raise_for_status()

        local_dir = Path(local_dir)
        local_dir.mkdir(parents=True, exist_ok=True)
        base_name = os.path.basename(path)
        name, ext = os.path.splitext(base_name)
        filename = f"{name}_v{version_id}{ext}"
        local_path = local_dir / filename

        total = int(resp.headers.get("Content-Length", 0))
        downloaded = 0
        start = time.time()

        with open(local_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total > 0:
                        pct = downloaded / total * 100
                        elapsed = time.time() - start
                        speed = downloaded / elapsed / 1024 / 1024 if elapsed > 0 else 0
                        print(
                            f"\r  Downloading {filename}: {pct:.1f}% ({speed:.1f} MB/s)",
                            end="",
                            flush=True,
                        )

        print(f"\n[download-version] Saved to {local_path} ({downloaded:,} bytes)")
        return {
            "local_path": str(local_path),
            "size": downloaded,
            "name": filename,
            "version_id": version_id,
        }

    # ─────────────────────────────────────────────
    # Activity
    # ─────────────────────────────────────────────

    def get_activity(self, folder_path=None):
        """Get recent activity/changes via delta query.

        Args:
            folder_path: Optional folder to scope. None = entire drive.

        Returns:
            List of dicts with keys: id, name, path, action, lastModified, size.
            Also writes onedrive_activity.csv.
        """
        self._require_token()
        if folder_path and folder_path != "/":
            path = folder_path.strip("/")
            # Get folder item ID first
            folder_item = self._get(f"{GRAPH_BASE}/me/drive/root:/{path}")
            url = f"{GRAPH_BASE}/me/drive/items/{folder_item['id']}/delta"
        else:
            url = f"{GRAPH_BASE}/me/drive/root/delta"

        results = []
        page = 0
        while url:
            page += 1
            print(f"  Fetching activity page {page}...", flush=True)
            data = self._get(url)
            for item in data.get("value", []):
                parent = item.get("parentReference", {})
                parent_path = parent.get("path", "")
                if ":/drive/root:" in parent_path:
                    parent_path = parent_path.split(":/drive/root:")[-1]
                elif parent_path.startswith("/drive/root:"):
                    parent_path = parent_path[len("/drive/root:"):]
                elif parent_path.startswith("/drive/root"):
                    parent_path = parent_path[len("/drive/root"):]
                if not parent_path:
                    parent_path = "/"

                item_path = f"{parent_path.rstrip('/')}/{item.get('name', '')}"

                # Determine action
                if item.get("deleted"):
                    action = "deleted"
                elif "folder" in item:
                    action = "folder"
                else:
                    action = "modified"

                results.append({
                    "id": item.get("id", ""),
                    "name": item.get("name", ""),
                    "path": item_path,
                    "action": action,
                    "lastModified": item.get("lastModifiedDateTime", ""),
                    "size": item.get("size", 0),
                })

            # Follow nextLink for more pages, but stop at deltaLink
            next_link = data.get("@odata.nextLink")
            if next_link:
                url = next_link
            else:
                url = None
                # Store deltaLink for future incremental queries if desired
                delta_link = data.get("@odata.deltaLink", "")
                if delta_link:
                    delta_path = self.output_dir / ".onedrive_delta_link.txt"
                    delta_path.write_text(delta_link)

        # Write CSV
        csv_path = self._output_path("activity", folder_path=folder_path)
        fieldnames = ["id", "name", "path", "action", "lastModified", "size"]
        with open(csv_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for r in sorted(results, key=lambda x: x.get("lastModified", ""), reverse=True):
                writer.writerow(r)

        print(f"[activity] {len(results)} items in delta. CSV: {csv_path}")
        return results


# ─────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────

def _browse_folder(od, start_path="/"):
    """Interactive folder browser. Returns selected folder path or None if cancelled."""
    browse_path = start_path
    while True:
        print(f"\n{'─' * 60}")
        print(f"📂 {browse_path}")
        print(f"{'─' * 60}")
        folders, files = od.list_folder_contents(browse_path)
        sorted_folders = sorted(folders, key=lambda x: x["name"])

        if browse_path != "/":
            print(f"  ../ (go up)")
        for i, f in enumerate(sorted_folders, 1):
            print(f"  {i:>3}) 📁 {f['name']}/  ({f['childCount']} items)")
        if files:
            print()
            for f in sorted(files, key=lambda x: x["name"]):
                size_kb = f["size"] / 1024 if f["size"] else 0
                print(f"       📄 {f['name']}  ({size_kb:.1f} KB)")

        print(f"""
Actions:
  #     → open folder (drill into it)
  s #   → SELECT that folder (e.g. 's 3')
  s     → SELECT current folder
  ..    → go up one level
  t     → type a path manually
  b     → back / cancel""")

        nav = input("\n> ").strip()

        if nav.lower() == "b":
            return None
        elif nav.lower() == "s":
            return browse_path
        elif nav.lower().startswith("s "):
            try:
                idx = int(nav.split()[1]) - 1
                if 0 <= idx < len(sorted_folders):
                    return sorted_folders[idx]["path"]
                else:
                    print("Invalid number.")
            except (ValueError, IndexError):
                print("Usage: s <number>  (e.g. 's 3')")
        elif nav.lower() == "t":
            typed = input("Enter full path: ").strip()
            if typed:
                browse_path = typed
        elif nav == "..":
            if browse_path != "/":
                browse_path = os.path.dirname(browse_path.rstrip("/")) or "/"
        elif nav.isdigit():
            idx = int(nav) - 1
            if 0 <= idx < len(sorted_folders):
                browse_path = sorted_folders[idx]["path"]
            else:
                print("Invalid number.")
        else:
            print("Invalid choice.")


def _browse_file(od, start_path="/"):
    """Interactive file browser. Returns selected file path or None if cancelled."""
    browse_path = start_path
    while True:
        print(f"\n{'─' * 60}")
        print(f"Browsing: {browse_path}")
        print(f"{'─' * 60}")
        folders, files = od.list_folder_contents(browse_path)

        num = 0
        if browse_path != "/":
            print(f"  ../ (go up)")
        for f in sorted(folders, key=lambda x: x["name"]):
            num += 1
            print(f"  {num:>3}) 📁 {f['name']}/  ({f['childCount']} items)")
        folder_count = num
        for f in sorted(files, key=lambda x: x["name"]):
            num += 1
            size_kb = f["size"] / 1024 if f["size"] else 0
            print(f"  {num:>3}) 📄 {f['name']}  ({size_kb:.1f} KB)")

        print(f"\n  t) TYPE a path manually")
        print(f"  b) BACK / cancel")

        nav = input("\nChoose # or action: ").strip().lower()

        if nav == "b":
            return None
        elif nav == "t":
            typed = input("Enter full file path: ").strip()
            if typed:
                return typed
        elif nav == "..":
            if browse_path != "/":
                browse_path = os.path.dirname(browse_path.rstrip("/")) or "/"
        elif nav.isdigit():
            idx = int(nav) - 1
            sorted_folders = sorted(folders, key=lambda x: x["name"])
            sorted_files = sorted(files, key=lambda x: x["name"])
            if idx < len(sorted_folders):
                browse_path = sorted_folders[idx]["path"]
            elif idx < len(sorted_folders) + len(sorted_files):
                return sorted_files[idx - len(sorted_folders)]["path"]
            else:
                print("Invalid number.")
        else:
            print("Invalid choice.")


def _pick_folder(od, current_folder, label="Target folder"):
    """Quick prompt to pick a folder: use current, browse, or type."""
    print(f"\n{label}: {current_folder}")
    choice = input("  Enter=use this | b=browse | or type a path: ").strip()
    if not choice:
        return current_folder
    elif choice.lower() == "b":
        result = _browse_folder(od, current_folder)
        return result  # None if cancelled
    else:
        return choice


def _pick_file(od, current_folder):
    """Quick prompt to pick a file: browse or type."""
    choice = input(f"\n  Enter=browse from '{current_folder}' | or type a file path: ").strip()
    if not choice:
        return _browse_file(od, current_folder)
    else:
        return choice


def interactive_mode():
    """Interactive menu-driven interface."""
    od = OneDriveTools()
    od.authenticate()
    current_folder = "/"

    while True:
        print(f"""
OneDrive Tools
{'━' * 40}
Current folder: {current_folder}
Output dir:     {od.output_dir}

 1) List files
 2) Search files
 3) Audit permissions
 4) Get file hashes
 5) Create view-only links
 6) Download file
 7) File versions
 8) Recent activity
 9) Revoke edit links
10) Replace edit → view-only
11) File metadata
12) Download specific version
 c) Change folder
 q) Quit
""")
        choice = input("Choose: ").strip().lower()

        try:
            if choice == "q":
                print("Goodbye!")
                break

            elif choice == "c":
                result = _browse_folder(od, current_folder)
                if result:
                    current_folder = result
                    print(f"\n✓ Working folder set to: {current_folder}")

            elif choice == "1":
                folder = _pick_folder(od, current_folder, "List files in")
                if not folder:
                    continue
                recursive = input("Include subfolders? [Y/n]: ").strip().lower() != "n"
                if not recursive:
                    folders, files = od.list_folder_contents(folder)
                    print(f"\n{'─' * 60}")
                    for f in sorted(folders, key=lambda x: x["name"]):
                        print(f"  📁 {f['name']}/  ({f['childCount']} items)")
                    for f in sorted(files, key=lambda x: x["name"]):
                        size_kb = f["size"] / 1024 if f["size"] else 0
                        print(f"  📄 {f['name']}  ({size_kb:.1f} KB)")
                    print(f"{'─' * 60}")
                    print(f"Folders: {len(folders)} | Files: {len(files)}")
                    # Save CSV
                    csv_path = od._output_path("files", folder_path=folder)
                    with open(csv_path, "w", newline="") as csvf:
                        writer = csv.DictWriter(csvf, fieldnames=["name", "path", "size", "lastModified"])
                        writer.writeheader()
                        for f in sorted(files, key=lambda x: x["name"]):
                            writer.writerow(f)
                    print(f"CSV: {csv_path}")
                else:
                    files = od.list_files(folder, recursive=True)
                    print(f"\n{'─' * 60}")
                    for f in sorted(files, key=lambda x: x["path"]):
                        size_kb = f["size"] / 1024 if f["size"] else 0
                        print(f"  {f['path']}  ({size_kb:.1f} KB)")
                    print(f"{'─' * 60}")
                    print(f"Total: {len(files)} files")
                    # Save CSV
                    csv_path = od._output_path("files", folder_path=folder)
                    fieldnames = ["name", "folder", "path", "size", "lastModified", "mimeType", "webUrl"]
                    with open(csv_path, "w", newline="") as csvf:
                        writer = csv.DictWriter(csvf, fieldnames=fieldnames, extrasaction="ignore")
                        writer.writeheader()
                        for f in sorted(files, key=lambda x: x["path"]):
                            writer.writerow(f)
                    print(f"CSV: {csv_path}")

            elif choice == "2":
                query = input("Search query: ").strip()
                if not query:
                    continue
                scope = input(f"Scope to folder? [{current_folder}] (Enter=yes, 'all'=entire drive): ").strip()
                folder = None if scope.lower() == "all" else current_folder
                results = od.search(query, folder_path=folder)
                for r in results:
                    size_kb = r["size"] / 1024 if r["size"] else 0
                    print(f"  {r['path']}  ({size_kb:.1f} KB)")

            elif choice == "3":
                folder = _pick_folder(od, current_folder, "Audit permissions in")
                if not folder:
                    continue
                recursive = input("Include subfolders? [Y/n]: ").strip().lower() != "n"
                od.audit_permissions(folder, recursive=recursive)

            elif choice == "4":
                folder = _pick_folder(od, current_folder, "Get hashes for")
                if not folder:
                    continue
                recursive = input("Include subfolders? [Y/n]: ").strip().lower() != "n"
                od.get_hashes(folder, recursive=recursive)

            elif choice == "5":
                folder = _pick_folder(od, current_folder, "Create view-only links in")
                if not folder:
                    continue
                recursive = input("Include subfolders? [Y/n]: ").strip().lower() != "n"
                print(f"\nScanning files in '{folder}'...")
                files = od.list_files(folder, recursive=recursive)
                print(f"Found {len(files)} files.")

                # Check existing links
                print("Checking for existing view-only links...")
                has_link = 0
                needs_link = 0
                for f in files:
                    perms = od._get_permissions(f["id"])
                    has_direct_view = any(
                        p["type"] == "view" and not p["inherited"] for p in perms
                    )
                    if has_direct_view:
                        has_link += 1
                    else:
                        needs_link += 1

                print(f"\n{'─' * 60}")
                print(f"  Already have view-only link: {has_link} (will SKIP)")
                print(f"  Need new view-only link:     {needs_link} (will CREATE)")
                print(f"{'─' * 60}")

                if needs_link == 0:
                    print("All files already have view-only links. Nothing to do.")
                    continue

                confirm = input(f"Create {needs_link} new view-only link(s)? [y/N]: ").strip().lower()
                if confirm == "y":
                    pw = input("Password-protect new links? (enter password or leave blank for none): ").strip() or None
                    od.create_view_links(folder, recursive=recursive, password=pw)
                else:
                    print("Cancelled.")

            elif choice == "6":
                file_path = _pick_file(od, current_folder)
                if not file_path:
                    continue
                local_dir = input(f"Save to directory [{od.output_dir}]: ").strip() or str(od.output_dir)
                od.download(file_path, local_dir=local_dir)

            elif choice == "7":
                file_path = _pick_file(od, current_folder)
                if not file_path:
                    continue
                od.list_versions(file_path)

            elif choice == "8":
                folder = _pick_folder(od, current_folder, "Get activity for")
                if not folder:
                    continue
                od.get_activity(folder_path=folder)

            elif choice == "9":
                folder = _pick_folder(od, current_folder, "Revoke edit links in")
                if not folder:
                    continue
                recursive = input("Include subfolders? [Y/n]: ").strip().lower() != "n"
                print(f"\nScanning for direct edit links in '{folder}'...")
                files = od.list_files(folder, recursive=recursive)
                edit_links = []
                for f in files:
                    perms = od._get_permissions(f["id"])
                    for p in perms:
                        if p["type"] == "edit" and not p["inherited"]:
                            edit_links.append((f, p))

                if not edit_links:
                    print("No direct edit links found.")
                    continue

                print(f"\n{'─' * 60}")
                print(f"Found {len(edit_links)} direct edit link(s):\n")
                for i, (f, p) in enumerate(edit_links, 1):
                    print(f"  {i}. {f['name']}")
                    print(f"     Current: edit | scope={p['scope']} | {p['url']}")
                    print(f"     Action:  REVOKE this link")
                    print()
                print(f"{'─' * 60}")

                confirm = input(f"Revoke all {len(edit_links)} edit link(s)? [y/N/numbers e.g. 1,3]: ").strip()
                if confirm.lower() == "y":
                    od.revoke_edit_links(folder, recursive=recursive)
                elif confirm and confirm[0].isdigit():
                    # Selective revocation
                    indices = [int(x.strip()) - 1 for x in confirm.split(",")]
                    for idx in indices:
                        if 0 <= idx < len(edit_links):
                            f, p = edit_links[idx]
                            od.revoke_permission(f["path"], p["id"])
                else:
                    print("Cancelled.")

            elif choice == "10":
                use_current = input(f"Use current folder '{current_folder}'? [Y/n]: ").strip().lower()
                if use_current == "n":
                    folder = _browse_folder(od, current_folder)
                    if not folder:
                        continue
                else:
                    folder = current_folder
                # Preview
                path = folder.strip("/")
                if not path:
                    folder_url = f"{GRAPH_BASE}/me/drive/root"
                else:
                    folder_url = f"{GRAPH_BASE}/me/drive/root:/{path}"
                folder_item = od._get(folder_url)
                perms = od._get_permissions(folder_item["id"])

                edit_perms = [p for p in perms if p["type"] == "edit" and not p["inherited"]]
                view_perms = [p for p in perms if p["type"] == "view" and not p["inherited"]]

                print(f"\n{'─' * 60}")
                print(f"Folder: {folder}")
                if edit_perms:
                    for p in edit_perms:
                        print(f"  ✗ edit | scope={p['scope']} | {p['url']}  ← will be REVOKED")
                else:
                    print("  (no direct edit links to revoke)")
                if view_perms:
                    for p in view_perms:
                        print(f"  ✓ view | scope={p['scope']} | {p['url']}  (already exists)")
                    print("  Action: Will keep existing view link")
                else:
                    print("  ✓ view | anonymous | (new link will be created)")

                # Count affected files
                files = od.list_files(folder, recursive=True)
                print(f"\nThis affects {len(files)} files that inherit from this folder.")
                print(f"{'─' * 60}")

                if not edit_perms and view_perms:
                    print("Nothing to change — already has view-only, no edit links.")
                    continue

                confirm = input("Proceed? [y/N]: ").strip().lower()
                if confirm == "y":
                    od.replace_edit_with_view(folder)
                else:
                    print("Cancelled.")

            elif choice == "11":
                file_path = _pick_file(od, current_folder)
                if not file_path:
                    continue
                meta = od.get_metadata(file_path)
                print(json.dumps(meta, indent=2, default=str))

            elif choice == "12":
                file_path = _pick_file(od, current_folder)
                if not file_path:
                    continue
                versions = od.list_versions(file_path)
                if not versions:
                    continue
                ver_id = input("Version ID to download: ").strip()
                if not ver_id:
                    continue
                local_dir = input(f"Save to directory [{od.output_dir}]: ").strip() or str(od.output_dir)
                od.download_version(file_path, ver_id, local_dir=local_dir)

            else:
                print("Invalid choice. Try again.")

        except KeyboardInterrupt:
            print("\n\nOperation cancelled.")
        except Exception as e:
            print(f"\nError: {e}")

    return


def main():
    # If no args, launch interactive mode
    if len(sys.argv) == 1:
        interactive_mode()
        return

    parser = argparse.ArgumentParser(
        description="OneDrive Tools — manage files, permissions, versions via Microsoft Graph",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 onedrive_tools.py list "/Documents"
  python3 onedrive_tools.py search "contract" "/Legal"
  python3 onedrive_tools.py download "/Documents/report.pdf" ./downloads
  python3 onedrive_tools.py metadata "/Documents/report.pdf"
  python3 onedrive_tools.py hashes "/Documents"
  python3 onedrive_tools.py audit "/Documents"
  python3 onedrive_tools.py links "/Documents"
  python3 onedrive_tools.py revoke "/Documents/file.pdf" "permission-id"
  python3 onedrive_tools.py revoke-edit-links "/Documents"
  python3 onedrive_tools.py replace-edit-view "/Documents"
  python3 onedrive_tools.py versions "/Documents/report.pdf"
  python3 onedrive_tools.py download-version "/Documents/report.pdf" "1.0" ./downloads
  python3 onedrive_tools.py activity "/Documents"
        """,
    )

    subparsers = parser.add_subparsers(dest="command", required=True)

    # list
    p_list = subparsers.add_parser("list", help="List files with metadata")
    p_list.add_argument("path", help="OneDrive folder path")
    p_list.add_argument("--no-recursive", action="store_true", help="Don't recurse into subfolders")

    # search
    p_search = subparsers.add_parser("search", help="Full-text search")
    p_search.add_argument("query", help="Search query")
    p_search.add_argument("path", nargs="?", default=None, help="Optional folder to scope search")

    # download
    p_dl = subparsers.add_parser("download", help="Download a file")
    p_dl.add_argument("path", help="OneDrive file path")
    p_dl.add_argument("local_dir", nargs="?", default=".", help="Local directory (default: current)")

    # metadata
    p_meta = subparsers.add_parser("metadata", help="Get file metadata")
    p_meta.add_argument("path", help="OneDrive file path")

    # hashes
    p_hash = subparsers.add_parser("hashes", help="Extract file hashes to CSV")
    p_hash.add_argument("path", help="OneDrive folder path")
    p_hash.add_argument("--no-recursive", action="store_true", help="Don't recurse into subfolders")

    # audit
    p_audit = subparsers.add_parser("audit", help="Audit permissions")
    p_audit.add_argument("path", help="OneDrive folder path")
    p_audit.add_argument("--no-recursive", action="store_true", help="Don't recurse into subfolders")

    # links
    p_links = subparsers.add_parser("links", help="Create view-only sharing links")
    p_links.add_argument("path", help="OneDrive folder path")
    p_links.add_argument("--no-recursive", action="store_true", help="Don't recurse into subfolders")
    p_links.add_argument("--password", default=None, help="Password-protect new links (OneDrive Personal only)")

    # revoke
    p_revoke = subparsers.add_parser("revoke", help="Revoke a specific permission")
    p_revoke.add_argument("path", help="OneDrive file path")
    p_revoke.add_argument("permission_id", help="Permission ID to revoke")

    # revoke-edit-links
    p_rev_edit = subparsers.add_parser("revoke-edit-links", help="Revoke all direct edit links")
    p_rev_edit.add_argument("path", help="OneDrive folder path")
    p_rev_edit.add_argument("--no-recursive", action="store_true", help="Don't recurse into subfolders")

    # replace-edit-view
    p_replace = subparsers.add_parser("replace-edit-view", help="Replace edit link with view link on a folder")
    p_replace.add_argument("path", help="OneDrive folder path")

    # versions
    p_ver = subparsers.add_parser("versions", help="List file versions")
    p_ver.add_argument("path", help="OneDrive file path")

    # download-version
    p_dlv = subparsers.add_parser("download-version", help="Download a specific file version")
    p_dlv.add_argument("path", help="OneDrive file path")
    p_dlv.add_argument("version_id", help="Version ID to download")
    p_dlv.add_argument("local_dir", nargs="?", default=".", help="Local directory (default: current)")

    # activity
    p_act = subparsers.add_parser("activity", help="Get recent activity")
    p_act.add_argument("path", nargs="?", default=None, help="Optional folder path")

    args = parser.parse_args()

    # Initialize and authenticate
    od = OneDriveTools()
    od.authenticate()
    print()

    # Dispatch
    if args.command == "list":
        files = od.list_files(args.path, recursive=not args.no_recursive)
        print(f"\nTotal: {len(files)} files")
        # Write CSV
        csv_path = od._output_path("files", folder_path=args.path)
        fieldnames = ["name", "folder", "path", "size", "lastModified", "mimeType", "webUrl"]
        with open(csv_path, "w", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            for r in sorted(files, key=lambda x: x["path"]):
                writer.writerow(r)
        print(f"CSV: {csv_path}")

    elif args.command == "search":
        results = od.search(args.query, folder_path=args.path)
        for r in results:
            size_kb = r["size"] / 1024 if r["size"] else 0
            print(f"  {r['path']}  ({size_kb:.1f} KB)")

    elif args.command == "download":
        od.download(args.path, local_dir=args.local_dir)

    elif args.command == "metadata":
        meta = od.get_metadata(args.path)
        print(json.dumps(meta, indent=2, default=str))

    elif args.command == "hashes":
        od.get_hashes(args.path, recursive=not args.no_recursive)

    elif args.command == "audit":
        od.audit_permissions(args.path, recursive=not args.no_recursive)

    elif args.command == "links":
        od.create_view_links(args.path, recursive=not args.no_recursive, password=args.password)

    elif args.command == "revoke":
        od.revoke_permission(args.path, args.permission_id)

    elif args.command == "revoke-edit-links":
        od.revoke_edit_links(args.path, recursive=not args.no_recursive)

    elif args.command == "replace-edit-view":
        od.replace_edit_with_view(args.path)

    elif args.command == "versions":
        od.list_versions(args.path)

    elif args.command == "download-version":
        od.download_version(args.path, args.version_id, local_dir=args.local_dir)

    elif args.command == "activity":
        od.get_activity(folder_path=args.path)


if __name__ == "__main__":
    main()
