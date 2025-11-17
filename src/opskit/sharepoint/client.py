import os
import time
import mimetypes
from urllib.parse import urlparse

import requests
from msal import ConfidentialClientApplication

GRAPH = "https://graph.microsoft.com/v1.0"


class SharePointClient:
    def __init__(
        self,
        tenant_id: str,
        client_id: str,
        client_secret: str,
        site_url: str,
        library: str = "Documents",
        retry_max: int = 5,
        retry_backoff: float = 2.0,
        verbose: bool = False,
    ):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.site_url = site_url
        self.library = library
        self.retry_max = retry_max
        self.retry_backoff = retry_backoff
        self.verbose = verbose

        self._token = None
        self._site_id = None
        self._drive_id = None

    # ---------- Public API ----------

    def ensure_connected(self):
        """Ensure access token, site id, and drive id are resolved."""
        if self._token is None:
            self._token = self._get_token()
        if self._site_id is None:
            self._site_id = self._get_site_id()
        if self._drive_id is None:
            self._drive_id = self._get_drive_id()

    @property
    def drive_id(self) -> str:
        self.ensure_connected()
        return self._drive_id

    def ensure_folder_path(self, folder_path: str, conflict_behavior: str = "replace") -> dict:
        """
        Ensure the given folder path exists in the drive and return the final driveItem.
        `folder_path` is relative to the document library root.
        """
        self.ensure_connected()
        folder_path = folder_path.strip("/")
        root = self._get_item_by_path("")  # drive root
        if not root:
            raise RuntimeError("Drive root not found")

        if not folder_path:
            return root

        current = root
        for part in folder_path.split("/"):
            candidate_path = (
                f"{self.get_server_relative_path(current)}/{part}"
                if current.get("parentReference")
                else part
            )
            existing = self._get_item_by_path(candidate_path)
            if existing is None:
                created = self._create_folder(current["id"], part, conflict_behavior)
                current = created
            else:
                current = existing
        return current

    def get_server_relative_path(self, item_json: dict) -> str:
        """
        Reconstruct the path used with /root:/path style calls for a driveItem.
        """
        parent = item_json.get("parentReference", {})
        parent_path = parent.get("path")
        name = item_json.get("name", "")

        if parent_path and parent_path.endswith(":"):
            base = ""
        elif parent_path and ":/" in parent_path:
            base = parent_path.split(":/", 1)[1]
        else:
            base = ""

        if base:
            return f"{base}/{name}".strip("/") if name else base.strip("/")
        return name.strip("/")

    def upload_file(
        self,
        local_file_path: str,
        dest_folder_path: str,
        small_upload_max: int = 4 * 1024 * 1024,
        chunk_size: int = 8 * 1024 * 1024,
        conflict_behavior: str = "replace",
        retry_max: int | None = None,
        retry_backoff: float | None = None,
        dry_run: bool = False,
    ):
        """
        Upload a single local file into dest_folder_path (relative to library root).
        """
        self.ensure_connected()

        file_name = os.path.basename(local_file_path)
        dest_path_with_name = f"{dest_folder_path.strip('/')}/{file_name}".strip("/")

        content_type = mimetypes.guess_type(local_file_path)[0] or "application/octet-stream"
        size = os.path.getsize(local_file_path)

        if dry_run:
            return

        if size <= small_upload_max:
            with open(local_file_path, "rb") as fh:
                self._small_upload(
                    dest_path_with_name,
                    fh.read(),
                    content_type,
                    conflict_behavior,
                    retry_max,
                    retry_backoff,
                )
        else:
            upload_url = self._create_upload_session(
                dest_path_with_name,
                conflict_behavior,
                retry_max,
                retry_backoff,
            )
            self._chunked_upload(
                upload_url,
                local_file_path,
                size,
                chunk_size,
                retry_max,
                retry_backoff,
            )

    # ---------- Internal helpers ----------

    def _request_with_retry(self, method, url, headers=None, retry_max=None, retry_backoff=None, **kwargs):
        rm = retry_max if retry_max is not None else self.retry_max
        rb = retry_backoff if retry_backoff is not None else self.retry_backoff

        for attempt in range(1, rm + 1):
            resp = requests.request(method, url, headers=headers, **kwargs)
            if resp.status_code in (429, 500, 502, 503, 504):
                wait = rb ** attempt
                if self.verbose:
                    print(f"[http] {method} {url} -> {resp.status_code}, retrying in {wait:.1f}s")
                time.sleep(wait)
                continue
            return resp
        return resp  # last attempt

    def _parse_site(self):
        parsed = urlparse(self.site_url)
        host = parsed.netloc
        path = parsed.path.strip("/")
        if path.startswith("sites/"):
            return host, path[len("sites/"):]
        if path.startswith("teams/"):
            return host, path[len("teams/"):]
        return host, path.split("/")[-1]

    def _get_token(self) -> str:
        app = ConfidentialClientApplication(
            self.client_id,
            authority=f"https://login.microsoftonline.com/{self.tenant_id}",
            client_credential=self.client_secret,
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in result:
            raise RuntimeError(f"Auth failed: {result.get('error_description') or result}")
        if self.verbose:
            print("[auth] acquired access token")
        return result["access_token"]

    def _get_site_id(self) -> str:
        host, site_path = self._parse_site()
        url = f"{GRAPH}/sites/{host}:/sites/{site_path}?$select=id"
        headers = {"Authorization": f"Bearer {self._token}"}
        resp = self._request_with_retry("GET", url, headers=headers)
        if not resp.ok:
            raise RuntimeError(f"Failed to resolve site id: {resp.status_code} {resp.text}")
        site_id = resp.json()["id"]
        if self.verbose:
            print(f"[site] site_id={site_id}")
        return site_id

    def _get_drive_id(self) -> str:
        url = f"{GRAPH}/sites/{self._site_id}/drives"
        headers = {"Authorization": f"Bearer {self._token}"}
        resp = self._request_with_retry("GET", url, headers=headers)
        if not resp.ok:
            raise RuntimeError(f"Failed to list drives: {resp.status_code} {resp.text}")
        for drive in resp.json().get("value", []):
            if drive.get("name") == self.library or drive.get("displayName") == self.library:
                drive_id = drive["id"]
                if self.verbose:
                    print(f"[drive] drive_id={drive_id}")
                return drive_id
        names = [d.get("name") or d.get("displayName") for d in resp.json().get("value", [])]
        raise RuntimeError(f"Drive '{self.library}' not found. Available: {names}")

    def _get_item_by_path(self, path_in_drive: str):
        clean_path = path_in_drive.strip("/")
        url = f"{GRAPH}/drives/{self._drive_id}/root"
        if clean_path:
            url += f":/{clean_path}"
        url += "?$select=id,name,folder,file,parentReference"
        headers = {"Authorization": f"Bearer {self._token}"}
        resp = self._request_with_retry("GET", url, headers=headers)
        if resp.status_code == 404:
            return None
        if not resp.ok:
            raise RuntimeError(f"Failed to get item '{path_in_drive}': {resp.status_code} {resp.text}")
        return resp.json()

    def _create_folder(self, parent_item_id: str, name: str, conflict_behavior: str):
        url = f"{GRAPH}/drives/{self._drive_id}/items/{parent_item_id}/children"
        headers = {"Authorization": f"Bearer {self._token}", "Content-Type": "application/json"}
        body = {
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": conflict_behavior,
        }
        resp = self._request_with_retry("POST", url, headers=headers, json=body)
        if not resp.ok:
            raise RuntimeError(f"Failed to create folder '{name}': {resp.status_code} {resp.text}")
        return resp.json()

    def _small_upload(
        self,
        dest_path_with_name: str,
        file_bytes: bytes,
        content_type: str,
        conflict_behavior: str,
        retry_max: int | None,
        retry_backoff: float | None,
    ):
        headers = {"Authorization": f"Bearer {self._token}", "Content-Type": content_type}
        url = (
            f"{GRAPH}/drives/{self._drive_id}/root:/{dest_path_with_name}"
            f":/content?@microsoft.graph.conflictBehavior={conflict_behavior}"
        )
        resp = self._request_with_retry(
            "PUT",
            url,
            headers=headers,
            data=file_bytes,
            retry_max=retry_max,
            retry_backoff=retry_backoff,
        )
        if not resp.ok:
            raise RuntimeError(
                f"Small upload failed for '{dest_path_with_name}': {resp.status_code} {resp.text}"
            )

    def _create_upload_session(
        self,
        dest_path_with_name: str,
        conflict_behavior: str,
        retry_max: int | None,
        retry_backoff: float | None,
    ) -> str:
        url = f"{GRAPH}/drives/{self._drive_id}/root:/{dest_path_with_name}:/createUploadSession"
        headers = {"Authorization": f"Bearer {self._token}", "Content-Type": "application/json"}
        body = {"@microsoft.graph.conflictBehavior": conflict_behavior, "deferCommit": False}
        resp = self._request_with_retry(
            "POST",
            url,
            headers=headers,
            json=body,
            retry_max=retry_max,
            retry_backoff=retry_backoff,
        )
        if not resp.ok:
            raise RuntimeError(
                f"Create upload session failed for '{dest_path_with_name}': {resp.status_code} {resp.text}"
            )
        return resp.json()["uploadUrl"]

    def _chunked_upload(
        self,
        upload_url: str,
        file_path: str,
        file_size: int,
        chunk_size: int,
        retry_max: int | None,
        retry_backoff: float | None,
    ):
        rm = retry_max if retry_max is not None else self.retry_max
        rb = retry_backoff if retry_backoff is not None else self.retry_backoff

        with open(file_path, "rb") as f:
            bytes_sent = 0
            chunk_index = 0
            while bytes_sent < file_size:
                chunk = f.read(chunk_size)
                if not chunk:
                    break
                start = bytes_sent
                end = bytes_sent + len(chunk) - 1
                headers = {
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {start}-{end}/{file_size}",
                }
                resp = self._request_with_retry(
                    "PUT",
                    upload_url,
                    headers=headers,
                    data=chunk,
                    retry_max=rm,
                    retry_backoff=rb,
                )
                if resp.status_code not in (200, 201, 202):
                    raise RuntimeError(
                        f"Chunk upload failed (chunk {chunk_index}, {start}-{end}): "
                        f"{resp.status_code} {resp.text}"
                    )
                bytes_sent += len(chunk)
                chunk_index += 1
                if self.verbose:
                    pct = (bytes_sent / file_size) * 100
                    print(
                        f"[upload] {os.path.basename(file_path)}: "
                        f"{bytes_sent}/{file_size} bytes ({pct:.1f}%)"
                    )


# These helpers are generic (not SharePoint-specific), but convenient to reuse
def iter_local(root_dir: str):
    for dirpath, dirnames, filenames in os.walk(root_dir):
        dirnames.sort()
        filenames.sort()
        yield dirpath, dirnames, filenames


def count_files(root_dir: str) -> int:
    total = 0
    for _, _, filenames in os.walk(root_dir):
        total += len(filenames)
    return total