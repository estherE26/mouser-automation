"""
Dropbox API client for downloading press release folders.
"""
import os
import tempfile
import requests
from typing import Optional


class DropboxClient:
    """Client for interacting with Dropbox API."""

    def __init__(self, access_token: Optional[str] = None):
        self.access_token = access_token or os.environ.get('DROPBOX_TOKEN')
        if not self.access_token:
            raise ValueError("Dropbox access token required")

        self.base_url = "https://api.dropboxapi.com/2"
        self.content_url = "https://content.dropboxapi.com/2"

    def _headers(self) -> dict:
        return {
            "Authorization": f"Bearer {self.access_token}",
            "Content-Type": "application/json"
        }

    def list_folder(self, path: str) -> list[dict]:
        """
        List contents of a Dropbox folder.

        Args:
            path: Dropbox path (e.g., "/Mouser/2026-01-19_PR_Folder")

        Returns:
            List of file metadata dicts
        """
        # Ensure path starts with /
        if not path.startswith('/'):
            path = '/' + path

        response = requests.post(
            f"{self.base_url}/files/list_folder",
            headers=self._headers(),
            json={"path": path}
        )
        response.raise_for_status()

        data = response.json()
        return data.get('entries', [])

    def download_file(self, dropbox_path: str, local_path: str) -> str:
        """
        Download a single file from Dropbox.

        Args:
            dropbox_path: Full Dropbox path to file
            local_path: Local path to save file

        Returns:
            Local file path
        """
        import json

        headers = {
            "Authorization": f"Bearer {self.access_token}",
            "Dropbox-API-Arg": json.dumps({"path": dropbox_path})
        }

        response = requests.post(
            f"{self.content_url}/files/download",
            headers=headers,
            stream=True
        )
        response.raise_for_status()

        with open(local_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)

        return local_path

    def download_folder(self, dropbox_path: str, local_dir: Optional[str] = None) -> str:
        """
        Download all files from a Dropbox folder.

        Args:
            dropbox_path: Dropbox folder path (e.g., "Mouser/2026-01-19_PR")
            local_dir: Local directory to save files (creates temp if None)

        Returns:
            Path to local directory containing downloaded files
        """
        # Normalize path
        if not dropbox_path.startswith('/'):
            dropbox_path = '/' + dropbox_path

        # Create local directory
        if local_dir is None:
            local_dir = tempfile.mkdtemp(prefix="mouser_pr_")
        else:
            os.makedirs(local_dir, exist_ok=True)

        # List and download files
        entries = self.list_folder(dropbox_path)

        downloaded_files = []
        for entry in entries:
            if entry.get('.tag') == 'file':
                filename = entry['name']
                dropbox_file_path = entry['path_display']
                local_file_path = os.path.join(local_dir, filename)

                self.download_file(dropbox_file_path, local_file_path)
                downloaded_files.append(filename)

        return local_dir

    def find_folder_by_name(self, folder_name: str, search_path: str = "/Mouser") -> Optional[str]:
        """
        Search for a folder by partial name match.

        Args:
            folder_name: Folder name to search for
            search_path: Root path to search in

        Returns:
            Full Dropbox path if found, None otherwise
        """
        try:
            entries = self.list_folder(search_path)

            for entry in entries:
                if entry.get('.tag') == 'folder':
                    if folder_name in entry['name']:
                        return entry['path_display']

            # Search in subfolders (month folders)
            for entry in entries:
                if entry.get('.tag') == 'folder':
                    try:
                        sub_entries = self.list_folder(entry['path_display'])
                        for sub_entry in sub_entries:
                            if sub_entry.get('.tag') == 'folder':
                                if folder_name in sub_entry['name']:
                                    return sub_entry['path_display']
                    except Exception:
                        continue

        except Exception as e:
            print(f"Error searching for folder: {e}")

        return None
