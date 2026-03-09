"""File loader with caching support for remote URLs."""

import tempfile
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Optional


class FileLoader:
    """
    Handles file loading from local paths or remote URLs with caching.

    Cache Path Generation:
    URLs are cached using a hierarchical directory structure that mirrors
    the URL structure, making cache files readable and debuggable. For example:

    URL: "http://disclosure.edinet-fsa.go.jp/taxonomy/jppfs/2022/jppfs_cor.xsd"
    Cache path: "xbrlp/cache/raw/disclosure.edinet-fsa.go.jp/taxonomy/jppfs/2022/jppfs_cor.xsd"

    URL: "http://example.com:8080/path/to/file.xml?param=value"
    Cache path: "xbrlp/cache/raw/example.com_8080/path/to/file.xml_param_value"

    This approach:
    - Makes cache files easily identifiable and debuggable
    - Preserves URL structure for easy navigation
    - Handles special characters safely (? → _, : → _, etc.)
    - Ensures the same URL always maps to the same cache path across processes
    - Maintains file extensions for proper file type identification
    """

    def __init__(
        self, cache_dir: Optional[Path] = None, *, ignore_failure: bool = False
    ):
        """
        Initialize the FileLoader.

        Args:
            cache_dir: Directory for caching downloaded files.
                      Defaults to xbrlp/.cache/raw/ relative to this file
            ignore_failure: When True, return None from fetch() if any error occurs
                             instead of raising the exception.
        """
        if cache_dir:
            self.cache_dir = Path(cache_dir)
        else:
            # Get the directory containing this file
            file_dir = Path(__file__).parent
            self.cache_dir = file_dir / ".cache" / "raw"
        self.cache_dir.mkdir(parents=True, exist_ok=True)
        self.ignore_failure = ignore_failure

    def fetch(self, path_or_url: str) -> Optional[Path]:
        """
        Fetch a file from a local path or remote URL.

        Args:
            path_or_url: Local file path or HTTP/HTTPS URL

        Returns:
            Path object pointing to the local file, or None if ignore_failure is
            enabled and an error occurs.

        Raises:
            FileNotFoundError: If local file doesn't exist
            urllib.error.URLError: If URL download fails
        """
        try:
            # Check if it's a URL
            if path_or_url.startswith(("http://", "https://")):
                return self._fetch_url(path_or_url)

            # Local file path
            path = Path(path_or_url)
            if not path.exists():
                raise FileNotFoundError(f"File not found: {path}")
            return path
        except Exception:
            if self.ignore_failure:
                return None
            raise

    def _fetch_url(self, url: str) -> Path:
        """
        Fetch a file from a URL with caching.

        Args:
            url: HTTP/HTTPS URL to download

        Returns:
            Path to the cached file
        """
        # Generate cache filename based on URL
        cache_path = self._get_cache_path(url)

        # Check if file exists in cache
        if cache_path.exists():
            return cache_path

        # Download file to cache
        self._download_file(url, cache_path)
        return cache_path

    def _get_cache_path(self, url: str) -> Path:
        """
        Generate a cache file path for a URL.

        Creates a hierarchical cache structure that mirrors the URL structure
        for easy debugging and identification. Special characters are replaced
        with underscores to ensure filesystem compatibility.

        Args:
            url: URL to generate cache path for

        Returns:
            Path object for the cache file
        """
        parsed = urllib.parse.urlparse(url)

        # Build host directory (handle port if present)
        if parsed.port:
            host_dir = f"{parsed.hostname}_{parsed.port}"
        else:
            host_dir = parsed.netloc

        # Get the path component (remove leading /)
        path_component = parsed.path.lstrip("/")

        # Handle query parameters if present
        if parsed.query:
            # Replace special chars in query string with underscores
            safe_query = (
                parsed.query.replace("&", "_").replace("=", "_").replace("/", "_")
            )
            # Append query to filename
            if path_component:
                # Add query to the filename part
                if "." in path_component:
                    # Insert before extension
                    base, ext = path_component.rsplit(".", 1)
                    path_component = f"{base}_{safe_query}.{ext}"
                else:
                    path_component = f"{path_component}_{safe_query}"
            else:
                path_component = f"index_{safe_query}"
        elif not path_component:
            # No path and no query, use index as filename
            path_component = "index"

        # Build the full cache path
        cache_path = self.cache_dir / host_dir / path_component

        return cache_path

    def _download_file(self, url: str, dest_path: Path) -> None:
        """
        Download a file from URL to destination path using atomic write.

        Uses a temporary file and atomic rename to prevent race conditions
        when multiple processes try to download the same URL simultaneously.

        Args:
            url: URL to download from
            dest_path: Path to save the file

        Raises:
            urllib.error.URLError: If download fails
        """
        # Ensure parent directory exists
        dest_path.parent.mkdir(parents=True, exist_ok=True)

        # Download to a temporary file first
        # Use same directory as dest_path to ensure atomic rename works
        with tempfile.NamedTemporaryFile(
            mode="wb",
            dir=dest_path.parent,
            prefix=".download_",
            suffix=dest_path.suffix,
            delete=False,
        ) as tmp_file:
            tmp_path = Path(tmp_file.name)
            try:
                # Download with urllib
                with urllib.request.urlopen(url) as response:
                    content = response.read()
                    tmp_file.write(content)

                # Atomic rename - if file exists, it will be replaced atomically
                # This prevents partial files and handles concurrent downloads
                tmp_path.replace(dest_path)
            except Exception:
                # Clean up temp file on error
                tmp_path.unlink(missing_ok=True)
                raise

    def clear_cache(self) -> None:
        """Clear all cached files."""
        if self.cache_dir.exists():
            for cache_file in self.cache_dir.iterdir():
                if cache_file.is_file():
                    cache_file.unlink()