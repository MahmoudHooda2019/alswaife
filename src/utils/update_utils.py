"""
Update Utilities - Check and download updates from GitHub
"""

import os
import re
import requests
import subprocess
import tempfile
from typing import Optional, Tuple

from utils.log_utils import log_error

# GitHub repository info
GITHUB_REPO = "MahmoudHooda2019/alswaife"
GITHUB_RAW_URL = f"https://raw.githubusercontent.com/{GITHUB_REPO}/main/src/version.py"
GITHUB_DOWNLOAD_URL = "https://github.com/MahmoudHooda2019/alswaife/raw/refs/heads/main/AlSawifeFactory-setup.exe"
SETUP_FILENAME = "AlSawifeFactory-setup.exe"


def get_current_version() -> str:
    """Get current installed version"""
    try:
        from version import __version__

        return __version__
    except ImportError:
        return "1.0"


def get_latest_version() -> Tuple[Optional[str], Optional[str]]:
    """
    Check GitHub for latest version by reading src/version.py file.
    Returns: (version, download_url) or (None, None) if failed
    """
    try:
        response = requests.get(GITHUB_RAW_URL, timeout=10)
        response.raise_for_status()

        content = response.text

        # Parse __version__ from the file content
        # Looking for: __version__ = "1.0" or __version__ = '1.0'
        match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', content)

        if match:
            version = match.group(1)
            return version, GITHUB_DOWNLOAD_URL

        log_error("Could not find __version__ in version.py")
        return None, None

    except requests.RequestException as e:
        log_error(f"Failed to check for updates: {e}")
        return None, None
    except Exception as e:
        log_error(f"Unexpected error checking updates: {e}")
        return None, None


def compare_versions(current: str, latest: str) -> bool:
    """
    Compare version strings.
    Returns True if latest > current (update available)
    """
    try:
        def parse_version(v):
            # Remove 'v' prefix if present
            v = v.lstrip("v")
            # Split by dots and convert to integers
            parts = []
            for part in v.split("."):
                # Handle versions like "1.0.1-beta"
                num_part = ""
                for char in part:
                    if char.isdigit():
                        num_part += char
                    else:
                        break
                parts.append(int(num_part) if num_part else 0)
            return parts
        
        current_parts = parse_version(current)
        latest_parts = parse_version(latest)
        
        # Pad shorter version with zeros
        max_len = max(len(current_parts), len(latest_parts))
        current_parts.extend([0] * (max_len - len(current_parts)))
        latest_parts.extend([0] * (max_len - len(latest_parts)))
        
        return latest_parts > current_parts
    
    except Exception as e:
        log_error(f"Version comparison failed: {e}")
        return False


def check_for_updates() -> Tuple[bool, str, str, Optional[str]]:
    """
    Check if updates are available.
    Returns: (update_available, current_version, latest_version, download_url)
    """
    current = get_current_version()
    latest, download_url = get_latest_version()
    
    if latest is None:
        return False, current, "غير متاح", None
    
    update_available = compare_versions(current, latest)
    return update_available, current, latest, download_url


def download_update(download_url: str, progress_callback=None, cancel_check=None) -> Optional[str]:
    """
    Download the update file.
    Args:
        download_url: URL to download from
        progress_callback: Function to call with progress percentage
        cancel_check: Function that returns True if download should be cancelled
    Returns: path to downloaded file or None if failed/cancelled
    """
    try:
        # Create temp directory for download
        temp_dir = tempfile.gettempdir()
        download_path = os.path.join(temp_dir, SETUP_FILENAME)
        
        # Download with progress
        response = requests.get(download_url, stream=True, timeout=60)
        response.raise_for_status()
        
        total_size = int(response.headers.get('content-length', 0))
        downloaded = 0
        
        with open(download_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                # Check if cancelled
                if cancel_check and cancel_check():
                    f.close()
                    # Delete partial file
                    if os.path.exists(download_path):
                        os.remove(download_path)
                    return None
                
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_callback and total_size > 0:
                        progress = (downloaded / total_size) * 100
                        progress_callback(progress)
        
        return download_path
    
    except requests.RequestException as e:
        log_error(f"Failed to download update: {e}")
        return None
    except Exception as e:
        log_error(f"Unexpected error downloading update: {e}")
        return None


def install_update(setup_path: str) -> bool:
    """
    Run the setup installer.
    Returns: True if installer started successfully
    """
    try:
        if not os.path.exists(setup_path):
            log_error(f"Setup file not found: {setup_path}")
            return False
        
        # Run the installer
        subprocess.Popen([setup_path], shell=True)
        return True
    
    except Exception as e:
        log_error(f"Failed to run installer: {e}")
        return False
