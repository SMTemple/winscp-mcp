"""WinSCP MCP server for FTP/SFTP site management.

Provides tools to list, browse, download, and upload sites from WinSCP saved sessions.
Uses WinSCP CLI scripting under the hood.

Configuration via config.json (next to this file) or environment variables:
  WINSCP_PATH     - path to winscp.com
  WINSCP_INI_PATH - path to WinSCP .ini config
  WINSCP_DOWNLOAD_ROOT - local directory for downloads
"""

import json
import os
import re
import subprocess
import tempfile
import threading
import time
from pathlib import Path
from typing import Optional
from urllib.parse import unquote

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("winscp-ftp")

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

_CONFIG_PATH = Path(__file__).parent / "config.json"
_config: dict = {}
if _CONFIG_PATH.exists():
    with open(_CONFIG_PATH) as f:
        _config = json.load(f)

WINSCP_PATH = os.environ.get("WINSCP_PATH", _config.get("winscp_path", r"D:\WinSCP\winscp.com"))
INI_PATH = os.environ.get("WINSCP_INI_PATH", _config.get("ini_path", ""))
DOWNLOAD_ROOT = os.environ.get("WINSCP_DOWNLOAD_ROOT", _config.get("download_root", r"D:\Local Sites"))
DEFAULT_REMOTE = _config.get("default_remote_path", "/public_html")
IGNORE_PATTERNS: list[str] = _config.get("ignore_patterns", [
    "error_log*", "*.bak", "*.log", "*.tar", "*.tar.gz", "*.tgz", "*.zip",
    "*.gz", "node_modules/", "vendor/", "vendors/", "cgi-bin/", ".ftpquota",
    "*.swp", "*.tmp", ".git/", "cache/", "backups/",
])

# Active downloads tracked in-memory
_downloads: dict[str, dict] = {}
_download_lock = threading.Lock()

# ---------------------------------------------------------------------------
# INI parsing helpers
# ---------------------------------------------------------------------------

def _parse_ini() -> dict[str, dict]:
    """Parse WinSCP INI file and return dict of session_name -> session_info."""
    if not INI_PATH or not os.path.exists(INI_PATH):
        return {}

    sessions = {}
    current_section = None
    current_data = {}

    with open(INI_PATH, encoding="utf-8", errors="replace") as f:
        for line in f:
            line = line.rstrip("\n\r")
            # Section header
            m = re.match(r"^\[Sessions\\(.+)\]$", line)
            if m:
                # Save previous section
                if current_section:
                    sessions[current_section] = current_data
                current_section = m.group(1)
                current_data = {}
                continue
            # Key=Value inside a session
            if current_section and "=" in line:
                key, _, val = line.partition("=")
                current_data[key.strip()] = val.strip()

        # Save last section
        if current_section:
            sessions[current_section] = current_data

    return sessions


def _decode_name(raw: str) -> str:
    """Decode URL-encoded session name for display."""
    return unquote(raw.replace("%20", " "))


def _build_filemask(extra_excludes: Optional[list[str]] = None) -> str:
    """Build WinSCP filemask string from ignore patterns."""
    excludes = list(IGNORE_PATTERNS)
    if extra_excludes:
        excludes.extend(extra_excludes)
    if not excludes:
        return ""
    return "|" + ";".join(excludes)


# ---------------------------------------------------------------------------
# WinSCP script execution
# ---------------------------------------------------------------------------

def _run_script(script_lines: list[str], timeout: int = 120) -> tuple[int, str]:
    """Write a temp script file and run WinSCP with it. Returns (exit_code, output)."""
    with tempfile.NamedTemporaryFile(
        mode="w", suffix=".txt", prefix="winscp_", delete=False, dir=tempfile.gettempdir()
    ) as f:
        f.write("\n".join(script_lines) + "\n")
        script_path = f.name

    try:
        cmd = f'"{WINSCP_PATH}" /ini="{INI_PATH}" /script="{script_path}"'
        result = subprocess.run(
            ["cmd.exe", "/c", cmd],
            capture_output=True, text=True, timeout=timeout,
            encoding="utf-8", errors="replace",
        )
        return result.returncode, result.stdout + result.stderr
    except subprocess.TimeoutExpired:
        return -1, "Command timed out"
    finally:
        try:
            os.unlink(script_path)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# Tools
# ---------------------------------------------------------------------------

@mcp.tool()
def list_sites(search: str = "") -> str:
    """List available FTP/SFTP sites from WinSCP saved sessions.

    Args:
        search: Optional search term to filter sites (case-insensitive).
                 Leave empty to list all sites.
    """
    sessions = _parse_ini()
    if not sessions:
        return "No sessions found. Check that ini_path is configured correctly in config.json."

    results = []
    for raw_name, data in sorted(sessions.items()):
        display = _decode_name(raw_name)
        if search and search.lower() not in display.lower():
            continue
        host = data.get("HostName", "?")
        user = data.get("UserName", "?")
        protocol = "SFTP" if data.get("FSProtocol") == "7" else "FTP"
        ftps = data.get("Ftps", "0")
        if ftps in ("1", "2", "3"):
            protocol = "FTPS"
        results.append(f"  {display}  |  {protocol}  |  {user}@{host}")

    if not results:
        return f"No sites matching '{search}'." if search else "No sites found."

    header = f"Found {len(results)} site(s):\n"
    return header + "\n".join(results)


@mcp.tool()
def get_site_info(site_name: str) -> str:
    """Get connection details for a specific WinSCP site.

    Args:
        site_name: The site name (or partial match). Examples: 'BarberInstitute', 'Erie Zoo'
    """
    sessions = _parse_ini()
    matches = []
    for raw_name, data in sessions.items():
        display = _decode_name(raw_name)
        if site_name.lower() in display.lower():
            matches.append((raw_name, display, data))

    if not matches:
        return f"No site found matching '{site_name}'."

    lines = []
    for raw, display, data in matches:
        host = data.get("HostName", "?")
        port = data.get("PortNumber", "21")
        user = data.get("UserName", "?")
        remote = data.get("RemoteDirectory", "/")
        fs = data.get("FSProtocol", "?")
        passive = "No (Active)" if data.get("FtpPasvMode", "1") == "0" else "Yes (Passive)"
        ftps = data.get("Ftps", "0")

        protocol = "FTP"
        if fs == "7":
            protocol = "SFTP"
        elif ftps in ("1", "2", "3"):
            protocol = "FTPS"

        lines.append(
            f"Session: {display}\n"
            f"  Session key: {raw}\n"
            f"  Host: {host}:{port}\n"
            f"  User: {user}\n"
            f"  Protocol: {protocol}\n"
            f"  Passive mode: {passive}\n"
            f"  Saved remote dir: {remote}"
        )

    return "\n\n".join(lines)


@mcp.tool()
def list_remote_files(site_name: str, remote_path: str = "") -> str:
    """List files and directories on a remote FTP/SFTP server.

    Args:
        site_name: The site name (or partial match)
        remote_path: Remote directory to list (default: /public_html)
    """
    sessions = _parse_ini()
    session_key = _find_session(sessions, site_name)
    if not session_key:
        return f"No site found matching '{site_name}'."

    if not remote_path:
        remote_path = DEFAULT_REMOTE

    script = [
        "option batch abort",
        "option confirm off",
        f"open \"{session_key}\"",
        f"ls \"{remote_path}/\"",
        "exit",
    ]
    code, output = _run_script(script, timeout=30)

    # Parse the ls output — extract file listing lines
    lines = output.strip().split("\n")
    file_lines = []
    for line in lines:
        # WinSCP ls output has permissions at start
        if re.match(r"^[Dd-][rwx-]", line.strip()):
            file_lines.append(line.strip())

    if not file_lines:
        return f"Could not list {remote_path}. WinSCP output:\n{output}"

    return f"Contents of {remote_path} ({len(file_lines)} items):\n" + "\n".join(file_lines)


@mcp.tool()
def download_site(
    site_name: str,
    remote_path: str = "",
    local_folder: str = "",
    exclude: str = "",
) -> str:
    """Download an entire site (or specific remote path) to the local machine.

    Runs asynchronously — returns a download ID to check status with download_status().
    Automatically skips files matching the global ignore list.

    Args:
        site_name: The site name (or partial match). Example: 'BarberInstitute'
        remote_path: Remote directory to download (default: /public_html)
        local_folder: Local folder name inside the download root (default: derived from site name)
        exclude: Extra comma-separated exclude patterns (e.g. '*.zip,backups/')
                 These are added ON TOP of the global ignore list.
    """
    sessions = _parse_ini()
    session_key = _find_session(sessions, site_name)
    if not session_key:
        return f"No site found matching '{site_name}'."

    if not remote_path:
        remote_path = DEFAULT_REMOTE

    # Determine local path
    if not local_folder:
        # Use the last part of the session key, decoded
        local_folder = _decode_name(session_key.split("/")[-1])

    local_path = os.path.join(DOWNLOAD_ROOT, local_folder)
    os.makedirs(local_path, exist_ok=True)

    # Build filemask
    extra = [p.strip() for p in exclude.split(",") if p.strip()] if exclude else None
    filemask = _build_filemask(extra)
    filemask_opt = f' -filemask="{filemask}"' if filemask else ""

    script = [
        "option batch continue",
        "option confirm off",
        f"open \"{session_key}\"",
        "option transfer binary",
        f'synchronize local{filemask_opt} "{local_path}" "{remote_path}"',
        "exit",
    ]

    # Write persistent script file (not temp — so we can resume/debug)
    script_path = os.path.join(local_path, "_winscp_download.txt")
    with open(script_path, "w") as f:
        f.write("\n".join(script) + "\n")

    # Generate download ID
    dl_id = f"dl_{int(time.time())}_{local_folder.replace(' ', '_')[:20]}"

    # Launch in background thread
    def _run():
        with _download_lock:
            _downloads[dl_id]["status"] = "running"
            _downloads[dl_id]["started"] = time.time()

        try:
            cmd = f'"{WINSCP_PATH}" /ini="{INI_PATH}" /script="{script_path}"'
            result = subprocess.run(
                ["cmd.exe", "/c", cmd],
                capture_output=True, text=True, timeout=3600,
                encoding="utf-8", errors="replace",
            )
            with _download_lock:
                _downloads[dl_id]["status"] = "completed" if result.returncode == 0 else "completed_with_errors"
                _downloads[dl_id]["exit_code"] = result.returncode
                _downloads[dl_id]["output_tail"] = result.stdout[-2000:] if result.stdout else ""
                _downloads[dl_id]["ended"] = time.time()
        except subprocess.TimeoutExpired:
            with _download_lock:
                _downloads[dl_id]["status"] = "timed_out"
                _downloads[dl_id]["ended"] = time.time()
        except Exception as e:
            with _download_lock:
                _downloads[dl_id]["status"] = "error"
                _downloads[dl_id]["error"] = str(e)
                _downloads[dl_id]["ended"] = time.time()

    with _download_lock:
        _downloads[dl_id] = {
            "site": site_name,
            "session_key": session_key,
            "remote_path": remote_path,
            "local_path": local_path,
            "filemask": filemask,
            "status": "starting",
            "script_path": script_path,
        }

    t = threading.Thread(target=_run, daemon=True)
    t.start()

    ignore_display = "\n".join(f"  - {p}" for p in IGNORE_PATTERNS)
    extra_display = ""
    if extra:
        extra_display = "\nExtra excludes:\n" + "\n".join(f"  - {p}" for p in extra)

    return (
        f"Download started!\n"
        f"  ID: {dl_id}\n"
        f"  Site: {_decode_name(session_key)}\n"
        f"  Remote: {remote_path}\n"
        f"  Local: {local_path}\n"
        f"  Filemask: {filemask}\n\n"
        f"Global ignore list:\n{ignore_display}{extra_display}\n\n"
        f"Use download_status('{dl_id}') to check progress."
    )


@mcp.tool()
def download_status(download_id: str = "") -> str:
    """Check the status of a running or completed download.

    Args:
        download_id: The download ID returned by download_site().
                     Leave empty to see all downloads.
    """
    with _download_lock:
        if not _downloads:
            return "No downloads tracked in this session."

        if download_id and download_id in _downloads:
            dl = _downloads[download_id]
            elapsed = ""
            if "started" in dl:
                end = dl.get("ended", time.time())
                mins = (end - dl["started"]) / 60
                elapsed = f"\n  Elapsed: {mins:.1f} minutes"

            result = (
                f"Download: {download_id}\n"
                f"  Status: {dl['status']}\n"
                f"  Site: {dl.get('site', '?')}\n"
                f"  Local: {dl.get('local_path', '?')}\n"
                f"  Remote: {dl.get('remote_path', '?')}"
                f"{elapsed}"
            )
            if dl.get("output_tail"):
                # Show last few lines
                tail = dl["output_tail"].strip().split("\n")[-10:]
                result += "\n  Last output:\n    " + "\n    ".join(tail)
            if dl.get("error"):
                result += f"\n  Error: {dl['error']}"
            return result

        if download_id:
            return f"No download found with ID '{download_id}'."

        # List all
        lines = [f"Tracked downloads ({len(_downloads)}):"]
        for did, dl in _downloads.items():
            lines.append(f"  {did}: {dl['status']} — {dl.get('site', '?')} -> {dl.get('local_path', '?')}")
        return "\n".join(lines)


@mcp.tool()
def upload_file(
    site_name: str,
    local_path: str,
    remote_path: str = "",
) -> str:
    """Upload one or more local files to a remote FTP/SFTP server.

    Args:
        site_name: The site name (or partial match). Example: 'eccleston law'
        local_path: Local file or files to upload. Comma-separated for multiple files.
                    Example: 'C:/Users/me/file.php' or 'C:/path/a.php,C:/path/b.php'
        remote_path: Remote directory to upload into (default: /public_html).
                     Must be a directory path, not a file path.
    """
    sessions = _parse_ini()
    session_key = _find_session(sessions, site_name)
    if not session_key:
        return f"No site found matching '{site_name}'."

    if not remote_path:
        remote_path = DEFAULT_REMOTE

    # Ensure remote_path ends with /
    if not remote_path.endswith("/"):
        remote_path += "/"

    # Parse file list
    files = [f.strip() for f in local_path.split(",") if f.strip()]
    missing = [f for f in files if not os.path.exists(f)]
    if missing:
        return f"Local file(s) not found:\n" + "\n".join(f"  - {f}" for f in missing)

    # Build put commands
    put_cmds = []
    for f in files:
        # Normalize to forward slashes for WinSCP
        f_norm = f.replace("\\", "/")
        put_cmds.append(f'put "{f_norm}" "{remote_path}"')

    script = [
        "option batch abort",
        "option confirm off",
        f"open \"{session_key}\"",
        "option transfer binary",
    ] + put_cmds + ["exit"]

    code, output = _run_script(script, timeout=120)

    if code == 0:
        file_names = [os.path.basename(f) for f in files]
        return (
            f"Upload successful!\n"
            f"  Site: {_decode_name(session_key)}\n"
            f"  Files: {', '.join(file_names)}\n"
            f"  Remote: {remote_path}"
        )
    else:
        return f"Upload failed (exit code {code}).\nWinSCP output:\n{output}"


@mcp.tool()
def upload_directory(
    site_name: str,
    local_path: str,
    remote_path: str = "",
    exclude: str = "",
    preview: bool = True,
) -> str:
    """Synchronize a local directory to a remote FTP/SFTP server (upload changed files).

    Uses WinSCP's synchronize command to only upload files that have changed.

    Args:
        site_name: The site name (or partial match). Example: 'eccleston law'
        local_path: Local directory to upload from. Example: 'D:/Local Sites/Eccleston'
        remote_path: Remote directory to sync to (default: /public_html)
        exclude: Extra comma-separated exclude patterns (e.g. '*.zip,backups/')
                 Added on top of the global ignore list.
        preview: If True (default), only preview changes without uploading.
                 Set to False to actually upload.
    """
    sessions = _parse_ini()
    session_key = _find_session(sessions, site_name)
    if not session_key:
        return f"No site found matching '{site_name}'."

    if not remote_path:
        remote_path = DEFAULT_REMOTE

    if not os.path.isdir(local_path):
        return f"Local directory not found: {local_path}"

    # Normalize path
    local_path = local_path.replace("\\", "/")

    # Build filemask
    extra = [p.strip() for p in exclude.split(",") if p.strip()] if exclude else None
    filemask = _build_filemask(extra)
    filemask_opt = f' -filemask="{filemask}"' if filemask else ""

    preview_flag = " -preview" if preview else ""

    script = [
        "option batch continue",
        "option confirm off",
        f"open \"{session_key}\"",
        "option transfer binary",
        f'synchronize remote{preview_flag}{filemask_opt} "{local_path}" "{remote_path}"',
        "exit",
    ]

    timeout = 60 if preview else 600
    code, output = _run_script(script, timeout=timeout)

    if preview:
        return (
            f"Upload preview (no changes made):\n"
            f"  Site: {_decode_name(session_key)}\n"
            f"  Local: {local_path}\n"
            f"  Remote: {remote_path}\n\n"
            f"WinSCP output:\n{output}\n\n"
            f"To actually upload, call upload_directory with preview=False."
        )
    else:
        status = "completed" if code == 0 else f"completed with errors (exit code {code})"
        return (
            f"Upload {status}.\n"
            f"  Site: {_decode_name(session_key)}\n"
            f"  Local: {local_path}\n"
            f"  Remote: {remote_path}\n\n"
            f"WinSCP output:\n{output}"
        )


@mcp.tool()
def get_ignore_list() -> str:
    """Show the current global ignore list for downloads."""
    if not IGNORE_PATTERNS:
        return "Ignore list is empty. All files will be downloaded."
    lines = ["Global ignore patterns (from config.json):"]
    for p in IGNORE_PATTERNS:
        lines.append(f"  - {p}")
    lines.append(f"\nConfig file: {_CONFIG_PATH}")
    return "\n".join(lines)


@mcp.tool()
def update_ignore_list(add: str = "", remove: str = "") -> str:
    """Add or remove patterns from the global ignore list.

    Changes are saved to config.json and persist across sessions.

    Args:
        add: Comma-separated patterns to add (e.g. '*.zip,backups/,cache/')
        remove: Comma-separated patterns to remove (e.g. '*.bak,*.log')
    """
    global IGNORE_PATTERNS

    added = []
    removed = []

    if add:
        for p in add.split(","):
            p = p.strip()
            if p and p not in IGNORE_PATTERNS:
                IGNORE_PATTERNS.append(p)
                added.append(p)

    if remove:
        for p in remove.split(","):
            p = p.strip()
            if p in IGNORE_PATTERNS:
                IGNORE_PATTERNS.remove(p)
                removed.append(p)

    # Save back to config.json
    _config["ignore_patterns"] = IGNORE_PATTERNS
    with open(_CONFIG_PATH, "w") as f:
        json.dump(_config, f, indent=2)

    lines = []
    if added:
        lines.append("Added: " + ", ".join(added))
    if removed:
        lines.append("Removed: " + ", ".join(removed))
    if not lines:
        lines.append("No changes made.")

    lines.append("\nCurrent ignore list:")
    for p in IGNORE_PATTERNS:
        lines.append(f"  - {p}")

    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_session(sessions: dict, search: str) -> Optional[str]:
    """Find the best matching session key for a search term."""
    search_lower = search.lower()

    # Exact match on last segment
    for raw_name in sessions:
        last_part = _decode_name(raw_name.split("/")[-1])
        if last_part.lower() == search_lower:
            return raw_name

    # Partial match on last segment
    for raw_name in sessions:
        last_part = _decode_name(raw_name.split("/")[-1])
        if search_lower in last_part.lower():
            return raw_name

    # Partial match on full path
    for raw_name in sessions:
        display = _decode_name(raw_name)
        if search_lower in display.lower():
            return raw_name

    return None


if __name__ == "__main__":
    mcp.run()
